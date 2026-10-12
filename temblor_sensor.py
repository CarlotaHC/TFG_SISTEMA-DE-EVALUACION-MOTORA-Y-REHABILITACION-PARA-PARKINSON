# -*- coding: utf-8 -*-
"""
temblor_sensor.py
=================
Módulo reutilizable para caracterizar el temblor mediante el acelerómetro de
una placa micro:bit conectada por USB (puerto serie).

Idea general
------------
1. El firmware de la micro:bit (ver ``microbit/acelerometro_makecode.py``) envía
   por serie una línea CSV por muestra:  ``t_ms,x,y,z``  a ~20 Hz.
2. ``TemblorSensor.iniciar()`` abre el puerto y lanza un HILO en segundo plano
   que lee esas líneas. Así el bucle del minijuego (cámara + dibujo) NO se
   bloquea esperando al puerto serie.
3. Al terminar la partida se llama a ``detener()`` y después se puede:
     - ``calcular_metricas()``  -> frecuencia dominante y amplitud por eje,
                                   sin escribir ningún archivo.
     - ``guardar_grafica()``    -> PNG con el espectro FFT de los 3 ejes.
     - ``guardar_csv()``        -> datos crudos (opcional).

Por qué se usa el reloj de la micro:bit y no el del PC
------------------------------------------------------
El buffer USB-serie introduce "jitter" en el instante en que el PC recibe cada
línea, por lo que el muestreo visto desde el PC NO es uniforme. Para calcular la
FFT hace falta una frecuencia de muestreo (fs) fiable, así que se usa la marca
de tiempo que la propia micro:bit incluye en cada línea (``t_firmware``) y se
estima ``fs = (N - 1) / duración``.

Cadena de procesado de cada eje
-------------------------------
    señal cruda -> quitar media (DC) -> filtro paso-banda Butterworth
                -> ventana de Hann -> FFT (rfft) -> normalizar por N
                -> detección de picos -> frecuencia dominante

Uso mínimo
----------
    sensor = TemblorSensor(puerto="COM3")
    sensor.iniciar()                 # bloquea ~2 s (arranque de la placa)
    ...                              # el juego corre mientras el hilo captura
    sensor.detener()
    metricas = sensor.calcular_metricas()
    sensor.guardar_grafica("P001")

Si la micro:bit no está conectada el módulo NO lanza excepciones: registra un
aviso y el juego continúa sin datos de temblor (``conectado = False``).
"""

import logging
import serial          # pyserial: comunicación con la micro:bit
import threading       # hilo de lectura en segundo plano
import time
import csv
from datetime import datetime

import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import butter, filtfilt, find_peaks

logger = logging.getLogger(__name__)

# Nº mínimo de muestras para que el filtrado y la FFT sean fiables.
# filtfilt con un filtro de orden N necesita > 3*2*N muestras; con orden 2 son
# > 12. Se usan 30 como margen de seguridad (a 20 Hz son ~1,5 s de captura).
_MIN_MUESTRAS = 30


class TemblorSensor:
    """Captura en segundo plano del acelerómetro de la micro:bit y análisis FFT."""

    def __init__(self, puerto="COM3", baudios=115200):
        """
        Parameters
        ----------
        puerto : str
            Puerto serie de la micro:bit ("COM3" en Windows, "/dev/ttyACM0"
            en Linux, "/dev/cu.usbmodem..." en macOS).
        baudios : int
            Velocidad del puerto. Debe coincidir con la de la micro:bit
            (115200 por defecto en MakeCode).
        """
        self.puerto  = puerto
        self.baudios = baudios

        self.ser       = None    # objeto serial.Serial (None si no hay conexión)
        self._hilo      = None   # hilo de lectura
        self._grabando  = False  # bandera que mantiene vivo el hilo
        self.conectado  = False  # True solo si el puerto se abrió correctamente

        # Datos crudos (una entrada por muestra recibida)
        self.t_firmware = []   # ms, reloj de la propia micro:bit
        self.datos_x    = []   # aceleración eje X (mili-g)
        self.datos_y    = []   # aceleración eje Y (mili-g)
        self.datos_z    = []   # aceleración eje Z (mili-g)

    # ================= CAPTURA =================
    def iniciar(self):
        """Abre el puerto serie y arranca el hilo de lectura. No bloquea
        después del arranque (solo espera ~2 s a que la placa se inicie)."""
        try:
            self.ser = serial.Serial(self.puerto, self.baudios, timeout=0.1)
            time.sleep(2)  # tiempo de arranque del micro:bit
            self.conectado = True
        except Exception as e:
            # Sin sensor el juego sigue funcionando; solo se pierde el temblor.
            logger.warning(
                "AVISO temblor: no se pudo abrir %s (%s). "
                "El juego continua sin registrar temblor.", self.puerto, e
            )
            self.ser = None
            self.conectado = False
            return

        self._grabando = True
        # daemon=True: el hilo no impide cerrar el programa si algo falla.
        self._hilo = threading.Thread(target=self._loop_lectura, daemon=True)
        self._hilo.start()

    def _loop_lectura(self):
        """Bucle del hilo: lee líneas ``t,x,y,z`` y las añade a las listas."""
        while self._grabando:
            try:
                linea = self.ser.readline().decode().strip()
                if linea:
                    t, x, y, z = map(float, linea.split(","))
                    self.t_firmware.append(t)
                    self.datos_x.append(x)
                    self.datos_y.append(y)
                    self.datos_z.append(z)
            except ValueError:
                # Línea corrupta o incompleta (típico al arrancar): se ignora.
                pass
            except Exception as e:
                # Error inesperado (puerto cerrado, hilo zombie...): se registra.
                logger.error("Error inesperado en lectura serie: %s", e)

    def detener(self):
        """Para el hilo de lectura y cierra el puerto. Es idempotente: se puede
        llamar varias veces sin que falle."""
        self._grabando = False
        if self._hilo is not None:
            self._hilo.join(timeout=1.0)
            self._hilo = None          # evita un segundo join
        if self.ser is not None:
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None            # evita un segundo close

    # ================= GUARDADO CSV =================
    def guardar_csv(self, paciente_id):
        """Guarda los datos crudos en ``temblor_<id>_<fecha>.csv``.

        Columnas: t_firmware_ms, x, y, z.
        Devuelve el nombre del archivo, o None si no había datos.
        """
        if len(self.t_firmware) == 0:
            logger.warning(
                "AVISO temblor: no hay datos que guardar (sin conexion al micro:bit)."
            )
            return None

        fecha  = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre = f"temblor_{paciente_id}_{fecha}.csv"

        with open(nombre, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["t_firmware_ms", "x", "y", "z"])
            for t, x, y, z in zip(self.t_firmware, self.datos_x, self.datos_y, self.datos_z):
                w.writerow([t, x, y, z])

        logger.info("CSV temblor guardado: %s", nombre)
        return nombre

    # ================= ANALISIS FFT =================
    def _bandpass(self, signal, fs):
        """Filtro paso-banda Butterworth (0,5 Hz - 8 Hz) aplicado con filtfilt.

        - Se elimina la deriva lenta (< 0,5 Hz: cambios de postura, gravedad) y
          el ruido de alta frecuencia (> 8 Hz), conservando la banda donde
          aparece el temblor parkinsoniano (~4-6 Hz).
        - ``filtfilt`` filtra hacia delante y hacia atrás, por lo que NO
          introduce desfase en la señal.

        Salvaguardas (para que nunca lance excepción):
          1. Se comprueba que lowcut < highcut antes de llamar a ``butter``
             (si fs es muy baja, la frecuencia de corte superior se queda por
             debajo de la inferior).
          2. Se comprueba la longitud mínima antes de ``filtfilt``.
          3. Se usa orden 2 (en vez de 4) para mayor estabilidad y menos
             artefactos de borde en señales cortas.

        Si alguna comprobación falla, se devuelve la señal sin filtrar.
        """
        nyquist = fs / 2
        lowcut  = 0.5
        # El corte superior nunca debe acercarse a Nyquist (0,9 * Nyquist).
        highcut = min(8.0, nyquist * 0.9)

        # highcut debe ser estrictamente mayor que lowcut
        if highcut <= lowcut:
            logger.warning(
                "fs=%.2f Hz demasiado baja para el filtro paso-banda; "
                "se devuelve señal sin filtrar.", fs
            )
            return signal

        # Con orden 2, filtfilt necesita > 12 muestras; _MIN_MUESTRAS es la guarda.
        if len(signal) < _MIN_MUESTRAS:
            logger.warning(
                "Señal demasiado corta (%d muestras) para filtfilt; "
                "se devuelve sin filtrar.", len(signal)
            )
            return signal

        # Las frecuencias de corte se normalizan respecto a Nyquist (0-1).
        b, a = butter(2, [lowcut / nyquist, highcut / nyquist], btype="band")
        return filtfilt(b, a, signal)

    def _analizar_eje(self, signal, fs):
        """Analiza un eje y devuelve ``(freq_dom, amplitud_pico, concluyente, freqs, amps)``.

        - ``freq_dom``     : frecuencia (Hz) del pico dominante, o None si no hay espectro.
        - ``amplitud_pico``: amplitud FFT normalizada de ese pico.
        - ``concluyente``  : True si el pico destaca claramente sobre el ruido
                             (criterio de prominencia); False si es solo el
                             máximo del espectro sin pico claro.
        - ``freqs, amps``  : espectro completo (> 0,5 Hz) para poder graficarlo.

        Decisiones de diseño:
        - Se trabaja sobre una copia (``np.array``), nunca se modifica la lista
          original de datos crudos.
        - Se usa ``np.hanning`` para reducir la fuga espectral. La ventana
          anula los extremos de la señal; es un efecto esperado.
        - Las amplitudes se dividen por N para que el umbral de detección sea
          independiente de la duración de la captura.
        - El umbral de prominencia es relativo (3 x la amplitud media del
          espectro), sin constantes en unidades de FFT cruda.
        """
        signal = np.array(signal, dtype=float)   # copia propia
        signal -= np.mean(signal)                # quita la componente continua (gravedad/offset)
        signal = self._bandpass(signal, fs)

        # Ventana de Hann: suaviza los bordes antes de la FFT.
        signal = signal * np.hanning(len(signal))

        fft   = np.fft.rfft(signal)                       # FFT de señal real
        freqs = np.fft.rfftfreq(len(signal), d=1.0 / fs)  # eje de frecuencias (Hz)

        # Normalizar por N: la amplitud deja de depender de la longitud.
        N    = len(signal)
        amps = np.abs(fft) / N

        # Se descartan las frecuencias <= 0,5 Hz (deriva / DC residual).
        mask  = freqs > 0.5
        freqs = freqs[mask]
        amps  = amps[mask]

        if len(amps) == 0:
            # Se devuelve None de forma explícita; los llamadores lo comprueban
            # antes de dibujar o mostrar resultados.
            return None, 0.0, False, freqs, amps

        # Un pico es "concluyente" si su prominencia supera 3 veces la media.
        umbral_prominencia = np.mean(amps) * 3.0
        peaks, _ = find_peaks(amps, prominence=umbral_prominencia)

        if len(peaks) > 0:
            # De todos los picos prominentes, se elige el de mayor amplitud.
            pico = peaks[np.argmax(amps[peaks])]
            return float(freqs[pico]), float(amps[pico]), True, freqs, amps

        # Sin pico claro: se devuelve el máximo del espectro pero marcado como
        # NO concluyente (probablemente ruido).
        pico = int(np.argmax(amps))
        return float(freqs[pico]), float(amps[pico]), False, freqs, amps

    # ================= FRECUENCIA DE MUESTREO =================
    def _calcular_fs(self):
        """Estima la frecuencia de muestreo a partir del reloj del firmware.

        Devuelve ``(fs, n)``. ``fs`` es None si hay menos de ``_MIN_MUESTRAS``
        muestras o si la duración es cero/casi nula (evita dividir por cero).
        """
        n = len(self.datos_x)
        if n < _MIN_MUESTRAS:
            return None, n

        # t_firmware está en milisegundos -> segundos
        duracion_s = (self.t_firmware[-1] - self.t_firmware[0]) / 1000.0
        if duracion_s < 1e-6:
            logger.warning(
                "Duracion de captura nula o casi nula (%.6f s); "
                "no se puede estimar fs.", duracion_s
            )
            return None, n

        # N muestras definen N-1 intervalos
        return (n - 1) / duracion_s, n

    # ================= METRICAS SIN GUARDAR =================
    def calcular_metricas(self):
        """Calcula frecuencia dominante y amplitud por eje (X, Y, Z) sin generar
        ningún archivo. Útil para mostrar los datos en pantalla justo al
        terminar la partida.

        Devuelve None si no hay datos suficientes o no se puede calcular fs.
        En caso contrario devuelve un diccionario::

            {"fs": 20.0,
             "X": {"freq_hz": 5.1, "amplitud": 12.3, "concluyente": True},
             "Y": {...},
             "Z": {...}}
        """
        fs, n = self._calcular_fs()
        if fs is None:
            logger.warning(
                "AVISO temblor: datos insuficientes o duracion nula "
                "(n=%d); no se calculan metricas.", n
            )
            return None

        resultados = {"fs": fs}
        for nombre_eje, datos in [("X", self.datos_x), ("Y", self.datos_y), ("Z", self.datos_z)]:
            freq_dom, amplitud, concluyente, _, _ = self._analizar_eje(np.array(datos), fs)
            resultados[nombre_eje] = {
                "freq_hz":     freq_dom,
                "amplitud":    amplitud,
                "concluyente": concluyente,
            }
        return resultados

    # ================= GUARDADO GRAFICA =================
    def guardar_grafica(self, paciente_id):
        """Genera y guarda un PNG con el espectro FFT de los ejes X, Y y Z.

        Devuelve siempre una tupla ``(nombre_png, resultados)`` o
        ``(None, None)`` si no hay datos suficientes, de modo que el llamador
        pueda desempaquetarla sin comprobar el tipo del retorno.

        En el gráfico, el marcador del pico es verde si el resultado es
        concluyente y gris si no lo es (la leyenda añade "(ruido?)").
        """
        fs, n = self._calcular_fs()
        if fs is None:
            logger.warning(
                "AVISO temblor: datos insuficientes o duracion nula "
                "(n=%d); no se genera grafica.", n
            )
            return None, None   # tupla consistente

        fig, ax = plt.subplots(figsize=(10, 5))
        resultados = {"fs": fs}

        for nombre_eje, datos in [("X", self.datos_x), ("Y", self.datos_y), ("Z", self.datos_z)]:
            freq_dom, amplitud, concluyente, freqs, amps = self._analizar_eje(np.array(datos), fs)
            etiqueta = nombre_eje if concluyente else f"{nombre_eje} (ruido?)"
            ax.plot(freqs, amps, label=etiqueta)

            # Solo se dibuja el marcador del pico si existe frecuencia dominante.
            if freq_dom is not None:
                marker_color = "green" if concluyente else "gray"
                ax.plot(freq_dom, amplitud, marker="o", color=marker_color)

            resultados[nombre_eje] = {
                "freq_hz":     freq_dom,
                "amplitud":    amplitud,
                "concluyente": concluyente,
            }

        ax.set_xlabel("Frecuencia (Hz)")
        ax.set_ylabel("Amplitud FFT normalizada")
        ax.set_title(f"Espectro de temblor - {paciente_id}  (fs={fs:.1f} Hz)")
        ax.legend()
        ax.grid(True, alpha=0.3)

        fecha      = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_png = f"temblor_fft_{paciente_id}_{fecha}.png"
        plt.tight_layout()
        plt.savefig(nombre_png, dpi=150, bbox_inches="tight")
        plt.close(fig)   # libera memoria (importante si se generan varias gráficas)

        logger.info("Grafica temblor guardada: %s", nombre_png)
        logger.info("Frecuencias dominantes: %s", resultados)
        return nombre_png, resultados   # tupla consistente
