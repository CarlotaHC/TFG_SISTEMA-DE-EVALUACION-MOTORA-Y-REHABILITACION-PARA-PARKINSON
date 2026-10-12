# -*- coding: utf-8 -*-
"""
handgrip_flor.py
================
Minijuego de rehabilitación "Cierre de Mano" (Hand Grip) con una flor que se
va completando.

Qué evalúa
----------
El paciente abre y cierra la mano por completo un número de veces elegido
(5, 10, 15 o personalizado) dentro de GAME_DURATION segundos. Cada cierre
completo rellena un punto de la flor: primero el centro y luego los pétalos.
A partir de la cámara se obtienen:

    - Amplitud de cada cierre (apertura máxima de la mano entre cierres, cm)
    - Velocidad de cierre de cada ciclo (máxima velocidad de cierre, cm/s)
    - Intervalo entre cierres (ITI, s)
    - Tendencia (pendiente de regresión lineal) de amplitud y velocidad
    - Variabilidad de la amplitud (coeficiente de variación, %)
    - Temblor (frecuencia dominante y amplitud por eje) con la micro:bit,
      a través de ``temblor_sensor.py``

Cómo se mide la "apertura de la mano"
-------------------------------------
Apertura global = media de las distancias punta-de-dedo -> muñeca de los 5
dedos (landmarks 4, 8, 12, 16 y 20 de MediaPipe), convertida a centímetros con
la referencia anatómica muñeca-base del dedo medio (9,5 cm). Con la mano
abierta la apertura es grande; al cerrar el puño disminuye.

Cómo se detecta un cierre
-------------------------
Se considera mano cerrada cuando al menos 4 de los 5 dedos tienen la punta más
cerca de la muñeca que su propia base (``is_hand_closed``). Un cierre se cuenta
al pasar de "abierta" a "cerrada" (flanco) y solo si ha pasado el tiempo de
espera obligatorio (COOLDOWN_CIERRE_S), que evita contar varios cierres por
temblor o por movimientos muy rápidos.

Flujo del programa
------------------
    1. Ventana Tkinter: registro del paciente y número de cierres.
    2. Bucle principal (cámara + MediaPipe Hands + dibujo con OpenCV).
    3. La partida termina al completar los cierres O al agotarse el tiempo.
    4. Al terminar se detiene el sensor de temblor y se muestran los resultados.
    5. Al pulsar Q se guardan: fila en Excel, gráfica PNG y gráfica FFT.

Nota: el script se ejecuta de arriba abajo (no hay función main()), por lo que
el estado del juego se guarda en variables globales.

Requiere: pip install opencv-python mediapipe==0.10.9 numpy matplotlib openpyxl
          pyserial scipy   (y ``temblor_sensor.py`` en la misma carpeta)
"""

import cv2
import mediapipe as mp
import time
import math
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
from collections import deque
from temblor_sensor import TemblorSensor

# ================= TEMBLOR (micro:bit) =================
PUERTO_MICROBIT = "COM3"  # Cambiar por tu puerto
# El objeto se crea aquí pero la captura NO empieza hasta temblor.iniciar(),
# que se llama justo antes de arrancar el cronómetro de la partida.
temblor = TemblorSensor(puerto=PUERTO_MICROBIT)
GUARDAR_CSV_TEMBLOR = False   # True = guardar además el CSV crudo (x,y,z por muestra)

# ================= MEDIAPIPE =================
# MediaPipe Hands devuelve 21 landmarks (puntos) por mano, con coordenadas
# normalizadas (0-1) respecto al tamaño de la imagen.
mp_hands       = mp.solutions.hands
mp_draw        = mp.solutions.drawing_utils
mp_draw_styles = mp.solutions.drawing_styles

hands = mp_hands.Hands(
    max_num_hands=1,               # solo se evalúa una mano
    model_complexity=1,            # 0 = rápido, 1 = más preciso
    min_detection_confidence=0.7,  # confianza mínima para detectar la mano
    min_tracking_confidence=0.7    # confianza mínima para seguirla entre frames
)

# ================= VENTANA =================
# La ventana se divide en: zona de juego (izquierda) + panel de métricas (derecha).
WINDOW_W = 1280
WINDOW_H = 720
PANEL_W  = 400
GAME_W   = WINDOW_W - PANEL_W
WIN_NAME = "Rehabilitacion - Cierre de Mano"
cv2.namedWindow(WIN_NAME, cv2.WINDOW_NORMAL)
cv2.resizeWindow(WIN_NAME, WINDOW_W, WINDOW_H)

# ================= CAMARA =================
cap = cv2.VideoCapture(0)   # 0 = cámara por defecto
cap.set(cv2.CAP_PROP_FRAME_WIDTH,  640)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

# ================= CONSTANTE ANATOMICA =================
# Referencia para convertir píxeles a centímetros: la distancia entre la muñeca
# (landmark 0) y la base del dedo medio (landmark 9) mide ~9,5 cm en un adulto.
# La escala se recalcula en cada frame, así el resultado no depende de lo cerca
# o lejos que esté la mano de la cámara.
HAND_REF_REAL_CM = 9.5   # muneca(0) -> base dedo medio(9): ~9.5 cm

# ================= DEDOS (para deteccion de cierre) =================
# Para cada dedo: índice del landmark de la PUNTA y de su BASE (articulación
# metacarpofalángica; en el pulgar, el landmark 2).
FINGERS = [
    {"name": "Pulgar",  "tip": 4,  "base": 2},
    {"name": "Indice",  "tip": 8,  "base": 5},
    {"name": "Medio",   "tip": 12, "base": 9},
    {"name": "Anular",  "tip": 16, "base": 13},
    {"name": "Menique", "tip": 20, "base": 17},
]
N_FINGERS = len(FINGERS)

# ================= MEDIDA GLOBAL DE APERTURA DE MANO =================
# La "apertura global de la mano" es la distancia media punta-muñeca de los
# 5 dedos (pulgar incluido). Aquí van los índices de landmarks de las puntas.
APERTURA_TIPS = [4, 8, 12, 16, 20]   # indices de landmarks de las puntas


# ================= PACIENTE =================
def obtener_datos_paciente():
    """Muestra una ventana Tkinter para registrar al paciente y elegir el nivel.

    Bloquea hasta que se pulsa "Iniciar Evaluacion" con datos válidos.
    Devuelve un diccionario ``{"id", "genero", "edad", "cierres"}``. Si se cierra
    la ventana sin iniciar, devuelve un diccionario vacío (y el programa fallará
    al acceder a ``paciente["id"]``).
    """
    datos = {}
    ventana = tk.Tk()
    ventana.title("Registro de Paciente")
    ventana.geometry("420x400")
    ventana.resizable(False, False)

    tk.Label(ventana, text="ID Paciente").pack(pady=5)
    e_id = tk.Entry(ventana, width=35)
    e_id.pack()

    tk.Label(ventana, text="Genero").pack(pady=5)
    genero_var = tk.StringVar(value="Masculino")
    frame_gen = tk.Frame(ventana)
    frame_gen.pack()
    tk.Radiobutton(frame_gen, text="Masculino", variable=genero_var, value="Masculino").pack(side="left", padx=10)
    tk.Radiobutton(frame_gen, text="Femenino",  variable=genero_var, value="Femenino").pack(side="left", padx=10)
    tk.Radiobutton(frame_gen, text="Otro",      variable=genero_var, value="Otro").pack(side="left", padx=10)

    tk.Label(ventana, text="Edad").pack(pady=5)
    e_edad = tk.Entry(ventana, width=35)
    e_edad.pack()

    # Línea separadora gris
    tk.Frame(ventana, height=1, bg="#cccccc").pack(fill="x", padx=20, pady=10)
    tk.Label(ventana, text="Numero de cierres a realizar", font=("Helvetica", 10, "bold")).pack()

    # Niveles predefinidos de dificultad (None = el usuario escribe el número).
    NIVELES = {
        "Basico (5 cierres)": 5,
        "Estandar (10 cierres)": 10,
        "Intenso (15 cierres)": 15,
        "Personalizado": None
    }
    nivel_var = tk.StringVar(value="Estandar (10 cierres)")
    frame_niv = tk.Frame(ventana)
    frame_niv.pack(pady=4)
    for etiqueta in NIVELES:
        tk.Radiobutton(frame_niv, text=etiqueta, variable=nivel_var,
                       value=etiqueta).pack(anchor="w", padx=20)

    # Campo numérico: solo se usa si se elige "Personalizado".
    frame_custom = tk.Frame(ventana)
    frame_custom.pack()
    tk.Label(frame_custom, text="Cierres:").pack(side="left", padx=5)
    e_cierres = tk.Entry(frame_custom, width=6)
    e_cierres.insert(0, "10")
    e_cierres.pack(side="left")
    tk.Label(frame_custom, text="(solo si Personalizado)", fg="gray").pack(side="left", padx=4)

    def iniciar():
        """Valida los campos (ID, edad y nº de cierres) y cierra la ventana."""
        if not e_id.get().strip():
            messagebox.showerror("Error", "Introduzca el ID del paciente")
            return
        if not e_edad.get().strip():
            messagebox.showerror("Error", "Introduzca la edad")
            return
        nivel_elegido = nivel_var.get()
        if NIVELES[nivel_elegido] is not None:
            cierres = NIVELES[nivel_elegido]
        else:
            # Personalizado: debe ser un entero positivo.
            try:
                cierres = int(e_cierres.get().strip())
                if cierres < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Numero de cierres invalido (>0)")
                return
        datos["id"]      = e_id.get().strip()
        datos["genero"]  = genero_var.get()
        datos["edad"]    = e_edad.get().strip()
        datos["cierres"] = cierres
        ventana.destroy()

    ttk.Button(ventana, text="Iniciar Evaluacion", command=iniciar).pack(pady=14)
    ventana.mainloop()
    return datos

paciente         = obtener_datos_paciente()
paciente_id      = paciente["id"]
genero           = paciente["genero"]
edad             = paciente["edad"]
TOTAL_CIERRES    = paciente["cierres"]   # objetivo de la partida
GAME_DURATION    = 25   # segundos maximos (ajustable)


# ================= ESTADO DEL JUEGO =================
score            = 0          # cierres completados
game_over        = False      # True al completar los cierres o agotar el tiempo
hand_closed      = False      # estado anterior de la mano (para detectar el flanco)
game_start_time  = None       # instante de inicio (None hasta el primer frame)
time_left        = GAME_DURATION

# ================= COOLDOWN ENTRE CIERRES =================
# Tiempo minimo (segundos) que debe pasar despues de registrar un cierre
# antes de que se pueda registrar el siguiente. Evita que la flor se
# rellene de golpe por temblores o cierres muy rapidos/repetidos.
COOLDOWN_CIERRE_S = 0.5
cooldown_hasta    = 0.0   # timestamp (time.time()) hasta el que no se acepta otro cierre


# ================= METRICAS POR CIERRE (mano global) =================
# Amplitud = apertura maxima entre cierre anterior y este
# Velocidad = velocidad de cierre maxima durante el gesto
grip_amplitudes  = []   # cm — apertura maxima de cada ciclo
grip_velocities  = []   # cm/s — velocidad de cierre de cada ciclo
grip_intervals   = []   # s — ITI entre cierres
grip_timestamps  = []   # s — instante de cada cierre desde inicio
metricas_finales = {}   # se rellena con calcular_metricas()

temblor_metricas   = None     # dict con freq/amplitud/concluyente por eje (X,Y,Z)
_temblor_procesado = False    # evita detener/calcular el temblor mas de una vez

# Acumuladores del ciclo actual (se guardan en register_close y se reinician).
_max_apertura    = 0.0  # acumulador de apertura maxima entre cierres
_max_vel_cierre  = 0.0  # velocidad maxima de cierre en el ciclo actual
ultimo_cierre_t  = None # instante del último cierre (para el ITI)

# Historico completo (para graficas PNG)
apertura_full    = []   # apertura global frame a frame
vel_full         = []   # velocidad global frame a frame
time_full        = []

# Buffers tiempo real (panel): deque circular de BUFFER_LEN puntos.
BUFFER_LEN       = 200
apertura_rt      = deque([0.0] * BUFFER_LEN, maxlen=BUFFER_LEN)
vel_rt           = deque([0.0] * BUFFER_LEN, maxlen=BUFFER_LEN)

prev_apertura    = None   # apertura del frame anterior (para derivar la velocidad)
prev_time        = None   # instante del frame anterior


# ================= FUNCIONES DE MEDIDA =================
def scale_factor(lm, cam_w, cam_h):
    """Factor de conversión píxeles -> cm para este frame.

    Se obtiene dividiendo la longitud real de referencia (9,5 cm, muñeca ->
    base del dedo medio) entre su longitud en píxeles. Devuelve 0.0 si la
    referencia mide < 1 px (mano mal detectada)."""
    wx = lm[0].x * cam_w;  wy = lm[0].y * cam_h
    mx = lm[9].x * cam_w;  my = lm[9].y * cam_h
    ref_px = math.hypot(mx - wx, my - wy)
    if ref_px < 1:
        return 0.0
    return HAND_REF_REAL_CM / ref_px


def apertura_mano_cm(lm, cam_w, cam_h, sf):
    """Apertura global de la mano (cm) = media de las distancias
    punta-muñeca de los dedos listados en APERTURA_TIPS (los 5 dedos).

    ``sf`` es el factor píxeles->cm calculado con ``scale_factor``."""
    wx = lm[0].x * cam_w;  wy = lm[0].y * cam_h
    dists = []
    for tip_idx in APERTURA_TIPS:
        tx = lm[tip_idx].x * cam_w;  ty = lm[tip_idx].y * cam_h
        dists.append(math.hypot(tx - wx, ty - wy) * sf)
    return float(np.mean(dists))


def is_finger_closed(lm, finger_def, cam_w, cam_h):
    """Un dedo está 'cerrado' si su punta está más cerca de la muñeca que su
    base. Es un criterio geométrico independiente de la escala de la mano."""
    wrist = lm[0]
    tip   = lm[finger_def["tip"]]
    base  = lm[finger_def["base"]]
    wx, wy = wrist.x * cam_w, wrist.y * cam_h
    tx, ty = tip.x   * cam_w, tip.y   * cam_h
    bx, by = base.x  * cam_w, base.y  * cam_h
    return math.hypot(tx-wx, ty-wy) < math.hypot(bx-wx, by-wy)


def is_hand_closed(lm, cam_w, cam_h):
    """La mano se considera cerrada si al menos 4 de los 5 dedos están
    cerrados (tolera que un dedo, p. ej. el pulgar, no cierre del todo)."""
    closed = sum(1 for f in FINGERS if is_finger_closed(lm, f, cam_w, cam_h))
    return closed >= 4


# ================= REGISTRO DE CIERRE =================
def register_close():
    """Registra un cierre completo: guarda las métricas del ciclo, activa el
    cooldown y comprueba si se alcanzó el objetivo de cierres."""
    global ultimo_cierre_t, score, _max_apertura, _max_vel_cierre, cooldown_hasta

    now = time.time()

    # Usar la apertura maxima del ciclo (acumulada frame a frame)
    # y la velocidad maxima de cierre (valor mas negativo del ciclo)
    grip_amplitudes.append(_max_apertura)
    grip_velocities.append(_max_vel_cierre)   # ya es magnitud positiva
    if game_start_time:
        grip_timestamps.append(now - game_start_time)

    # ITI entre cierres. Se descartan valores imposibles: < 50 ms (doble
    # detección) o > 15 s (pausa larga, no es un ritmo continuo).
    if ultimo_cierre_t is not None:
        interval = now - ultimo_cierre_t
        if 0.05 < interval < 15.0:
            grip_intervals.append(interval)

    ultimo_cierre_t  = now
    _max_apertura    = 0.0      # reinicia los acumuladores del ciclo
    _max_vel_cierre  = 0.0
    score           += 1

    # Activar el tiempo de espera obligatorio antes del proximo cierre
    cooldown_hasta = now + COOLDOWN_CIERRE_S

    # Si se alcanzó el objetivo, la partida termina sin esperar al tiempo.
    if score >= TOTAL_CIERRES:
        global game_over
        game_over = True


# ================= FIN DE PARTIDA: TEMBLOR =================
def procesar_fin_temblor():
    """Detiene la captura del micro:bit y calcula sus metricas en el
    instante en que termina la partida, para poder mostrarlas ya en la
    pantalla de fin de juego (antes de guardar nada en disco).

    Es idempotente: solo actúa la primera vez que se llama."""
    global temblor_metricas, _temblor_procesado
    if _temblor_procesado:
        return
    _temblor_procesado = True
    temblor.detener()
    temblor_metricas = temblor.calcular_metricas()


# ================= CALCULO DE METRICAS FINALES =================
def calcular_metricas():
    """Calcula las métricas clínicas finales a partir de las listas por cierre.

    Devuelve un diccionario con:
        amp_media / vel_media   : medias de amplitud (cm) y velocidad (cm/s)
        pendiente_amp / _vel    : pendiente de la recta de regresión lineal
                                  (unidad por nº de cierre)
        decrece_amp / _vel      : True si la pendiente es negativa
                                  (decremento a lo largo de la tarea)
        variabilidad_pct        : coeficiente de variación de la amplitud (%)
                                  = desviación típica / media * 100
    """
    m = {}
    n = len(grip_amplitudes)
    nv = len(grip_velocities)

    m["amp_media"] = float(np.mean(grip_amplitudes)) if n >= 1 else 0.0
    m["vel_media"] = float(np.mean(grip_velocities)) if nv >= 1 else 0.0

    # Pendiente amplitud (regresion lineal cierre a cierre).
    # Se necesitan al menos 2 cierres para ajustar una recta.
    if n >= 2:
        x = np.arange(n, dtype=float)
        pend_a, _ = np.polyfit(x, grip_amplitudes, 1)
        m["pendiente_amp"] = float(pend_a)
        m["decrece_amp"]   = pend_a < 0
    else:
        m["pendiente_amp"] = 0.0
        m["decrece_amp"]   = False

    # Pendiente velocidad
    if nv >= 2:
        x = np.arange(nv, dtype=float)
        pend_v, _ = np.polyfit(x, grip_velocities, 1)
        m["pendiente_vel"] = float(pend_v)
        m["decrece_vel"]   = pend_v < 0
    else:
        m["pendiente_vel"] = 0.0
        m["decrece_vel"]   = False

    # Variabilidad = (std / amp_media) * 100  (ddof=1 -> desviación muestral)
    if n >= 2 and m["amp_media"] > 0:
        std_a = float(np.std(grip_amplitudes, ddof=1))
        m["variabilidad_pct"] = (std_a / m["amp_media"]) * 100.0
    else:
        m["variabilidad_pct"] = 0.0

    return m


# ================= GUARDAR EXCEL =================
def guardar_resultados():
    """Añade una fila con los resultados de la sesión a ``sesiones_grip.xlsx``.

    Si el archivo no existe se crea con su fila de cabecera; si existe, se
    añade al final (así se acumulan todas las sesiones en un solo Excel).
    Las columnas de temblor quedan vacías si no hubo datos del sensor."""
    global metricas_finales
    if not metricas_finales:
        metricas_finales = calcular_metricas()
    m = metricas_finales
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    archivo = "sesiones_grip.xlsx"
    try:
        wb = load_workbook(archivo)
        ws = wb.active
    except:  # noqa: E722 - el archivo no existe aún: se crea uno nuevo
        wb = Workbook()
        ws = wb.active
        ws.title = "Sesiones"
        ws.append([
            "Fecha", "ID", "Genero", "Edad", "Cierres", "CierresSolicitados",
            "AmpMedia_cm", "VelMedia_cms",
            "PendienteAmp", "DecreceAmp",
            "PendienteVel", "DecreceVel",
            "Variabilidad_pct",
            "Temblor_FreqX_Hz", "Temblor_AmpX", "Temblor_ConcluyenteX",
            "Temblor_FreqY_Hz", "Temblor_AmpY", "Temblor_ConcluyenteY",
            "Temblor_FreqZ_Hz", "Temblor_AmpZ", "Temblor_ConcluyenteZ"
        ])

    fila = [
        fecha, paciente_id, genero, edad, score, TOTAL_CIERRES,
        round(m["amp_media"], 2),
        round(m["vel_media"], 2),
        round(m["pendiente_amp"], 6),
        "SI" if m["decrece_amp"] else "NO",
        round(m["pendiente_vel"], 6),
        "SI" if m["decrece_vel"] else "NO",
        round(m["variabilidad_pct"], 2)
    ]

    # Tres columnas por eje (frecuencia, amplitud, concluyente); None si no hay dato.
    for eje in ("X", "Y", "Z"):
        if temblor_metricas and temblor_metricas.get(eje, {}).get("freq_hz") is not None:
            d = temblor_metricas[eje]
            fila.append(round(d["freq_hz"], 2))
            fila.append(round(d["amplitud"], 1))
            fila.append("SI" if d["concluyente"] else "NO")
        else:
            fila.extend([None, None, None])

    ws.append(fila)
    wb.save(archivo)
    print("Excel guardado:", archivo)


# ================= GUARDAR GRAFICAS PNG =================
def guardar_graficas():
    """Genera un PNG 2x2 con las gráficas de la sesión:

        [0,0] amplitud por cierre + línea de tendencia + media
        [0,1] velocidad por cierre + línea de tendencia + media
        [1,0] apertura de mano continua (con los cierres marcados en rojo)
        [1,1] velocidad continua (rojo = cierre, verde = apertura)
    """
    global metricas_finales
    metricas_finales = calcular_metricas()
    m = metricas_finales

    fecha_str  = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_png = f"grafica_handgrip_{paciente_id}_{fecha_str}.png"

    fig, axes = plt.subplots(2, 2, figsize=(14, 8))
    fig.suptitle(
        f"Sesion Grip  |  ID: {paciente_id} | Genero: {genero} | Edad: {edad}",
        fontsize=13, fontweight='bold'
    )

    # -- Grafica 1: Amplitud por cierre con tendencia --
    ax1 = axes[0, 0]
    if grip_amplitudes:
        x_c = np.arange(len(grip_amplitudes))
        ax1.plot(x_c, grip_amplitudes, color='#4db8ff', linewidth=1.2,
                 marker='o', markersize=5, label='Amplitud por cierre')
        if len(grip_amplitudes) >= 2:
            coef = np.polyfit(x_c, grip_amplitudes, 1)
            ax1.plot(x_c, np.polyval(coef, x_c),
                     color='#ff4444', linewidth=1.5, linestyle='--', label='Tendencia')
        ax1.axhline(m["amp_media"], color='yellow', linewidth=0.9, linestyle=':',
                    alpha=0.8, label=f'Media={m["amp_media"]:.2f}cm')
    ax1.set_title("Amplitud de apertura por cierre")
    ax1.set_ylabel("Distancia (cm)"); ax1.set_xlabel("Cierre #")
    ax1.legend(fontsize=7); ax1.grid(True, alpha=0.3)

    # -- Grafica 2: Velocidad por cierre con tendencia --
    ax2 = axes[0, 1]
    if grip_velocities:
        x_c = np.arange(len(grip_velocities))
        ax2.plot(x_c, grip_velocities, color='#ffa040', linewidth=1.2,
                 marker='s', markersize=5, label='Velocidad por cierre')
        if len(grip_velocities) >= 2:
            coef = np.polyfit(x_c, grip_velocities, 1)
            ax2.plot(x_c, np.polyval(coef, x_c),
                     color='#ff4444', linewidth=1.5, linestyle='--', label='Tendencia')
        ax2.axhline(m["vel_media"], color='yellow', linewidth=0.9, linestyle=':',
                    alpha=0.8, label=f'Media={m["vel_media"]:.2f}cm/s')
    ax2.set_title("Velocidad de cierre por cierre")
    ax2.set_ylabel("Velocidad (cm/s)"); ax2.set_xlabel("Cierre #")
    ax2.legend(fontsize=7); ax2.grid(True, alpha=0.3)

    # -- Grafica 3: Apertura global continua --
    ax3 = axes[1, 0]
    if time_full and apertura_full:
        ax3.plot(time_full, apertura_full,
                 color='#4db8ff', linewidth=0.8, label='Apertura mano (cm)')
        if grip_timestamps and grip_amplitudes:
            ax3.scatter(grip_timestamps, grip_amplitudes,
                        color='red', s=25, zorder=5, label='Cierre')
    ax3.set_title("Apertura de mano continua")
    ax3.set_ylabel("Distancia (cm)"); ax3.set_xlabel("Tiempo (s)")
    ax3.legend(fontsize=7); ax3.grid(True, alpha=0.3)
    ax3.set_xlim(left=0); ax3.set_ylim(bottom=0)

    # -- Grafica 4: Velocidad continua (negativa = cierre, positiva = apertura) --
    ax4 = axes[1, 1]
    if time_full and vel_full:
        ax4.plot(time_full, vel_full,
                 color='#ffa040', linewidth=0.8, label='Velocidad (cm/s)')
        ax4.axhline(0, color='white', linewidth=0.6, linestyle='--', alpha=0.5)
        ax4.fill_between(time_full, vel_full, 0,
                         where=[v < 0 for v in vel_full],
                         color='#ff6060', alpha=0.25, label='Cierre')
        ax4.fill_between(time_full, vel_full, 0,
                         where=[v > 0 for v in vel_full],
                         color='#60ff80', alpha=0.25, label='Apertura')
    ax4.set_title("Velocidad cierre/apertura continua")
    ax4.set_ylabel("Velocidad (cm/s)"); ax4.set_xlabel("Tiempo (s)")
    ax4.legend(fontsize=7); ax4.grid(True, alpha=0.3)
    ax4.set_xlim(left=0)

    plt.tight_layout()
    plt.savefig(nombre_png, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print("Grafica guardada:", nombre_png)


# ================= HELPERS DE DIBUJO =================
def put_text(frame, text, pos, scale=0.8, color=(255, 255, 255), thickness=2):
    """Escribe texto en el frame. OpenCV no soporta tildes ni 'ñ' con las
    fuentes Hershey, así que se convierten a '?' para evitar caracteres rotos."""
    safe = text.encode('ascii', 'replace').decode('ascii')
    cv2.putText(frame, safe, pos, cv2.FONT_HERSHEY_SIMPLEX,
                scale, color, thickness, cv2.LINE_AA)


def draw_temblor_block(frame, x, y, scale=0.62):
    """Dibuja frecuencia y amplitud por eje (X, Y, Z) del temblor.

    Verde = pico concluyente; ocre = no concluyente (probable ruido);
    gris = sin datos. Devuelve la coordenada y tras el bloque dibujado."""
    put_text(frame, "-- Temblor (acelerometro) --", (x, y),
             scale=0.55, color=(160, 160, 180), thickness=1); y += 26

    if not temblor.conectado:
        put_text(frame, "Sensor no conectado", (x, y),
                 scale=scale, color=(150, 150, 150), thickness=1); y += 28
        return y

    if temblor_metricas is None:
        put_text(frame, "Datos insuficientes", (x, y),
                 scale=scale, color=(150, 150, 150), thickness=1); y += 28
        return y

    for eje in ("X", "Y", "Z"):
        d = temblor_metricas[eje]
        if d["freq_hz"] is None:
            texto = f"Eje {eje}: sin datos"
            color = (120, 120, 120)
        else:
            texto = f"Eje {eje}: {d['freq_hz']:.2f} Hz   amp {d['amplitud']:.0f}"
            if not d["concluyente"]:
                texto += "  (no concluyente / ruido)"
            color = (80, 220, 80) if d["concluyente"] else (180, 160, 90)
        put_text(frame, texto, (x, y), scale=scale, color=color, thickness=1)
        y += 26
    return y


def draw_graph(frame, data, gx, gy, gw, gh,
               color, label, unit, y_min=None, y_max=None, zero_line=False):
    """Dibuja una gráfica de línea sencilla (estilo osciloscopio) con OpenCV.

    Parameters
    ----------
    data : iterable de float
        Valores a representar (normalmente un deque circular).
    gx, gy, gw, gh : int
        Posición (esquina superior izquierda) y tamaño del recuadro en píxeles.
    y_min, y_max : float | None
        Límites del eje Y; si son None se autoescalan con los datos.
    zero_line : bool
        Si True y el rango cruza el 0, dibuja la línea del cero.
    """
    # Fondo
    cv2.rectangle(frame, (gx, gy), (gx+gw, gy+gh), (40, 40, 55), -1)
    vals = list(data)
    if not vals:
        return
    lo  = y_min if y_min is not None else min(vals)
    hi  = y_max if y_max is not None else max(vals)
    rng = hi - lo if hi != lo else 1.0   # evita dividir por cero con señal plana
    if zero_line and lo < 0 < hi:
        yz = gy + gh - int((0 - lo) / rng * gh)
        cv2.line(frame, (gx, yz), (gx+gw, yz), (80, 80, 100), 1)
    # Rejilla horizontal central
    for frac in [0.5]:
        yg = gy + int((1 - frac) * gh)
        cv2.line(frame, (gx, yg), (gx+gw, yg), (55, 55, 70), 1)

    # Conversión de cada valor a coordenadas de píxel (el eje Y está invertido
    # en las imágenes: y=0 es la parte superior).
    step = gw / max(len(vals) - 1, 1)
    pts  = []
    for idx, v in enumerate(vals):
        px_ = int(gx + idx * step)
        py_ = int(gy + gh - ((v - lo) / rng) * gh)
        py_ = max(gy, min(gy + gh, py_))   # recorta al recuadro
        pts.append((px_, py_))
    for idx in range(1, len(pts)):
        cv2.line(frame, pts[idx-1], pts[idx], color, 1)

    # Borde y valor actual
    cv2.rectangle(frame, (gx, gy), (gx+gw, gy+gh), (90, 90, 120), 1)
    cur = vals[-1]
    put_text(frame, f"{label}:{cur:.3f}{unit}",
             (gx, gy - 4), scale=0.38, color=color, thickness=1)


# ================= VISUAL DE FLOR =================
FLOR_CX = GAME_W // 2        # centro de la flor (x)
FLOR_CY = WINDOW_H // 2      # centro de la flor (y); subido un poco para dejar espacio al tallo

PETALO_R = 85    # radio de distribucion de petalos
PUNTO_R  = 22    # radio de cada punto (mas grande)

def _posiciones_flor(total):
    """Calcula la posición y el color de cada punto de la flor.

    Devuelve una lista de ``(x, y, color_BGR)`` con ``total`` elementos:
    el primero es el centro (amarillo) y el resto son pétalos (blancos)
    repartidos uniformemente en círculo, empezando por la parte superior."""
    posiciones = []
    # Centro amarillo
    posiciones.append((FLOR_CX, FLOR_CY, (0, 215, 255)))

    n_petalos = total - 1
    if n_petalos <= 0:
        return posiciones
    for k in range(n_petalos):
        # -90º = arriba; se avanza 360/n_petalos grados por pétalo.
        angulo = math.radians(-90 + k * 360.0 / n_petalos)
        px_ = int(FLOR_CX + PETALO_R * math.cos(angulo))
        py_ = int(FLOR_CY + PETALO_R * math.sin(angulo))
        posiciones.append((px_, py_, (240, 240, 240)))
    return posiciones

# Se calcula una sola vez, ya que TOTAL_CIERRES no cambia durante la partida.
_flor_pos = _posiciones_flor(TOTAL_CIERRES)


def draw_flor(frame):
    """
    Dibuja la flor con tallo y hoja.
    Cada cierre rellena un punto: primero el centro (amarillo), luego los petalos (blancos).
    Al completar: rectangulo verde fino alrededor.
    Mientras el cooldown esta activo, el siguiente punto se muestra "congelado"
    (sin parpadeo) para indicar que hay que esperar.
    """
    completados = score

    # --- Tallo ---
    tallo_top = FLOR_CY + PETALO_R + PUNTO_R + 2
    tallo_bot = FLOR_CY + PETALO_R + PUNTO_R + 70
    cv2.line(frame, (FLOR_CX, tallo_top), (FLOR_CX, tallo_bot), (34, 139, 34), 3)

    # Hoja izquierda (elipse inclinada)
    hoja_cx = FLOR_CX - 18
    hoja_cy = tallo_top + 28
    cv2.ellipse(frame, (hoja_cx, hoja_cy), (22, 10), -35, 0, 360, (34, 160, 34), -1)
    cv2.ellipse(frame, (hoja_cx, hoja_cy), (22, 10), -35, 0, 360, (20, 110, 20),  1)

    # Guias vacias (circulos tenues) para todas las posiciones
    for i, (px_, py_, col) in enumerate(_flor_pos):
        cv2.circle(frame, (px_, py_), PUNTO_R, (90, 90, 90), 1)

    # Puntos completados (solidos)
    for i in range(min(completados, len(_flor_pos))):
        px_, py_, col = _flor_pos[i]
        cv2.circle(frame, (px_, py_), PUNTO_R, col, -1)
        border_col = (0, 160, 200) if i == 0 else (160, 160, 160)
        cv2.circle(frame, (px_, py_), PUNTO_R, border_col, 2)

    # Punto siguiente
    if completados < len(_flor_pos):
        nx, ny, nc = _flor_pos[completados]
        en_cooldown = time.time() < cooldown_hasta
        if en_cooldown:
            # Congelado: aro fijo en gris/naranja para indicar "espera"
            cv2.circle(frame, (nx, ny), PUNTO_R + 6, (90, 110, 140), 2)
        else:
            # Parpadeo animado: ya se puede registrar el siguiente cierre
            t_anim = int(time.time() * 3) % 2
            cv2.circle(frame, (nx, ny), PUNTO_R + 5 + t_anim * 5, nc, 1)

    # Rectangulo verde fino al completar
    if completados >= TOTAL_CIERRES:
        margen = PETALO_R + PUNTO_R + 14
        x1 = FLOR_CX - margen;  y1 = FLOR_CY - margen
        x2 = FLOR_CX + margen;  y2 = FLOR_CY + margen
        cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 220, 80), 2)
        put_text(frame, "Completado!", (FLOR_CX - 82, y2 + 28),
                 scale=0.78, color=(0, 220, 80), thickness=2)

    # Contador debajo del tallo
    put_text(frame, f"{min(completados, TOTAL_CIERRES)}/{TOTAL_CIERRES}",
             (FLOR_CX - 22, tallo_bot + 22),
             scale=0.72, color=(100, 160, 100), thickness=1)


# ================= PANEL LATERAL =================
def draw_panel(frame):
    """Panel lateral DURANTE la partida: cierres, ITI/amplitud/velocidad
    medias y las gráficas en tiempo real de apertura y velocidad."""
    px = GAME_W
    cv2.rectangle(frame, (px, 0), (WINDOW_W, WINDOW_H), (25, 25, 35), -1)
    cv2.line(frame, (px, 0), (px, WINDOW_H), (70, 70, 100), 2)

    put_text(frame, "GRIP MONITOR", (px+15, 32), scale=0.72, color=(180, 220, 255))

    y = 58
    put_text(frame, f"Cierres: {score}/{TOTAL_CIERRES}", (px+12, y),
             scale=0.58, color=(150, 200, 255), thickness=1); y += 24

    if grip_intervals:
        avg_iti = sum(grip_intervals) / len(grip_intervals)
        put_text(frame, f"ITI medio: {avg_iti:.3f} s", (px+12, y),
                 scale=0.52, color=(200, 255, 200), thickness=1)
    else:
        put_text(frame, "ITI medio: --", (px+12, y),
                 scale=0.52, color=(140, 140, 140), thickness=1)
    y += 22

    if grip_amplitudes:
        avg_amp = sum(grip_amplitudes) / len(grip_amplitudes)
        put_text(frame, f"Amp. media: {avg_amp:.2f} cm", (px+12, y),
                 scale=0.52, color=(130, 200, 255), thickness=1)
    else:
        put_text(frame, "Amp. media: --", (px+12, y),
                 scale=0.52, color=(140, 140, 140), thickness=1)
    y += 22

    if grip_velocities:
        avg_vel = sum(grip_velocities) / len(grip_velocities)
        put_text(frame, f"Vel. media: {avg_vel:.2f} cm/s", (px+12, y),
                 scale=0.52, color=(255, 210, 120), thickness=1)
    else:
        put_text(frame, "Vel. media: --", (px+12, y),
                 scale=0.52, color=(140, 140, 140), thickness=1)
    y += 26

    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1)
    y += 12

    gw = PANEL_W - 28
    gh = 100
    gx = px + 14

    put_text(frame, "Apertura mano (cm):", (gx, y),
             scale=0.44, color=(160, 200, 255), thickness=1); y += 13
    draw_graph(frame, apertura_rt, gx, y, gw, gh,
               color=(80, 180, 255), label="Ap", unit="cm", y_min=0); y += gh + 18

    put_text(frame, "Velocidad cierre (cm/s):", (gx, y),
             scale=0.44, color=(255, 200, 100), thickness=1); y += 13
    draw_graph(frame, vel_rt, gx, y, gw, gh,
               color=(255, 160, 40), label="Vel", unit="cm/s",
               y_min=None, zero_line=True); y += gh + 10

    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1)
    y += 6
    put_text(frame, "Vel<0: cierre  |  Vel>0: apertura", (px+10, y),
             scale=0.35, color=(100, 100, 130), thickness=1)

    # FPS medio de los últimos 30 frames (fps_buffer se define más abajo,
    # antes del bucle principal)
    fps_val = sum(fps_buffer) / len(fps_buffer) if fps_buffer else 0
    put_text(frame, f"FPS: {fps_val:.0f}", (px+12, WINDOW_H - 12),
             scale=0.42, color=(80, 80, 110), thickness=1)


# ================= PANEL RESULTADOS FINALES =================
def draw_panel_results(frame):
    """Panel lateral AL TERMINAR: resumen de métricas (amplitud, velocidad,
    tendencias, variabilidad con barra y temblor)."""
    global metricas_finales
    if not metricas_finales:
        metricas_finales = calcular_metricas()
    m = metricas_finales

    px = GAME_W
    cv2.rectangle(frame, (px, 0), (WINDOW_W, WINDOW_H), (25, 25, 35), -1)
    cv2.line(frame, (px, 0), (px, WINDOW_H), (70, 70, 100), 2)

    put_text(frame, "RESULTADOS", (px+30, 38), scale=0.9, color=(34, 197, 94), thickness=2)

    y = 65
    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1); y += 16

    put_text(frame, f"Cierres: {score}/{TOTAL_CIERRES}",
             (px+15, y), scale=0.72, color=(180, 220, 255)); y += 32

    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1); y += 14

    # Amplitud
    put_text(frame, "-- Amplitud de apertura --",
             (px+12, y), scale=0.55, color=(160, 160, 180), thickness=1); y += 26
    put_text(frame, f"Amplitud media: {m['amp_media']:.2f} cm",
             (px+12, y), scale=0.68, color=(130, 200, 255)); y += 32

    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1); y += 14

    # Velocidad
    put_text(frame, "-- Velocidad de cierre --",
             (px+12, y), scale=0.55, color=(160, 160, 180), thickness=1); y += 26
    put_text(frame, f"Velocidad media: {m['vel_media']:.2f} cm/s",
             (px+12, y), scale=0.68, color=(255, 210, 100)); y += 32

    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1); y += 14

    # Pendientes (verde = sin decremento; rojo en BGR = pendiente negativa)
    put_text(frame, "-- Tendencia (pendiente) --",
             (px+12, y), scale=0.55, color=(160, 160, 180), thickness=1); y += 26
    col_a = (80, 220, 80) if not m["decrece_amp"] else (80, 80, 255)
    txt_a = "SI DECRECE" if m["decrece_amp"] else "NO decrece"
    put_text(frame, f"Amp: {m['pendiente_amp']:+.5f}  [{txt_a}]",
             (px+12, y), scale=0.58, color=col_a); y += 28
    col_v = (80, 220, 80) if not m["decrece_vel"] else (80, 80, 255)
    txt_v = "SI DECRECE" if m["decrece_vel"] else "NO decrece"
    put_text(frame, f"Vel: {m['pendiente_vel']:+.5f}  [{txt_v}]",
             (px+12, y), scale=0.58, color=col_v); y += 34

    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1); y += 14

    # Variabilidad (umbral visual del 20 %: por debajo verde, por encima naranja)
    put_text(frame, "-- Variabilidad --",
             (px+12, y), scale=0.55, color=(160, 160, 180), thickness=1); y += 26
    vc = (200, 255, 180) if m["variabilidad_pct"] < 20 else (255, 160, 80)
    put_text(frame, f"Variabilidad: {m['variabilidad_pct']:.1f} %",
             (px+12, y), scale=0.72, color=vc); y += 28
    # Barra de variabilidad: la barra completa equivale al 50 %.
    bx = px + 12; bw = PANEL_W - 28; bh = 10
    pct_c = min(m["variabilidad_pct"] / 50.0, 1.0)
    cv2.rectangle(frame, (bx, y), (bx + bw, y + bh), (60, 60, 80), -1)
    fc = (80, 200, 80) if m["variabilidad_pct"] < 20 else (255, 160, 50)
    cv2.rectangle(frame, (bx, y), (bx + int(bw * pct_c), y + bh), fc, -1)
    cv2.rectangle(frame, (bx, y), (bx + bw, y + bh), (120, 120, 150), 1)
    y += bh + 16

    cv2.line(frame, (px+10, y), (WINDOW_W-10, y), (60, 60, 80), 1); y += 14
    draw_temblor_block(frame, px+12, y, scale=0.5)


# ================= PANTALLA GAME OVER =================
def draw_game_over(frame):
    """Pantalla de fin de partida: oscurece el juego y muestra el resumen
    de métricas sobre él. El título cambia según se completaran o no los
    cierres solicitados."""
    global metricas_finales
    if not metricas_finales:
        metricas_finales = calcular_metricas()
    m = metricas_finales

    # Oscurecer el fondo con una capa negra semitransparente (60 %).
    overlay = frame.copy()
    cv2.rectangle(overlay, (0, 0), (GAME_W, WINDOW_H), (0, 0, 0), -1)
    cv2.addWeighted(overlay, 0.6, frame, 0.4, 0, frame)
    cx = GAME_W // 2

    titulo = "Completado!" if score >= TOTAL_CIERRES else "Tiempo terminado!"
    put_text(frame, titulo, (cx-180, WINDOW_H//2-155),
             scale=1.15, color=(34, 197, 94), thickness=3)
    put_text(frame, f"Cierres: {score}/{TOTAL_CIERRES}",
             (cx-90, WINDOW_H//2-112), scale=0.9)

    y = WINDOW_H // 2 - 75
    put_text(frame, f"Amplitud media:   {m['amp_media']:.2f} cm",
             (cx-235, y), scale=0.75, color=(130, 200, 255)); y += 34
    put_text(frame, f"Velocidad media:  {m['vel_media']:.2f} cm/s",
             (cx-235, y), scale=0.75, color=(255, 210, 120)); y += 34
    col_a = (80, 220, 80) if not m["decrece_amp"] else (80, 100, 255)
    dec_a = "DECRECE" if m["decrece_amp"] else "estable"
    put_text(frame, f"Tend. amplitud:   {m['pendiente_amp']:+.5f}  ({dec_a})",
             (cx-235, y), scale=0.72, color=col_a); y += 32
    col_v = (80, 220, 80) if not m["decrece_vel"] else (80, 100, 255)
    dec_v = "DECRECE" if m["decrece_vel"] else "estable"
    put_text(frame, f"Tend. velocidad:  {m['pendiente_vel']:+.5f}  ({dec_v})",
             (cx-235, y), scale=0.72, color=col_v); y += 34
    vc = (200, 255, 160) if m["variabilidad_pct"] < 20 else (255, 160, 80)
    put_text(frame, f"Variabilidad:     {m['variabilidad_pct']:.1f} %",
             (cx-235, y), scale=0.75, color=vc); y += 40

    y = draw_temblor_block(frame, cx-235, y, scale=0.65)

    put_text(frame, "Pulsa Q para salir", (cx-130, y),
             scale=0.72, color=(150, 150, 150), thickness=1)


# ================= BARRA SUPERIOR =================
def draw_top_bar(frame):
    """Barra superior: título, datos del paciente, tiempo restante y cierres."""
    cv2.rectangle(frame, (0, 0), (GAME_W, 65), (21, 94, 88), -1)
    put_text(frame, "Rehabilitacion - Cierre de Mano", (20, 40), scale=0.9)
    put_text(frame, f"ID: {paciente_id} | Genero: {genero} | Edad: {edad}",
             (20, 60), scale=0.42)
    put_text(frame, f"Tiempo: {max(time_left, 0)}s", (GAME_W - 190, 40), scale=0.85)
    put_text(frame, f"Cierres: {score}/{TOTAL_CIERRES}",
             (20, 90), scale=0.6, color=(20, 180, 150))


# ================= FPS =================
fps_buffer   = deque(maxlen=30)   # FPS instantáneos de los últimos 30 frames
last_frame_t = time.time()


# ================= LOOP PRINCIPAL =================
print("Iniciando... Pulsa Q para salir.")

while True:
    now_t = time.time()

    # --- 1) Captura del frame -------------------------------------------------
    ret, cam_frame = cap.read()
    if not ret:
        print("ERROR: No se puede leer la camara.")
        break

    fps_buffer.append(1.0 / max(now_t - last_frame_t, 1e-6))
    last_frame_t = now_t

    # Efecto espejo: así el movimiento en pantalla coincide con el de la mano.
    cam_frame = cv2.flip(cam_frame, 1)
    cam_h, cam_w = cam_frame.shape[:2]

    # MediaPipe espera RGB; OpenCV trabaja en BGR.
    rgb     = cv2.cvtColor(cam_frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb)

    # Inicio de la partida: se arranca el sensor y DESPUÉS el cronómetro.
    if game_start_time is None:
        temblor.iniciar()               # bloquea ~2 s (arranque del micro:bit)
        game_start_time = time.time()   # el cronometro empieza DESPUES de la espera
        now_t = game_start_time         # refrescar: now_t se tomo antes de la espera

    # --- 2) Medida de la mano -------------------------------------------------
    if results.multi_hand_landmarks and not game_over:
        # Dibuja el esqueleto de la mano sobre la imagen de cámara.
        for hl in results.multi_hand_landmarks:
            mp_draw.draw_landmarks(
                cam_frame, hl,
                mp_hands.HAND_CONNECTIONS,
                mp_draw_styles.get_default_hand_landmarks_style(),
                mp_draw_styles.get_default_hand_connections_style()
            )

        lm = results.multi_hand_landmarks[0].landmark
        sf = scale_factor(lm, cam_w, cam_h)

        # Apertura global de la mano en centimetros
        ap_cm = apertura_mano_cm(lm, cam_w, cam_h, sf)
        apertura_rt.append(ap_cm)

        # Velocidad (derivada numérica de la apertura entre dos frames).
        # Negativa = la mano se cierra; positiva = se abre.
        dt = now_t - prev_time if prev_time is not None else 0
        if dt > 0 and prev_apertura is not None:
            vel = (ap_cm - prev_apertura) / dt
        else:
            vel = 0.0
        vel_rt.append(vel)

        # Guardar historico
        apertura_full.append(ap_cm)
        vel_full.append(vel)
        time_full.append(now_t - game_start_time)

        # Acumular apertura maxima entre cierres (solo si la mano es visible y real)
        if ap_cm > 2.0:   # umbral minimo para ignorar frames de ruido/mano oculta
            if ap_cm > _max_apertura:
                _max_apertura = ap_cm
            # Velocidad de cierre maxima: la derivada mas negativa (cierre rapido)
            if vel < 0 and abs(vel) > _max_vel_cierre:
                _max_vel_cierre = abs(vel)

        prev_apertura = ap_cm
        prev_time     = now_t

        # --- 3) Detección de cierre ------------------------------------------
        # Un cierre es el FLANCO de "abierta" -> "cerrada" y solo se acepta
        # fuera del cooldown obligatorio.
        closed_now  = is_hand_closed(lm, cam_w, cam_h)
        en_cooldown = now_t < cooldown_hasta
        if closed_now and not hand_closed and not en_cooldown:
            register_close()
        hand_closed = closed_now

    else:
        # Mano no detectada (o partida terminada): se rellenan ceros para
        # mantener alineadas las series.
        apertura_rt.append(0.0)
        vel_rt.append(0.0)
        if game_start_time and not game_over:
            apertura_full.append(0.0)
            vel_full.append(0.0)
            time_full.append(now_t - game_start_time)
        hand_closed = False

    # --- 4) Temporizador (solo si no se completó por cierres) -----------------
    if not game_over:
        elapsed   = now_t - game_start_time
        time_left = max(0, GAME_DURATION - int(elapsed))
        if time_left <= 0:
            game_over = True

    # En cuanto acaba la partida se cierra la captura de temblor y se calculan
    # las métricas (para mostrarlas ya en la pantalla final).
    if game_over:
        procesar_fin_temblor()

    # --- 5) Dibujo ------------------------------------------------------------
    # Lienzo base (BGR): zona de juego clara + panel lateral oscuro.
    frame = np.full((WINDOW_H, WINDOW_W, 3), (245, 250, 255), dtype=np.uint8)
    frame[:, GAME_W:] = (25, 25, 35)

    # Mensaje de ayuda: espera obligatoria o instrucción de cierre.
    if not game_over:
        restante_cd = cooldown_hasta - now_t
        if restante_cd > 0:
            put_text(frame, f"Espera {restante_cd:.1f}s para el siguiente cierre",
                     (GAME_W//2 - 230, 115), scale=0.66, color=(255, 160, 60))
        else:
            put_text(frame, "Cierra la mano completamente",
                     (GAME_W//2 - 195, 115), scale=0.72, color=(100, 100, 100))

    draw_flor(frame)
    draw_top_bar(frame)

    if game_over:
        draw_panel_results(frame)
        draw_game_over(frame)
    else:
        draw_panel(frame)

    # Cámara en vivo (miniatura en la esquina inferior izquierda)
    cam_disp_w  = 240
    cam_disp_h  = int(cam_disp_w * cam_h / cam_w)   # mantiene la proporción
    cam_resized = cv2.resize(cam_frame, (cam_disp_w, cam_disp_h))
    y1 = WINDOW_H - cam_disp_h - 10;  y2 = WINDOW_H - 10
    x1 = 10;                           x2 = 10 + cam_disp_w
    frame[y1:y2, x1:x2] = cam_resized
    cv2.rectangle(frame, (x1-2, y1-2), (x2+2, y2+2), (21, 94, 88), 2)
    put_text(frame, "Camara en vivo", (x1, y1-6),
             scale=0.5, color=(21, 94, 88), thickness=1)

    cv2.imshow(WIN_NAME, frame)
    # waitKey(1) procesa los eventos de la ventana; 'q' termina el programa.
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# ================= CIERRE =================
# Se cierra el sensor por si la ventana se cerró antes de ver la pantalla final.
procesar_fin_temblor()   # por si se cerro la ventana antes de ver la pantalla de fin

# Guardado de resultados en disco (solo ocurre al salir del bucle).
if GUARDAR_CSV_TEMBLOR:
    temblor.guardar_csv(paciente_id)
temblor.guardar_grafica(paciente_id)

guardar_resultados()
guardar_graficas()

# Liberar recursos
cap.release()
cv2.destroyAllWindows()
hands.close()

# Resumen final en consola
print("\n--- RESULTADOS FINALES ---")
print(f"ID: {paciente_id} | Genero: {genero} | Edad: {edad}")
print(f"Cierres: {score}/{TOTAL_CIERRES}")
if metricas_finales:
    m = metricas_finales
    print(f"Amplitud media:    {m['amp_media']:.2f} cm")
    print(f"Velocidad media:   {m['vel_media']:.2f} cm/s")
    print(f"Pendiente amp:     {m['pendiente_amp']:+.6f}  -> {'DECRECE' if m['decrece_amp'] else 'No decrece'}")
    print(f"Pendiente vel:     {m['pendiente_vel']:+.6f}  -> {'DECRECE' if m['decrece_vel'] else 'No decrece'}")
    print(f"Variabilidad:      {m['variabilidad_pct']:.1f} %")
if temblor_metricas:
    for eje in ("X", "Y", "Z"):
        d = temblor_metricas[eje]
        if d["freq_hz"] is not None:
            nota = "" if d["concluyente"] else "  (no concluyente / ruido)"
            print(f"Temblor eje {eje}: {d['freq_hz']:.2f} Hz  amp={d['amplitud']:.0f}{nota}")
