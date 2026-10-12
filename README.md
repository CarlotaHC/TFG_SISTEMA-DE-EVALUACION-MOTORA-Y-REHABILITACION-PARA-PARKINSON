# TFG: Diseño e implementación de un prototipo basado en visión artificial para el análisis de movimientos en el contexto de la enfermedad de Parkinson

Este repositorio contiene el código desarrollado como parte del Trabajo de Fin de Grado (TFG) de Carlota Hernández Carrero (Ingeniería de la Salud, ETSI Informática, Universidad de Sevilla), orientado al diseño e implementación de un prototipo software capaz de evaluar la función motora de la mano mediante visión por computador y de medir el temblor con un acelerómetro micro:bit.

- **Curso:** 2025-2026

## Contexto y Motivación

La enfermedad de Parkinson se caracteriza por signos motores como la bradicinesia (lentitud y reducción progresiva de la amplitud del movimiento) y el temblor de reposo. Su valoración clínica se apoya en pruebas como el *finger tapping* y el cierre repetido de la mano, evaluadas habitualmente de forma visual y subjetiva por el profesional.

Este TFG propone una alternativa accesible y cuantitativa: un sistema basado en una cámara web y una placa micro:bit que, sin necesidad de hardware especializado, es capaz de:

- Detectar los 21 puntos de la mano en tiempo real mediante MediaPipe y OpenCV.
- Calcular parámetros motores asociados a la bradicinesia: **amplitud**, **velocidad**, **tendencia (decremento)** y **variabilidad** de los movimientos.
- Medir el temblor con un acelerómetro (micro:bit, 20 Hz) y obtener su **frecuencia dominante** mediante FFT.
- Integrar las pruebas en dos minijuegos que facilitan la realización y el seguimiento de los ejercicios de rehabilitación:
  - **Fingertap (juego del pájaro):** finger tapping índice-pulgar.
  - **Handgrip (juego de la flor):** cierre completo de la mano.

Los resultados se guardan en archivos Excel y en gráficas PNG.

## Estructura del Proyecto

```
├── microbit/
│   └── acelerometro_microbit.py   # Programa de la micro:bit (MakeCode, Python). Envía "t,x,y,z" a 20 Hz por puerto serie
├── src/
│   ├── temblor_sensor.py          # Captura en segundo plano y análisis FFT del temblor
│   ├── fingertap_pajaro.py        # Juego Fingertap (pájaro)
│   └── handgrip_flor.py           # Juego Handgrip (flor)
├── requirements.txt               # Dependencias de Python
└── README.md
```

## Requisitos

**Hardware**

- Ordenador con cámara web.
- micro:bit v1 con cable USB (opcional: sin ella los juegos funcionan, pero no se registra el temblor).

**Software**

- Python 3.10 (recomendado). `mediapipe 0.10.9` no es compatible con versiones muy recientes de Python.
- Tkinter (incluido con Python en Windows y macOS; en Linux: `sudo apt install python3-tk`).
- Se ha utilizado Visual Studio Code para ejecutar los minijuegos.

**Dependencias** (ver `requirements.txt`):

```bash
pip install -r requirements.txt
```

Las principales librerías utilizadas son:

| Librería | Uso |
|---|---|
| opencv-python | Captura de cámara y dibujo de la interfaz |
| mediapipe==0.10.9 | Detección de los 21 puntos de la mano |
| numpy | Cálculo numérico |
| scipy | Filtro Butterworth y detección de picos |
| matplotlib | Generación de las gráficas PNG |
| openpyxl | Exportación a Excel |
| pyserial | Comunicación con la micro:bit |

> [!WARNING]
> Instalar `pyserial`, **no** `serial`: son paquetes distintos.

## Preparar la micro:bit

1. Abrir [MakeCode](https://makecode.microbit.org) y crear un proyecto nuevo.
2. Cambiar a modo Python y pegar el contenido de `microbit/acelerometro_microbit.py`.
3. Descargar el `.hex` y copiarlo a la unidad `MICROBIT`.
4. Con la micro:bit conectada por USB, averiguar el puerto serie.
5. Editar en los dos juegos la línea siguiente y poner el puerto de su equipo:

```python
PUERTO_MICROBIT = "COM3"
```

> [!NOTE]
> El puerto serie no puede estar abierto en otro programa (por ejemplo, la consola de MakeCode) mientras se ejecuta el juego.

## Uso

### 1. Ejecutar un minijuego

Desde la carpeta del repositorio:

```bash
python src/fingertap_pajaro.py
python src/handgrip_flor.py
```

Colocar la micro:bit en la mano contralateral (la que no juega) para medir el temblor en reposo.

### 2. Registro del paciente

En la ventana de registro, introducir el ID del paciente, género, edad y nivel de dificultad, y pulsar **Iniciar Evaluación**.

### 3. Realizar la prueba

Colocar la mano frente a la cámara, con buena iluminación.

| Juego | Ejercicio | Duración |
|---|---|---|
| Fingertap | Unir índice y pulgar repetidamente. El pájaro sube con cada tap y cae si se para | 20 s |
| Handgrip | Cerrar la mano por completo y volver a abrirla en cada repetición. Cada cierre rellena un punto de la flor | Hasta completar los cierres elegidos (5, 10, 15 o personalizado) o 25 s |

### 4. Resultados

Al terminar se muestra la pantalla de resultados (amplitud, velocidad, tendencia, variabilidad y temblor por eje).

## Controles durante la ejecución

| Tecla | Acción |
|---|---|
| q | Cerrar el juego y guardar los archivos de salida |

## Resultados generados

Se generan en la carpeta desde la que se ejecuta el programa:

| Archivo | Contenido |
|---|---|
| `sesiones_fingertap.xlsx` | Una fila por sesión (juego Fingertap) |
| `sesiones_grip.xlsx` | Una fila por sesión (juego Handgrip) |
| `grafica_fingertap_*.png` | Gráficas de la sesión Fingertap |
| `grafica_handgrip_*.png` | Gráficas de la sesión Handgrip |
| `temblor_fft_*.png` | Espectro FFT del temblor (ejes X, Y, Z) |
| `temblor_*.csv` | Datos crudos del acelerómetro (solo si `GUARDAR_CSV_TEMBLOR = True`) |

## Parámetros ajustables

Los siguientes parámetros pueden modificarse directamente en el código de los juegos:

| Parámetro | Descripción |
|---|---|
| `PUERTO_MICROBIT` | Puerto serie de la micro:bit |
| `GAME_DURATION` | Duración máxima de la partida (segundos) |
| `GUARDAR_CSV_TEMBLOR` | Guarda o no el CSV crudo del acelerómetro |
| `dist_cm < 1.0` | Umbral de tap en Fingertap (cm) |
| `COOLDOWN_CIERRE_S` | Espera mínima entre cierres en Handgrip |

## Solución de problemas

| Problema | Solución |
|---|---|
| "No se puede leer la cámara" | Cerrar otras aplicaciones que usen la cámara o cambiar `cv2.VideoCapture(0)` por `cv2.VideoCapture(1)` |
| "Sensor no conectado" | Comprobar el puerto en `PUERTO_MICROBIT` y que ningún otro programa lo tenga abierto |
| "Datos insuficientes" en el temblor | La captura es demasiado corta (mínimo 30 muestras, aprox. 1,5 s) |
| Error al instalar mediapipe | Comprobar la versión de Python (ver Requisitos) |

## Notas

> [!NOTE]
> El sistema está pensado para usarse con buena iluminación y la mano completamente visible frente a la cámara. El rendimiento puede variar en otras condiciones.

> [!IMPORTANT]
> Proyecto académico. No es un producto sanitario ni sustituye la valoración clínica profesional.

## Cita

Para citar este software:

> Hernández Carrero, Carlota. (2026). *Diseño e implementación de un prototipo basado en visión artificial para el análisis de movimientos en el contexto de la enfermedad de Parkinson*. GitHub. https://github.com/usuario/repositorio
