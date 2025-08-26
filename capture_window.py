# -*- coding: utf-8 -*-
"""
Captura una ventana en Windows, extrae el número de serie tras la etiqueta 'Carrete:'
y guarda un JPG nombrado con ese serial. Prioriza accesibilidad; usa OCR como respaldo.

Dependencias:
  pip install pyautogui pygetwindow pillow pywinauto pytesseract opencv-python

Si Tesseract no está en PATH, ajusta DEFAULT_TESSERACT_PATH.
"""

import argparse
import re
import sys
import time
from pathlib import Path

# ------------------ CONFIGURACIÓN PREDETERMINADA ------------------
DEFAULT_WINDOW_TITLE = "Pedido.txt - Notepad"  # Cambia al título de tu aplicación
DEFAULT_OUTPUT_DIR = r"C:/Users/garci/Desktop/capturar_pantalla/img"  # Carpeta destino
DEFAULT_TESSERACT_PATH = r"C:/Program Files/Tesseract-OCR/tesseract.exe"  # tesseract.exe

# ------------------ IMPORTS CON CONTROL DE ERRORES ------------------
try:
    import pyautogui
    import pygetwindow as gw
except Exception as e:
    print("Faltan dependencias base (pyautogui, pygetwindow, pillow).")
    print("Instala: pip install pyautogui pygetwindow pillow")
    sys.exit(1)

try:
    from pywinauto import Application
    from pywinauto.controls.uia_controls import EditWrapper
    PYWINAUTO_OK = True
except Exception:
    PYWINAUTO_OK = False

try:
    import pytesseract
    import cv2
    import numpy as np
    from PIL import Image
    OCR_OK = True
except Exception:
    OCR_OK = False

# ------------------ UTILIDADES BÁSICAS ------------------
def ensure_windows():
    if sys.platform != "win32":
        raise OSError("Este script está diseñado para Windows.")

def find_window(title_substr: str | None):
    """Ventana activa si no se pasa título; si se pasa, busca exacto y luego por substring."""
    if title_substr is None:
        return gw.getActiveWindow()
    wins = gw.getAllWindows()
    exact = [w for w in wins if w.title.strip() == title_substr.strip()]
    if exact:
        return exact[0]
    t = title_substr.lower()
    partial = [w for w in wins if t in w.title.lower()]
    return partial[0] if partial else None

def bring_to_front(win):
    try:
        win.activate(); time.sleep(0.25)
        if win.isMinimized:
            win.restore(); time.sleep(0.25)
        win.activate(); time.sleep(0.25)
    except Exception:
        pass

def get_window_box(win):
    left, top, right, bottom = win.left, win.top, win.right, win.bottom
    w = max(0, right - left); h = max(0, bottom - top)
    if w < 10 or h < 10:
        raise ValueError("La ventana no es visible o su tamaño es inválido.")
    return (left, top, w, h)

def screenshot_region(region) -> "Image.Image":
    try:
        return pyautogui.screenshot(region=region)
    except Exception as e:
        raise RuntimeError(f"No se pudo capturar la pantalla: {e}") from e

def save_jpg(pil_img: "Image.Image", out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        pil_img.save(out_path, format="JPEG", quality=95, subsampling=0, optimize=True)
    except PermissionError as e:
        raise PermissionError(f"Permiso denegado al guardar en: {out_path}") from e
    except FileNotFoundError as e:
        raise FileNotFoundError(f"Ruta no válida: {out_path}") from e
    except Exception as e:
        raise RuntimeError(f"No se pudo guardar la imagen: {e}") from e

# ------------------ EXTRACCIÓN DE SERIAL ------------------
# Endurece el patrón: seriales que comienzan con 'K' y longitud mínima 9
SERIAL_REGEXES = [
    r"\bK[A-Z0-9-]{8,}\b",     # 'K' + 8+ alfanuméricos/guiones
]

def regex_pick_serial(text: str) -> str | None:
    for pattern in SERIAL_REGEXES:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if m:
            return m.group(0).upper()
    return None

def serial_after_label(text: str, label="Carrete") -> str | None:
    """
    Busca explícitamente la etiqueta y toma el token que sigue:
    Carrete: <SERIAL>
    """
    # Línea con 'Carrete:' y captura token siguiente
    m = re.search(rf"{label}\s*:\s*([A-Za-z0-9\-]{{8,}})", text, re.IGNORECASE)
    if m:
        candidate = m.group(1).upper()
        # valida contra regex más estricto si aplica
        strong = regex_pick_serial(candidate)
        return strong or candidate
    return None

def serial_via_accessibility(window_title: str) -> str | None:
    """Lee el texto real del control Edit y extrae tras 'Carrete:'; fallback a regex."""
    if not PYWINAUTO_OK:
        return None
    try:
        app = Application(backend="uia").connect(title_re=rf".*{re.escape(window_title)}.*", timeout=5)
        dlg = app.window(title_re=rf".*{re.escape(window_title)}.*")
        dlg.set_focus()
        edits = dlg.descendants(control_type="Edit")
        for e in edits:
            try:
                wrapper = EditWrapper(e.element_info)
                txt = wrapper.get_value() or ""
                if not txt.strip():
                    continue
                # 1) prioriza 'Carrete:'
                s = serial_after_label(txt, "Carrete")
                if s:
                    return s
                # 2) fallback: regex global
                s = regex_pick_serial(txt)
                if s:
                    return s
            except Exception:
                continue
    except Exception:
        return None
    return None

def serial_via_ocr(pil_img: "Image.Image", tesseract_cmd: str | None) -> str | None:
    """OCR anclado a la etiqueta 'Carrete' y con lista blanca; sin diccionarios para evitar 'PEDIDOS'."""
    if not OCR_OK:
        return None
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    # Preprocesado: escala, gris, binarización Otsu
    img = np.array(pil_img)
    bgr = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    scale = 1.5
    h, w = bgr.shape[:2]
    bgr = cv2.resize(bgr, (int(w*scale), int(h*scale)), interpolation=cv2.INTER_LINEAR)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 9, 75, 75)
    _, thr = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Config OCR: psm 6, idiomas eng+spa, lista blanca y sin diccionarios
    config = (
        "--psm 6 -l eng+spa "
        "-c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- "
        "-c load_system_dawg=0 -c load_freq_dawg=0"
    )

    # 1) data por tokens para localizar 'Carrete'
    try:
        data = pytesseract.image_to_data(thr, config=config, output_type=pytesseract.Output.DICT)
        texts = data["text"]
        line_num = data["line_num"]
        idxs = [i for i, w in enumerate(texts) if w and re.fullmatch(r"(?i)carrete[:]?", w)]
        if idxs:
            i0 = idxs[0]
            ln = line_num[i0]
            # tokens posteriores en la misma línea primero; si no, siguientes 10 tokens
            same_line = [i for i in range(i0 + 1, len(texts)) if line_num[i] == ln]
            candidates = same_line or list(range(i0 + 1, min(i0 + 12, len(texts))))
            for i in candidates:
                tok = (texts[i] or "").strip().upper()
                if re.fullmatch(r"[A-Z0-9-]{9,}", tok):  # mínimo 9
                    strong = regex_pick_serial(tok)
                    if strong:
                        return strong
                    return tok
    except Exception:
        pass

    # 2) fallback: texto completo y extracción por etiqueta/regex
    try:
        raw = pytesseract.image_to_string(thr, config=config)
    except Exception:
        raw = pytesseract.image_to_string(pil_img, config=config)

    s = serial_after_label(raw, "Carrete")
    if s:
        return s
    return regex_pick_serial(raw)

def decide_output_path(outdir: Path, serial: str) -> Path:
    safe = re.sub(r"[^A-Za-z0-9._-]", "_", serial)
    return outdir.joinpath(f"{safe}.jpg")

# ------------------ MAIN ------------------
def main():
    ensure_windows()

    # Se mantienen argumentos para sobrescribir, pero con defaults internos
    parser = argparse.ArgumentParser(description="Captura ventana, extrae serial tras 'Carrete:' y guarda JPG con ese nombre.")
    parser.add_argument("--title", type=str, default=DEFAULT_WINDOW_TITLE, help="Título de la ventana (por defecto: config interna).")
    parser.add_argument("--outdir", type=str, default=DEFAULT_OUTPUT_DIR, help="Carpeta destino (por defecto: config interna).")
    parser.add_argument("--serial", type=str, default=None, help="Forzar serial manual.")
    parser.add_argument("--tess", type=str, default=DEFAULT_TESSERACT_PATH, help="Ruta a tesseract.exe si no está en PATH.")
    args = parser.parse_args()

    outdir = Path(args.outdir)
    try:
        outdir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"Error creando carpeta destino: {e}")
        sys.exit(5)

    win = find_window(args.title)
    if not win:
        print(f"Error: no se encontró la ventana '{args.title}'.")
        sys.exit(2)

    bring_to_front(win)
    try:
        region = get_window_box(win)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(3)

    time.sleep(0.25)  # estabilidad de dibujo
    pil_img = screenshot_region(region)

    # Extracción de serial: 1) forzado 2) accesibilidad 3) OCR
    serial = args.serial
    if not serial:
        serial = serial_via_accessibility(win.title)
    if not serial:
        serial = serial_via_ocr(pil_img, args.tess)

    if not serial:
        print("Advertencia: no se pudo extraer el serial automáticamente. Se usará timestamp.")
        serial = time.strftime("captura_%Y%m%d_%H%M%S")

    outfile = decide_output_path(outdir, serial)

    try:
        save_jpg(pil_img, outfile)
    except Exception as e:
        print(f"Error al guardar captura: {e}")
        sys.exit(4)

    print(f"Captura guardada en: {outfile}")
    print(f"Serial: {serial}")

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    main()
