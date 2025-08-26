# -*- coding: utf-8 -*-
"""
Captura la ventana indicada, extrae el número de serie del input 'carrete'
y guarda un JPG nombrado con ese serial.
Compatibilidad: Windows.

Ejemplos:
  python capture_window.py --title "Pedido.txt - Notepad" --outdir "C:/Users/garci/Desktop/capturar_pantalla/img"
  python capture_window.py --title "Pedido.txt - Notepad" --outdir "C:/Users/garci/Desktop/capturar_pantalla/img"  --tess "C:/Program Files/Tesseract-OCR/tesseract.exe"
  # Si conoces el serial y quieres forzarlo:
  python capture_window.py --title "Pedido.txt - Notepad" --serial "K900302392"

Dependencias: pyautogui, pygetwindow, pillow, pywinauto, pytesseract, opencv-python
"""

import argparse
import os
import re
import sys
import time
from pathlib import Path

# ---- Dependencias de captura/UI ----
try:
    import pyautogui
    import pygetwindow as gw
except Exception:
    print("Error: faltan dependencias. Instala con: pip install pyautogui pygetwindow pillow")
    sys.exit(1)

# pywinauto para extraer el valor del input 'carrete' si la app expone accesibilidad UIA
try:
    from pywinauto import Application
    from pywinauto.controls.uia_controls import EditWrapper
    from pywinauto.findwindows import ElementNotFoundError
    PYWINAUTO_OK = True
except Exception:
    PYWINAUTO_OK = False

# OCR de respaldo
try:
    import pytesseract
    import cv2
    from PIL import Image
    OCR_OK = True
except Exception:
    OCR_OK = False


def ensure_windows():
    if sys.platform != "win32":
        raise OSError("Este script está diseñado para Windows.")


def find_window(title_substr: str | None):
    if title_substr is None:
        return gw.getActiveWindow()
    all_windows = gw.getAllWindows()
    exact = [w for w in all_windows if w.title.strip() == title_substr.strip()]
    if exact:
        return exact[0]
    title_lower = title_substr.lower()
    partial = [w for w in all_windows if title_lower in w.title.lower()]
    return partial[0] if partial else None


def bring_to_front(win):
    try:
        win.activate(); time.sleep(0.2)
        if win.isMinimized:
            win.restore(); time.sleep(0.2)
        win.activate(); time.sleep(0.2)
    except Exception:
        pass


def get_window_box(win):
    left, top, right, bottom = win.left, win.top, win.right, win.bottom
    w = max(0, right - left)
    h = max(0, bottom - top)
    if w < 10 or h < 10:
        raise ValueError("La ventana no es visible o tiene tamaño inválido.")
    return (left, top, w, h)


def screenshot_region(region):
    try:
        img = pyautogui.screenshot(region=region)
        return img
    except Exception as e:
        raise RuntimeError(f"No se pudo capturar la pantalla: {e}") from e


def save_jpg(pil_img: Image.Image, outfile: Path):
    outfile.parent.mkdir(parents=True, exist_ok=True)
    try:
        pil_img.save(outfile, format="JPEG", quality=95, subsampling=0, optimize=True)
    except PermissionError as e:
        raise PermissionError(f"Permiso denegado al guardar en: {outfile}") from e
    except FileNotFoundError as e:
        raise FileNotFoundError(f"Ruta no válida: {outfile}") from e
    except Exception as e:
        raise RuntimeError(f"No se pudo guardar la imagen: {e}") from e


# ---------- Extracción de serial ----------
SERIAL_REGEXES = [
    r"\b[Kk][A-Za-z0-9]{5,20}\b",     # Ej: K900302392
    r"\b[A-Za-z][A-Za-z0-9-]{5,20}\b" # fallback más laxo
]

def regex_pick_serial(text: str) -> str | None:
    # Prioriza el primero que haga match.
    for pattern in SERIAL_REGEXES:
        m = re.search(pattern, text)
        if m:
            return m.group(0).upper()
    return None


def serial_via_accessibility(window_title: str, hint_label: str = "carrete") -> str | None:
    """Intenta leer el valor del control Edit llamado 'carrete' usando UIA."""
    if not PYWINAUTO_OK:
        return None
    try:
        app = Application(backend="uia").connect(title_re=rf".*{re.escape(window_title)}.*", timeout=5)
        dlg = app.window(title_re=rf".*{re.escape(window_title)}.*")
        dlg.set_focus()
        # Busca Edits con nombre que contenga 'carrete'
        edits = dlg.descendants(control_type="Edit")
        # Intento 1: por name_accessible
        for e in edits:
            try:
                name = e.element_info.name or ""
                if hint_label.lower() in name.lower():
                    wrapper = EditWrapper(e.element_info)
                    val = wrapper.get_value()
                    if val:
                        serial = regex_pick_serial(val) or val.strip()
                        if serial:
                            return serial.upper()
            except Exception:
                continue
        # Intento 2: por label cercano (Label seguido de Edit)
        # Busca textos estáticos para 'carrete' y toma el siguiente Edit
        texts = dlg.descendants(control_type="Text")
        targets = [t for t in texts if (t.element_info.name or "").strip().lower().__contains__(hint_label)]
        if targets:
            # toma cualquier Edit del diálogo y escoge el más cercano
            # enfoque simple: primero Edit del contenedor
            for e in edits:
                try:
                    wrapper = EditWrapper(e.element_info)
                    val = wrapper.get_value()
                    if val:
                        serial = regex_pick_serial(val) or val.strip()
                        if serial:
                            return serial.upper()
                except Exception:
                    continue
    except (ElementNotFoundError, TimeoutError, Exception):
        return None
    return None


def serial_via_ocr(pil_img: Image.Image, tesseract_cmd: str | None) -> str | None:
    """OCR general sobre la ventana completa. Opcionalmente usa una pre-segmentación con OpenCV."""
    if not OCR_OK:
        return None
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    # Convertir PIL -> OpenCV BGR
    try:
        import numpy as np
        cv_img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    except Exception:
        cv_img = None

    candidate_text = ""
    try:
        if cv_img is not None:
            # Preprocesado: escala + gris + binarización suave
            scale = 1.5
            h, w = cv_img.shape[:2]
            cv_img = cv2.resize(cv_img, (int(w*scale), int(h*scale)), interpolation=cv2.INTER_LINEAR)
            gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
            gray = cv2.bilateralFilter(gray, 9, 75, 75)
            _, thr = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            config = "--psm 6"
            candidate_text = pytesseract.image_to_string(thr, config=config)
        else:
            candidate_text = pytesseract.image_to_string(pil_img, config="--psm 6")
    except Exception:
        return None

    if not candidate_text:
        return None

    # Si encontramos la palabra 'carrete', intenta tomar el token que sigue
    lines = [l.strip() for l in candidate_text.splitlines() if l.strip()]
    joined = " ".join(lines)
    serial = None

    # 1) busca patrón cercano a 'carrete'
    m = re.search(r"(carrete|carreté|carrete:)\s*([A-Za-z0-9\-]{5,20})", joined, re.IGNORECASE)
    if m:
        serial = m.group(2).upper()

    # 2) si no, usa regex general
    if not serial:
        serial = regex_pick_serial(joined)

    return serial


def decide_output_path(outdir: Path, serial: str) -> Path:
    safe = re.sub(r"[^A-Za-z0-9._-]", "_", serial)
    return outdir.joinpath(f"{safe}.jpg")


def main():
    ensure_windows()

    parser = argparse.ArgumentParser(description="Captura un JPG de una ventana y nombra el archivo con el serial del 'carrete'.")
    parser.add_argument("--title", type=str, default=None, help="Título exacto o parcial de la ventana. Si se omite, usa la activa.")
    parser.add_argument("--outdir", type=str, required=True, help="Carpeta destino del JPG.")
    parser.add_argument("--serial", type=str, default=None, help="Serial manual. Si se pasa, se omite extracción.")
    parser.add_argument("--tess", type=str, default=None, help="Ruta a tesseract.exe si no está en PATH.")
    args = parser.parse_args()

    outdir = Path(args.outdir)
    if not outdir.exists():
        try:
            outdir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Error creando carpeta destino: {e}")
            sys.exit(5)

    # 1) localizar ventana
    win = find_window(args.title)
    if win is None:
        t = args.title or "(ventana activa)"
        print(f"Error: no se encontró la ventana: {t}")
        sys.exit(2)

    bring_to_front(win)
    try:
        region = get_window_box(win)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(3)

    # Pequeña espera para estabilidad de dibujo
    time.sleep(0.2)

    # 2) captura imagen de la ventana
    pil_img = screenshot_region(region)

    # 3) obtener serial
    serial = args.serial
    if not serial:
        # Intento 1: accesibilidad
        serial = serial_via_accessibility(win.title, hint_label="carrete")
    if not serial:
        # Intento 2: OCR
        serial = serial_via_ocr(pil_img, args.tess)

    if not serial:
        print("Advertencia: no se pudo extraer el serial automáticamente.")
        print("Sugerencias:")
        print("- Verifica que el input se llame 'carrete' o esté visible.")
        print("- Ejecuta con --serial para forzar el nombre.")
        print("- Si usas OCR, instala Tesseract y pasa --tess con su ruta.")
        # Aún así guardar con timestamp para no perder la captura
        timestamp_name = time.strftime("captura_%Y%m%d_%H%M%S")
        outfile = outdir.joinpath(f"{timestamp_name}.jpg")
    else:
        outfile = decide_output_path(outdir, serial)

    # 4) guardar JPG
    try:
        save_jpg(pil_img, outfile)
    except Exception as e:
        print(f"Error al guardar captura: {e}")
        sys.exit(4)

    if serial:
        print(f"Captura guardada: {outfile}")
        print(f"Serial detectado: {serial}")
    else:
        print(f"Captura guardada sin serial: {outfile}")

if __name__ == "__main__":
    # Nota DPI: si el recorte se desalineara por escalado, ajusta la escala de pantalla de Windows.
    pyautogui.FAILSAFE = True
    main()
