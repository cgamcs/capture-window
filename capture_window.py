# -*- coding: utf-8 -*-
"""
Captura una imagen JPG de la ventana activa o de una ventana cuyo título se especifique.
Compatibilidad: Windows.
Dependencias: pyautogui, pygetwindow, Pillow (PIL).
Uso:
  python capture_window.py --title "Calculadora gráfica - GeoGebra - Brave" --out "C:/Users/garci/Desktop/capturar_pantalla/img/captura.jpg"
  # O sin --title para usar la ventana activa
"""

import argparse
import os
import sys
import time
from pathlib import Path

# Librerías de captura y ventanas
try:
    import pyautogui
    import pygetwindow as gw
except Exception as e:
    print("Error: faltan dependencias. Instala con: pip install pyautogui pygetwindow pillow")
    sys.exit(1)

from PIL import Image  # noqa: F401  # Solo para asegurar presencia de Pillow

def ensure_windows():
    """Verifica que el SO sea Windows."""
    if sys.platform != "win32":
        raise OSError("Este script está diseñado para Windows.")

def find_window(title_substr: str | None):
    """
    Encuentra una ventana:
    - Si title_substr es None, retorna la ventana activa.
    - Si se proporciona, intenta coincidencia exacta; si no, por substring.
    Retorna un objeto Window de pygetwindow o None.
    """
    if title_substr is None:
        return gw.getActiveWindow()

    # Buscar coincidencia exacta primero
    all_windows = gw.getAllWindows()
    exact = [w for w in all_windows if w.title.strip() == title_substr.strip()]
    if exact:
        return exact[0]

    # Buscar por substring (insensible a mayúsculas)
    title_lower = title_substr.lower()
    partial = [w for w in all_windows if title_lower in w.title.lower()]
    return partial[0] if partial else None

def bring_to_front(win):
    """
    Intenta activar y llevar la ventana al frente.
    Algunos programas pueden bloquear el foco; se intenta dos veces.
    """
    try:
        win.activate()
        time.sleep(0.3)
        # Si sigue minimizada, restaurar
        if win.isMinimized:
            win.restore()
            time.sleep(0.3)
        # A veces requiere un segundo intento
        win.activate()
        time.sleep(0.3)
    except Exception:
        # Continuar incluso si no se puede activar
        pass

def get_window_box(win):
    """
    Devuelve la caja (left, top, width, height) de la ventana.
    Valida que sea visible y con tamaño válido.
    """
    left, top, right, bottom = win.left, win.top, win.right, win.bottom
    width = max(0, right - left)
    height = max(0, bottom - top)
    if width < 10 or height < 10:
        raise ValueError("La ventana tiene dimensiones no válidas o no es visible.")
    return (left, top, width, height)

def save_screenshot(region, out_path: Path):
    """
    Toma y guarda la captura de una región específica de pantalla.
    Maneja errores de permisos y rutas.
    """
    # Crear directorio destino si no existe
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Capturar
    try:
        img = pyautogui.screenshot(region=region)  # region=(left, top, width, height)
    except Exception as e:
        raise RuntimeError(f"No se pudo capturar la pantalla: {e}") from e

    # Guardar como JPG con calidad alta
    try:
        img.save(out_path, format="JPEG", quality=95, subsampling=0, optimize=True)
    except PermissionError as e:
        raise PermissionError(f"Permiso denegado al guardar en: {out_path}. "
                              f"Elige otra carpeta o ejecuta con permisos elevados.") from e
    except FileNotFoundError as e:
        raise FileNotFoundError(f"Ruta no válida: {out_path}") from e
    except Exception as e:
        raise RuntimeError(f"No se pudo guardar la imagen: {e}") from e

def main():
    ensure_windows()

    parser = argparse.ArgumentParser(description="Captura un JPG de una ventana específica en Windows.")
    parser.add_argument("--title", type=str, default=None,
                        help="Título exacto o parcial de la ventana. Si se omite, usa la ventana activa.")
    parser.add_argument("--out", type=str, required=True,
                        help="Ruta de salida del JPG. Ejemplo: C:/Users/TuUsuario/Pictures/captura.jpg")
    args = parser.parse_args()

    out_path = Path(args.out)
    if out_path.suffix.lower() not in (".jpg", ".jpeg"):
        # Forzar extensión .jpg si el usuario no la puso
        out_path = out_path.with_suffix(".jpg")

    # Buscar ventana
    win = find_window(args.title)
    if win is None:
        t = args.title or "(ventana activa)"
        print(f"Error: no se encontró la ventana: {t}")
        # Sugerencia mínima
        print("Verifica que el programa esté abierto y visible. Puedes pasar parte del nombre con --title.")
        sys.exit(2)

    # Llevar al frente y calcular región
    bring_to_front(win)

    try:
        region = get_window_box(win)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(3)

    # Pequeña espera para que Windows termine de dibujar
    time.sleep(0.2)

    # Capturar y guardar
    try:
        save_screenshot(region, out_path)
    except Exception as e:
        print(f"Error al guardar captura: {e}")
        sys.exit(4)

    print(f"Captura guardada en: {out_path}")

if __name__ == "__main__":
    main()