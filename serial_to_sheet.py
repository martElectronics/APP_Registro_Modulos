#!/usr/bin/env python3
"""
serial_to_sheet.py — Captura el CSV por módulo del BMS Master (tecla 'g' por
Serial) y lo sube a una Google Sheet vía Apps Script Web App.

Flujo: abre el puerto del BMS -> envía 'g' -> captura el bloque entre
<<<CSV_BEGIN>>> y <<<CSV_END>>> -> POST a la URL del Apps Script. Opcionalmente
guarda una copia local en Data/ para el flujo de Excel (make_dataset.py).

Uso:
    python serial_to_sheet.py            # interactivo: ENTER = registrar
    python serial_to_sheet.py --once     # una captura y sale
    python serial_to_sheet.py --port COM5

Requisitos:  pip install pyserial requests   (o: pip install -r requirements_serial.txt)
"""

import sys
import os
import json
import time
import argparse
from datetime import datetime

import serial
import serial.tools.list_ports
import requests

# ===================== CONFIG (EDITA ESTO) =====================
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw6K9xdPUp5YF_3YiseZvpmxiVqwPHwwPWn-TprwhxczLF0G1efbIlt61nySNUcD4Hj/exec"
TOKEN           = "MART_cambia_esto_2026"   # DEBE coincidir con el TOKEN del Apps Script
SERIAL_PORT     = None        # None = autodetecta; o fija p.ej. "COM5"
BAUDRATE        = 115200
SAVE_LOCAL      = True         # guarda copia CRUDA en Data/ para el flujo xlsx (make_dataset.py)
DATA_DIR        = "Data"
CONFIG_JSON     = "Config.json"  # mapeo orden de lectura del BMS -> módulo físico real

BEGIN = "<<<CSV_BEGIN>>>"
END   = "<<<CSV_END>>>"
CAPTURE_TIMEOUT_S = 8.0        # margen para que el BMS lea V y T y vuelque el bloque
# ===============================================================


def find_port():
    """Devuelve el puerto: el fijado en CONFIG, o autodetecta (pregunta si hay varios)."""
    if SERIAL_PORT:
        return SERIAL_PORT
    ports = list(serial.tools.list_ports.comports())
    if not ports:
        sys.exit("[ERROR] No hay puertos serie. Conecta el BMS por USB.")
    if len(ports) == 1:
        return ports[0].device
    print("Puertos disponibles:")
    for i, p in enumerate(ports):
        print(f"  [{i}] {p.device}  {p.description}")
    sel = input(f"Elige puerto [0-{len(ports) - 1}] (ENTER=0): ").strip()
    return ports[int(sel) if sel.isdigit() else 0].device


def capture_csv(ser):
    """Envía 'g' y captura el bloque entre marcadores. Devuelve el CSV (str) o None."""
    ser.reset_input_buffer()
    ser.write(b"g")
    lines, capturing = [], False
    t0 = time.time()
    while time.time() - t0 < CAPTURE_TIMEOUT_S:
        raw = ser.readline().decode("utf-8", errors="ignore").strip()
        if not raw:
            continue
        if raw == BEGIN:
            capturing, lines = True, []
            continue
        if raw == END:
            return "\n".join(lines) if lines else None
        if capturing:
            lines.append(raw)
    return None  # timeout sin ver END


def upload(csv_text):
    """POST del CSV al Apps Script. Devuelve (status_code, body)."""
    r = requests.post(APPS_SCRIPT_URL,
                      data={"token": TOKEN, "data": csv_text},
                      timeout=20)
    return r.status_code, r.text.strip()


def save_local(csv_text):
    """Guarda una copia CRUDA en Data/ (orden de lectura) con nombre fechado.
    La lee make_dataset.py, que aplica su PROPIO remapeo con Config.json -> NO se
    remapea aquí (si no, doble remapeo)."""
    os.makedirs(DATA_DIR, exist_ok=True)
    fname = datetime.now().strftime("Registro_%d_%m_%Y-%H-%M.csv")
    path = os.path.join(DATA_DIR, fname)
    with open(path, "w", encoding="utf-8") as f:
        f.write(csv_text + "\n")
    return path


def load_module_map():
    """Lee Config.json['Modulos']: etiqueta del orden de lectura del BMS (M01..) ->
    nº de módulo FÍSICO real. Devuelve {} si no hay fichero/clave."""
    try:
        with open(CONFIG_JSON, "r", encoding="utf-8") as f:
            return json.load(f).get("Modulos", {}) or {}
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def remap_for_upload(csv_text, mapping):
    """Para la SUBIDA a Google: reetiqueta cada módulo del orden de lectura (M01..)
    a su nº físico real (Config.json) y ORDENA las filas por ese nº, para que cada
    módulo caiga siempre en la misma posición y se pueda seguir en el tiempo.
    Sin mapeo, usa el propio número de la etiqueta (identidad) y ordena igual."""
    lines = [l for l in csv_text.splitlines() if l.strip()]
    if not lines:
        return csv_text
    header, out = lines[0], []
    for line in lines[1:]:
        if line.startswith("Modulo"):
            continue
        cols = line.split(";")
        real = mapping.get(cols[0])                      # p.ej. "M02" -> 10
        if real is None:                                  # sin mapeo -> su propio nº
            digits = "".join(ch for ch in cols[0] if ch.isdigit())
            real = int(digits) if digits else 0
        cols[0] = f"M{int(real):02d}"
        out.append((int(real), ";".join(cols)))
    out.sort(key=lambda x: x[0])
    return "\n".join([header] + [row for _, row in out])


def one_capture(ser):
    print("-> Pidiendo CSV al BMS ('g')...")
    csv_text = capture_csv(ser)
    if not csv_text:
        print("[ERROR] No llego el bloque CSV. Revisa: BMS conectado, firmware con "
              "la tecla 'g', y el puerto/baudrate correctos.")
        return
    n = max(0, len(csv_text.splitlines()) - 1)   # menos la cabecera
    print(f"   Capturados {n} modulos.")
    if SAVE_LOCAL:
        print(f"   Copia local (cruda): {save_local(csv_text)}")

    # Para Google: traduce orden de lectura -> módulo físico real y ordena.
    mapping = load_module_map()
    if not mapping:
        print("   [aviso] sin Config.json['Modulos']: subo en orden de lectura "
              "(sin traducir a módulo real).")
    upload_csv = remap_for_upload(csv_text, mapping)

    try:
        code, body = upload(upload_csv)
    except requests.RequestException as e:
        print(f"[ERROR] No se pudo conectar al Apps Script: {e}")
        return
    if code == 200 and body.startswith("ok"):
        # El Apps Script devuelve "ok: <nombre de la hoja creada>".
        print(f"OK -> subido a la Google Sheet. {body}")
    else:
        print(f"[ERROR] Apps Script respondio {code}: {body}")
        if "bad token" in body:
            print("        -> el TOKEN de este script no coincide con el del Apps Script.")


def main():
    ap = argparse.ArgumentParser(description="BMS -> Google Sheet (Apps Script)")
    ap.add_argument("--once", action="store_true", help="una captura y salir")
    ap.add_argument("--port", help="puerto serie (p.ej. COM5)")
    args = ap.parse_args()

    global SERIAL_PORT
    if args.port:
        SERIAL_PORT = args.port

    port = find_port()
    print(f"Abriendo {port} @ {BAUDRATE}...")
    with serial.Serial(port, BAUDRATE, timeout=1) as ser:
        time.sleep(2)   # deja que el VCP (ST-Link) se estabilice
        if args.once:
            one_capture(ser)
            return
        print("Listo. Pulsa ENTER para registrar (Ctrl+C para salir).")
        try:
            while True:
                input()
                one_capture(ser)
        except KeyboardInterrupt:
            print("\nFin.")


if __name__ == "__main__":
    main()
