# -*- coding: utf-8 -*-
"""
Servidor de Cedulero: sirve la interfaz web y expone endpoints para
subir acuses de pago (PDF) y descargar las cedulas (Excel) generadas.

Uso local:
    python app.py
Luego abre http://127.0.0.1:5000 en el navegador.

En produccion (Render u otro host) se ejecuta con gunicorn, ver Procfile.
"""

import io
import os
import uuid
import zipfile
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory

from procesador import procesar_pdf_bytes

BASE_DIR = Path(__file__).resolve().parent
PLANTILLA_PATH = BASE_DIR / "plantilla" / "PLANTILLA CEDULA.xlsx"

app = Flask(__name__, static_folder=None)

# Cedulas generadas en esta sesion, guardadas en memoria (no en disco).
# Es una herramienta local de un solo usuario: se reinicia con el servidor.
ARCHIVOS = {}


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.post("/api/procesar-uno")
def procesar_uno():
    if not PLANTILLA_PATH.exists():
        return jsonify({"error": f"No se encontro la plantilla en {PLANTILLA_PATH}"}), 500

    f = request.files.get("archivo")
    if not f:
        return jsonify({"error": "No se recibio ningun archivo."}), 400

    nombre = f.filename or "documento.pdf"
    if not nombre.lower().endswith(".pdf"):
        return jsonify({
            "nombre": nombre, "estado": "error",
            "avisos": ["Este archivo no es un PDF."],
            "banco": None, "fecha_pago": None, "importe_total": None,
            "descarga": None,
        })

    resultado = procesar_pdf_bytes(nombre, f.read(), PLANTILLA_PATH)

    token = None
    if resultado["xlsx_bytes"] is not None:
        token = uuid.uuid4().hex
        ARCHIVOS[token] = {
            "nombre": Path(nombre).stem + ".xlsx",
            "bytes": resultado["xlsx_bytes"],
        }

    return jsonify({
        "id": token,
        "nombre": resultado["nombre"],
        "estado": resultado["estado"],
        "avisos": resultado["avisos"],
        "banco": resultado["banco"],
        "fecha_pago": resultado["fecha_pago"],
        "importe_total": resultado["importe_total"],
        "descarga": f"/api/descargar/{token}" if token else None,
    })


@app.get("/api/descargar/<token>")
def descargar(token):
    archivo = ARCHIVOS.get(token)
    if not archivo:
        return jsonify({"error": "Ese archivo ya no esta disponible. Vuelve a procesar el PDF."}), 404
    return send_file(
        io.BytesIO(archivo["bytes"]),
        as_attachment=True,
        download_name=archivo["nombre"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/descargar-todo")
def descargar_todo():
    tokens = [t for t in request.args.get("tokens", "").split(",") if t]
    encontrados = {t: ARCHIVOS[t] for t in tokens if t in ARCHIVOS}
    if not encontrados:
        return jsonify({"error": "No hay cedulas disponibles para descargar."}), 404

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for archivo in encontrados.values():
            zf.writestr(archivo["nombre"], archivo["bytes"])
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="cedulas.zip",
        mimetype="application/zip",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
