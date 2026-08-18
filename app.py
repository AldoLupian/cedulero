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

from procesador import procesar_pdfs

BASE_DIR = Path(__file__).resolve().parent
PLANTILLA_PATH = BASE_DIR / "plantilla" / "PLANTILLA CEDULA.xlsx"

app = Flask(__name__, static_folder=None)

# Cedulas generadas en esta sesion, guardadas en memoria (no en disco).
# Es una herramienta local de un solo usuario: se reinicia con el servidor.
ARCHIVOS = {}


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.post("/api/procesar")
def procesar():
    """Recibe todos los acuses de una tanda y devuelve una cedula por contribuyente."""
    if not PLANTILLA_PATH.exists():
        return jsonify({"error": f"No se encontro la plantilla en {PLANTILLA_PATH}"}), 500

    subidos = request.files.getlist("archivos")
    if not subidos:
        return jsonify({"error": "No se recibio ningun archivo."}), 400

    pdfs, errores = [], []
    for f in subidos:
        nombre = f.filename or "documento.pdf"
        if not nombre.lower().endswith(".pdf"):
            errores.append({"nombre": nombre, "estado": "error",
                            "avisos": ["Este archivo no es un PDF."]})
            continue
        pdfs.append((nombre, f.read()))

    cedulas, errores_lectura = procesar_pdfs(pdfs, PLANTILLA_PATH)
    errores.extend(errores_lectura)

    salida = []
    for c in cedulas:
        token = uuid.uuid4().hex
        ARCHIVOS[token] = {"nombre": c["nombre"] + ".xlsx", "bytes": c["xlsx_bytes"]}
        salida.append({
            "id": token,
            "nombre": c["nombre"],
            "rfc": c["rfc"],
            "razon_social": c["razon_social"],
            "archivos": c["archivos"],
            "estado": c["estado"],
            "avisos": c["avisos"],
            "banco": c["banco"],
            "fecha_pago": c["fecha_pago"],
            "importe_total": c["importe_total"],
            "filas": c["filas"],
            "descarga": f"/api/descargar/{token}",
        })

    return jsonify({"cedulas": salida, "errores": errores})


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
