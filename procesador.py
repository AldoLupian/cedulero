# -*- coding: utf-8 -*-
"""
Logica de extraccion y llenado de cedulas, adaptada de procesar.py para
trabajar en memoria (bytes) en vez de carpetas, para poder usarse desde
un servidor web.
"""

import re
import io
import unicodedata
import difflib
from datetime import datetime

import pdfplumber
import openpyxl

FILA_INICIAL = 39
FILA_FINAL = 63  # ultima fila disponible en la tabla de la plantilla

BANCOS_LISTA = [
    "ABACO CASA DE BOLSA, S.A. DE C.V.", "AMERICAN EXPRESS", "BANAMEX",
    "BANCO DEL ATLANTICO", "ACCIONES Y VALORES DE MEXICO, S.A. DE C.V.",
    "BANCO DEL EJERCITO, FUERZA AEREA Y ARMADA", "BANCO MERCANTIL DEL NORTE",
    "BANCO NACIONAL DE COMERCIO EXTERIOR", "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS",
    "CITIBANK", "SCOTIA BANK INVERLAT", "IXE BANCO", "BANCO DEL SURESTE, S.A.",
    "BANCO DEL CREDITO RURAL DEL ITSMO", "BANCRECER", "BANREGIO", "BANCA AFIRME, S.A.",
    "AGENCIAS DEL GOBIERNO DEL ESTADO / RECAUDADORA", "BANCO DEL BAJIO, S.A.",
    "BANCO INBURSA", "HSBC", "SANTANDER SERFIN", "BBVA BANCOMER", "BANCO AZTECA",
    "BANCA MIFEL", "AGENCIAS DEL GOBIERNO DEL ESTADO / PAGO EN BANCO", "BANSI, S.A.",
    "BANCO MULTIVA", "BANCO INTERACCIONES, S.A. DE C.V.",
    "CI BANCO, S.A. INSTITUCION DE BANCA MULTIPLE",
    "BANK OF TOKYO-MITSUBISHI, UFJ (MEXICO), S.A. INST. BANCA MULTIPLE",
    "BANCO MONEX, S.A. INSTITUCION DE BANCA MULTIPLE", "BANCO BASE S.A.",
    "BANCO VE POR MAS, S.A. INST. BANCA MULTIPLE", "INTERCAM GRUPO FINANCIERO",
    "MUFG BANK MEXICO, S.A. INSTIT BANCA MULTIPLE", "PAGO EN LA TESOFE",
    "BANCO BANCREA, S.A., INSTITUCION DE BANCA MULTIPLE", "BANCO ACTINVER, S.A. DE C.V.",
    "CBM BANCO, A.A INSTITUCION DE BANCA MULTIPLE", "BANCO AUTOFIN MEXICO, S.A. DE C.V.",
    "BANCO DEL BIENESTAR, S.A.", "KPTL MEXICO BANK, S.A. INSTITUCION DE BANCA MULTIPLE",
    "BANKAOOL, S.A. INSTITUCION DE BANCA MULTIPLE",
]

UMBRAL_CONFIANZA_BANCO = 0.40

LABEL_WORDS = {
    "institucin", "de", "crdito", "fecha", "del", "pago", "lnea", "medio",
    "presentacin", "importe", "no", "pagado", "operacin", "llave", "captura",
}


def normaliza(texto):
    return re.sub(r"[^a-zA-Z]", "", texto).lower()


def normaliza_ascii(texto):
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return texto.upper()


def a_numero(texto):
    if texto is None:
        return None
    limpio = texto.replace(",", "").replace("$", "").strip()
    try:
        return int(float(limpio))
    except ValueError:
        return None


def extraer_texto_completo(pdf):
    return "\n".join(p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages)


def extraer_conceptos(texto_completo):
    bloques = re.split(r"Concepto de pago\s*\d+:\s*", texto_completo)[1:]
    conceptos = []
    for bloque in bloques:
        bloque = re.split(r"Sello digital|SECCI[OÓ]N|INFORMACI[OÓ]N REGISTRADA", bloque)[0]
        nombre = bloque.strip().split("\n", 1)[0].strip()
        m_cargo = re.search(r"A cargo:\s*([\d,\.]+)", bloque)
        m_act = re.search(r"(?:Actualizaciones|Parte actualizada):\s*([\d,\.]+)", bloque, re.IGNORECASE)
        m_rec = re.search(r"Recargos:\s*([\d,\.]+)", bloque)
        conceptos.append({
            "concepto": nombre,
            "a_cargo": a_numero(m_cargo.group(1)) if m_cargo else 0,
            "actualizaciones": a_numero(m_act.group(1)) if m_act else None,
            "recargos": a_numero(m_rec.group(1)) if m_rec else None,
        })
    return conceptos


def extraer_info_pago(pdf):
    for page in pdf.pages:
        palabras = page.extract_words(x_tolerance=1, y_tolerance=1)
        if not any("RECIBIDO" in w["text"].upper() for w in palabras):
            continue

        top_inicio = next(w["top"] for w in palabras if "RECIBIDO" in w["text"].upper())
        top_fin = next((w["top"] for w in palabras if "digital" in w["text"].lower()), 10_000)

        es_ruido_firma = lambda t: len(t) > 25 and re.match(r"^[A-Za-z0-9+/=]+$", t)

        relevantes = [
            w for w in palabras
            if top_inicio + 5 < w["top"] < top_fin - 5 and not es_ruido_firma(w["text"])
        ]

        relevantes.sort(key=lambda w: w["top"])
        filas_agrupadas = []
        for w in relevantes:
            if filas_agrupadas and abs(w["top"] - filas_agrupadas[-1][0]["top"]) <= 4:
                filas_agrupadas[-1].append(w)
            else:
                filas_agrupadas.append([w])

        filas_valor = []
        for palabras_fila_sin_orden in filas_agrupadas:
            palabras_fila = sorted(palabras_fila_sin_orden, key=lambda w: w["x0"])
            normalizadas = [normaliza(w["text"]) for w in palabras_fila]
            no_vacias = [n for n in normalizadas if n]
            es_etiqueta = no_vacias and all(n in LABEL_WORDS for n in no_vacias)
            if es_etiqueta:
                continue
            izquierda = " ".join(w["text"] for w in palabras_fila if w["x0"] < 300).strip()
            derecha = " ".join(w["text"] for w in palabras_fila if w["x0"] >= 300).strip()
            filas_valor.append((izquierda, derecha))

        info = {
            "banco": None, "fecha_pago": None, "linea_captura": None,
            "no_operacion": None, "llave_pago": None,
        }
        if len(filas_valor) >= 1:
            info["banco"], info["fecha_pago"] = filas_valor[0]
        if len(filas_valor) >= 2:
            info["linea_captura"], _medio = filas_valor[1]
        if len(filas_valor) >= 3:
            _importe, info["no_operacion"] = filas_valor[2]
        if len(filas_valor) >= 4:
            _vacio, info["llave_pago"] = filas_valor[3]

        return info, len(filas_valor)

    return None, 0


PALABRAS_GENERICAS_BANCO = {
    "BANCO", "BANCA", "S", "A", "SA", "DE", "C", "V", "CV", "INSTITUCION",
    "INSTIT", "INST", "MULTIPLE", "GRUPO", "FINANCIERO", "NACIONAL",
}


def _normaliza_banco(nombre):
    limpio = re.sub(r"[^A-Z ]", " ", normaliza_ascii(nombre))
    palabras = [w for w in limpio.split() if w not in PALABRAS_GENERICAS_BANCO]
    return "".join(palabras)


def mapear_banco(nombre_pdf):
    if not nombre_pdf:
        return None, 0.0
    objetivo = _normaliza_banco(nombre_pdf)
    mejor_opcion, mejor_score = None, 0.0
    for opcion in BANCOS_LISTA:
        opcion_norm = _normaliza_banco(opcion)
        if not opcion_norm:
            continue
        sm = difflib.SequenceMatcher(None, objetivo, opcion_norm)
        match = sm.find_longest_match(0, len(objetivo), 0, len(opcion_norm))
        score = match.size / max(len(opcion_norm), 1)
        if score > mejor_score:
            mejor_opcion, mejor_score = opcion, score
    return mejor_opcion, mejor_score


def mapear_tipo_impuesto(concepto):
    c = normaliza_ascii(concepto)
    retencion = "RETEN" in c
    if "IEPS" in c:
        return "IEPS"
    if "IVA" in c or "VALOR AGREGADO" in c:
        return "IVA RET" if retencion else "IVA"
    if "ISR" in c or "SOBRE LA RENTA" in c:
        return "ISR RET" if retencion else "ISR"
    return "OTR RET" if retencion else "OTR"


def parsear_fecha(texto):
    if not texto:
        return None
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", texto)
    if not m:
        return None
    dia, mes, anio = m.groups()
    return datetime(int(anio), int(mes), int(dia))


def extraer_fecha_presentacion(texto_completo):
    m = re.search(
        r"Fecha y hora de presentaci[oó]n:\s*(\d{2})/(\d{2})/(\d{4})",
        texto_completo, re.IGNORECASE,
    )
    if not m:
        return None
    dia, mes, anio = m.groups()
    return datetime(int(anio), int(mes), int(dia))


def procesar_pdf_bytes(nombre_archivo, pdf_bytes, plantilla_path):
    """Procesa un PDF (bytes) y devuelve un dict con el resultado.

    No toca el disco salvo para leer la plantilla. El Excel generado
    se devuelve como bytes en 'xlsx_bytes' cuando el proceso tiene exito
    (aunque sea con advertencias).
    """
    avisos = []

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto_completo = extraer_texto_completo(pdf)
            conceptos = extraer_conceptos(texto_completo)
            info_pago, _n_filas_valor = extraer_info_pago(pdf)
            fecha_presentacion = extraer_fecha_presentacion(texto_completo)
    except Exception as e:
        return {
            "nombre": nombre_archivo,
            "estado": "error",
            "avisos": [f"No se pudo abrir el PDF: {e}"],
            "banco": None, "fecha_pago": None, "importe_total": None,
            "xlsx_bytes": None,
        }

    if not conceptos:
        return {
            "nombre": nombre_archivo,
            "estado": "error",
            "avisos": ["No se encontraron conceptos de pago en el PDF. "
                       "Verifica que sea un acuse de pago con linea de captura del SAT."],
            "banco": None, "fecha_pago": None, "importe_total": None,
            "xlsx_bytes": None,
        }

    if info_pago is None:
        avisos.append("No se encontro la seccion 'Informacion del pago recibido'. "
                       "No se pudieron llenar banco, fecha, no. de operacion ni llave de pago.")
        info_pago = {"banco": None, "fecha_pago": None, "linea_captura": None,
                     "no_operacion": None, "llave_pago": None}

    if len(conceptos) > (FILA_FINAL - FILA_INICIAL + 1):
        avisos.append(f"El PDF trae {len(conceptos)} conceptos pero la plantilla solo tiene "
                       f"{FILA_FINAL - FILA_INICIAL + 1} filas disponibles. Se recorto.")
        conceptos = conceptos[: FILA_FINAL - FILA_INICIAL + 1]

    banco_mapeado, score_banco = mapear_banco(info_pago["banco"])
    if info_pago["banco"] and score_banco < UMBRAL_CONFIANZA_BANCO:
        avisos.append(f"No se pudo identificar el banco con certeza (\"{info_pago['banco']}\"). "
                       "Se dejo el texto tal cual del PDF; revisa y selecciona manualmente en la lista del Excel.")
        banco_final = info_pago["banco"]
    else:
        banco_final = banco_mapeado

    fecha_pago = parsear_fecha(info_pago["fecha_pago"])
    no_operacion = a_numero(info_pago["no_operacion"]) or info_pago["no_operacion"]
    llave_pago = info_pago["llave_pago"]
    linea_captura = (info_pago["linea_captura"] or "").replace(" ", "")
    periodo_pago = fecha_presentacion.strftime("%d/%m/%y") if fecha_presentacion else None

    wb = openpyxl.load_workbook(plantilla_path)
    ws = wb.active

    importe_total = 0
    fila = FILA_INICIAL
    for c in conceptos:
        total_linea = (c["a_cargo"] or 0) + (c["actualizaciones"] or 0) + (c["recargos"] or 0)
        if total_linea == 0:
            fila += 1
            continue

        ws[f"M{fila}"] = no_operacion
        ws[f"S{fila}"] = llave_pago
        ws[f"Y{fila}"] = linea_captura
        ws[f"AI{fila}"] = banco_final
        ws[f"AN{fila}"] = mapear_tipo_impuesto(c["concepto"])
        ws[f"AQ{fila}"] = periodo_pago
        ws[f"AT{fila}"] = fecha_pago
        ws[f"AX{fila}"] = c["a_cargo"]
        importe_total += c["a_cargo"] or 0
        if c["actualizaciones"]:
            ws[f"BB{fila}"] = c["actualizaciones"]
            importe_total += c["actualizaciones"]
        if c["recargos"]:
            ws[f"BL{fila}"] = c["recargos"]
            importe_total += c["recargos"]
        fila += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return {
        "nombre": nombre_archivo,
        "estado": "advertencia" if avisos else "correcto",
        "avisos": avisos,
        "banco": banco_final,
        "fecha_pago": fecha_pago.strftime("%d/%m/%Y") if fecha_pago else None,
        "importe_total": importe_total,
        "xlsx_bytes": buffer.getvalue(),
    }
