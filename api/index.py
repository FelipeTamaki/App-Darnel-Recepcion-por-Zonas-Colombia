import base64
from pathlib import Path

from flask import Flask, jsonify, redirect, request

from shared_logic import (
    build_darnel_preview,
    build_zaplast_preview,
    darnel_cruzar,
    darnel_generar_excel,
    darnel_leer_catalogo,
    darnel_leer_pedido,
    zaplast_generar_excel,
    zaplast_load_masterdata,
    zaplast_procesar,
)

BASE_DIR = Path(__file__).resolve().parent.parent
app = Flask(__name__)


def _json_error(message: str, status: int = 400):
    return jsonify({"ok": False, "error": message}), status


@app.get("/")
def index():
    return redirect("/index.html", code=307)


@app.get("/api/health")
def health():
    return jsonify({"ok": True})


@app.post("/api/process/darnel")
def process_darnel():
    pedido_file = request.files.get("pedido")
    catalogo_file = request.files.get("catalogo")
    if not pedido_file or not catalogo_file:
        return _json_error("Subí el pedido Darnel y el catálogo Pilarica.")

    try:
        pedido = darnel_leer_pedido(pedido_file)
        catalogo_file.stream.seek(0)
        mapa, catalogo = darnel_leer_catalogo(catalogo_file)
        resultado, sin_match = darnel_cruzar(pedido, mapa, catalogo)
        if not resultado:
            return _json_error("No se encontraron productos válidos para generar el reporte.", 422)

        nombre_pedido = Path(pedido_file.filename or "pedido").stem
        workbook_bytes = darnel_generar_excel(resultado, nombre_pedido)
        preview = build_darnel_preview(resultado, nombre_pedido)
        return jsonify(
            {
                "ok": True,
                "type": "darnel",
                "downloadName": f"reporte_por_zona_{nombre_pedido}.xlsx",
                "workbookBase64": base64.b64encode(workbook_bytes).decode("ascii"),
                "warnings": sin_match,
                "preview": preview,
            }
        )
    except Exception as exc:
        return _json_error(str(exc), 500)


@app.post("/api/process/zaplast")
def process_zaplast():
    pdf_files = request.files.getlist("pdfs")
    if not pdf_files:
        return _json_error("Subí al menos un PDF de Nota de Pedido.")

    try:
        masterdata = zaplast_load_masterdata()
        payload_files = [(file.filename or "archivo.pdf", file.read()) for file in pdf_files]
        df, sin_match = zaplast_procesar(payload_files, masterdata)
        if df.empty:
            return _json_error("No se encontraron artículos en los PDFs subidos.", 422)

        workbook_bytes = zaplast_generar_excel(df)
        preview = build_zaplast_preview(df)
        pedidos_label = "_".join(sorted(df["Nro Pedido"].unique()))
        return jsonify(
            {
                "ok": True,
                "type": "zaplast",
                "downloadName": f"pedidos_por_zona_{pedidos_label}.xlsx",
                "workbookBase64": base64.b64encode(workbook_bytes).decode("ascii"),
                "warnings": sin_match,
                "preview": preview,
            }
        )
    except Exception as exc:
        return _json_error(str(exc), 500)
