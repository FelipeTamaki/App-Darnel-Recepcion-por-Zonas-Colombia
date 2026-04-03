# Recepción por Zonas | Darnel

Aplicación web para generar reportes Excel de recepción por zona con dos flujos:

1. `Pedidos Darnel`
   Cruza un pedido Darnel con el catálogo Pilarica y arma el reporte agrupado por zona.
2. `Notas de Pedido ZAPLAST`
   Procesa uno o varios PDFs y genera el reporte usando `Masterdata.xlsx`.

La lógica del Excel sigue en Python, pero ahora la interfaz está preparada para Vercel:

- frontend estático en `public/`
- API Python en `api/index.py`
- preview web por pestañas que replica el contenido del workbook final
- descarga directa del `.xlsx`

## Estructura

```text
├── api/
│   └── index.py
├── public/
│   ├── app.js
│   ├── styles.css
│   └── assets/
│       └── darnel-logo.jpg
├── shared_logic.py
├── Masterdata.xlsx
├── vercel.json
├── requirements.txt
└── app_1.py
```

`app_1.py` queda como referencia de la versión anterior en Streamlit. El flujo nuevo recomendado es el de Vercel.

## Deploy en Vercel

1. Subí el repo a GitHub.
2. En Vercel, importá el repositorio.
3. No hace falta `build command`.
4. Vercel detecta:
   - archivos estáticos en `public/`
   - función Python en `api/index.py`
5. Deploy.

El archivo `vercel.json` ya deja configuradas las rutas para que:

- `/` sirva el frontend
- `/api/process/darnel` procese pedidos Darnel
- `/api/process/zaplast` procese PDFs ZAPLAST

## Desarrollo local

```bash
pip install -r requirements.txt
flask --app api/index.py run
```

Después abrí [http://127.0.0.1:5000](http://127.0.0.1:5000).

## Dependencias

```text
Flask
pandas
openpyxl
pdfplumber
```
