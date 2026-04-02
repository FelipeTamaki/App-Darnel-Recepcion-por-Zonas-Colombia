# 📦 Recepción por Zonas

🔗 **[Abrir aplicación](https://app-darnel-recepcion-por-zonas-colombia.streamlit.app/)**

App web con dos herramientas para gestionar la recepción de mercadería agrupada por zona de almacenamiento. Al ingresar, podés elegir cuál usar.

---

## 🗂️ Funcionalidades

### 🌎 1 — Pedidos Darnel (Colombia)

Cruzá un pedido Darnel con el catálogo interno de Pilarica y generá un reporte Excel agrupado por zona.

**Cómo funciona:**

1. Subís el pedido Darnel (`.xlsx`) — soporta uno o múltiples bloques de productos
2. Subís el catálogo Pilarica (`.xlsx`) con las hojas `ABM ART. DEPURADO` y `RESUMEN`
3. La app cruza cada producto en dos pasos:

```
Pedido Darnel
  └── Código Darnel  →  ABM ART. DEPURADO (Articulo Formularios)
                              └── Código Pilarica  →  RESUMEN
                                                          └── Zona + Cant por Pallet
```

4. Descargás el Excel con el reporte agrupado por zona

**Columnas del reporte:**

| Columna | Descripción |
|---|---|
| Zona | Zona de almacenamiento |
| Código Darnel | Código del proveedor |
| Código Pilarica | Código interno |
| Descripción | Nombre del artículo |
| Cant x Un Empaque | Cantidad recibida |
| Uds/Pallet | Unidades por pallet |
| Pallets | Cant x Un Empaque ÷ Uds/Pallet |

---

### 🗂️ 2 — Notas de Pedido ZAPLAST

Procesá una o varias Notas de Pedido en PDF y generá el reporte por zona automáticamente. El Masterdata ya está integrado en la app, no hace falta subir nada más.

**Cómo funciona:**

1. Subís uno o más PDFs de Notas de Pedido ZAPLAST (formato estándar)
2. La app extrae de cada PDF: cliente, número de pedido, código de artículo, descripción y cantidad
3. Cruza cada artículo con el `Masterdata.xlsx` (embebido en el repo) por código — con fallback por nombre si el código no matchea
4. Obtiene la **zona** y la **cantidad por pallet** del Masterdata
5. Descargás el Excel con el reporte agrupado por zona

**Columnas del reporte:**

| Columna | Descripción |
|---|---|
| Zona | Zona de almacenamiento |
| Código Artículo | Código ZAPLAST |
| Descripción | Nombre del artículo |
| Cliente | Razón social del cliente |
| Cantidad | Unidades pedidas |
| Uds/Pallet | Unidades por pallet (del Masterdata) |
| Pallets | Cantidad ÷ Uds/Pallet |

---

## 📁 Archivos del repositorio

```
├── app.py               # Aplicación Streamlit (hub con ambas funcionalidades)
├── Masterdata.xlsx      # Masterdata ZAPLAST embebido (usado por la func. 2)
├── requirements.txt     # Dependencias
└── README.md
```

---

## 🚀 Deploy en Streamlit Cloud

1. Clonar o forkear este repositorio
2. Ir a [share.streamlit.io](https://share.streamlit.io)
3. Conectar la cuenta de GitHub y seleccionar el repositorio
4. Seleccionar `app.py` como archivo principal
5. Click en **Deploy**

## 💻 Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📦 Dependencias

```
streamlit
pandas
openpyxl
pdfplumber
```
