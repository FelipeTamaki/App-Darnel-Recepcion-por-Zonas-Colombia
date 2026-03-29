# 📦 Procesador de Pedidos Darnel

🔗 **[Abrir aplicación](https://app-darnel-recepcion-por-zonas-colombia.streamlit.app/)**

App web para cruzar notas de pedido Darnel con el catálogo interno de Pilarica y generar un reporte Excel agrupado por zona.

## ¿Qué hace?

1. Lee el archivo de pedido Darnel (`.xlsx`) — soporta pedidos con uno o múltiples bloques de productos
2. Cruza cada producto con el catálogo interno usando la hoja **ABM ART. DEPURADO** como puente entre el código Darnel y el código Pilarica
3. Obtiene la **zona** y la **cantidad por pallet** desde la hoja **RESUMEN**
4. Genera un Excel descargable agrupado por zona con totales

## Columnas del reporte generado

| Columna | Descripción |
|---|---|
| Zona | Zona de almacenamiento |
| Código Darnel | Código del proveedor (del pedido) |
| Código Pilarica | Código interno |
| Descripción | Nombre del artículo |
| Cant x Un Empaque | Cantidad que llega (del pedido) |
| Uds/Pallet | Unidades que entran por pallet |
| Pallets | Cant x Un Empaque ÷ Uds/Pallet |

## Inputs requeridos

| Archivo | Descripción |
|---|---|
| **Pedido Darnel** | Nota de pedido `.xlsx` (puede tener múltiples bloques de productos) |
| **Catálogo Pilarica** | Excel de stock/zonas con las hojas `ABM ART. DEPURADO 23.02.2025` y `RESUMEN` |

## Cómo funciona el match entre archivos

Los códigos del pedido Darnel y los códigos internos de Pilarica son distintos. El cruce se hace en dos pasos:
```
Pedido Darnel
  └── Código Darnel (columna "ID ARTICULO")
        └── ABM ART. DEPURADO (columna "Articulo Formularios")
              └── Código Pilarica (columna "Articulo")
                    └── RESUMEN → Zona + Cant por Pallet
```

## Archivos del repositorio
```
├── app.py              # Aplicación Streamlit
├── requirements.txt    # Dependencias
└── README.md
```

## Deploy en Streamlit Cloud

1. Hacer fork o clonar este repositorio en GitHub
2. Ir a [share.streamlit.io](https://share.streamlit.io)
3. Conectar la cuenta de GitHub y seleccionar el repositorio
4. Seleccionar `app.py` como archivo principal
5. Click en **Deploy**

## Correr localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Dependencias

- `streamlit`
- `pandas`
- `openpyxl`
