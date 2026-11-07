# Finanzas del consultorio

Planilla para llevar los ingresos y gastos de un consultorio de cosmetología + nutrición y saber cuánta ganancia queda cada mes.

## Cómo abrirla en Google Sheets

1. Descargá `Finanzas-Consultorio.xlsx`.
2. Subilo a Google Drive → clic derecho → **Abrir con → Hojas de cálculo de Google**.
3. En Sheets: **Archivo → Guardar como Hojas de cálculo de Google** (así queda como planilla nativa y se puede usar desde el celular).

## Hojas

| Hoja | Para qué |
|---|---|
| Cómo usar | Instrucciones cortas |
| Panel | Elegís mes y año: cobrado, comisiones, gastos, **ganancia**, margen, meta, por servicio, por categoría y el año mes a mes con gráfico |
| Sesiones | Una fila por cada sesión cobrada |
| Gastos | Una fila por cada gasto del trabajo |
| Servicios | Catálogo con precio, duración y costo de insumos → ganancia por sesión y por hora |
| Listas | Categorías de gasto (fijo/variable), medios de pago con su % de comisión, rubros |

Cremita = se completa · Gris = se calcula solo. Los precios, comisiones y la meta son de ejemplo.

## Regenerar

`python3 generar_planilla.py` (necesita `openpyxl`).
