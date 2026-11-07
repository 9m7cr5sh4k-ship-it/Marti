# Finanzas del consultorio

Planilla para llevar los ingresos y gastos de un consultorio de cosmetología + nutrición y saber cuánta ganancia queda cada mes.

## Cómo abrirla en Google Sheets

1. Descargá `Finanzas-Consultorio.xlsx`.
2. Subilo a Google Drive → clic derecho → **Abrir con → Hojas de cálculo de Google**.
3. En Sheets: **Archivo → Guardar como Hojas de cálculo de Google** (así queda como planilla nativa y se puede usar desde el celular).

## Hojas

| Hoja | Para qué |
|---|---|
| Resumen | Período a elegir: ingresos, gastos, ganancia neta, margen, sesiones, ticket promedio, comisiones y meta (con comparación contra el mes anterior), detalle por servicio, por categoría, por área y evolución anual con gráfico |
| Ingresos | Una fila por cada sesión cobrada |
| Gastos | Una fila por cada gasto del consultorio |
| Servicios | Precio, duración y costo de insumos → margen por sesión y por hora |
| Configuración | Nombre, meta mensual, medios de pago con su comisión, categorías de gasto, áreas |
| Guía | Instrucciones de uso |

Las columnas con fondo gris se calculan solas. Los precios de los servicios y las comisiones son de referencia.

## Regenerar

`python3 generar_planilla.py` (necesita `openpyxl`).
