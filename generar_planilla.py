"""Genera la planilla de finanzas del consultorio (cosmetología + nutrición).

Uso:  python3 generar_planilla.py   ->  Finanzas-Consultorio.xlsx
Después se sube a Google Drive y se abre con Google Sheets.
"""
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date

SALIDA = "Finanzas-Consultorio.xlsx"
FILAS = 1000  # filas preparadas para cargar en Sesiones y Gastos
ANIO = 2026

# Paleta: verde salvia + rosa empolvado
OSCURO = "3E5C58"
ROSA = "F3D6D8"
ROSA_FUERTE = "C9737B"
CREMA = "FFF9EC"   # celdas para completar
GRIS = "F1F1F1"    # celdas automáticas
FUENTE = "Arial"

PESOS = '"$"#,##0;-"$"#,##0;"-"'
PORC = '0%;-0%;"-"'
FECHA = "dd/mm/yyyy"

f_titulo = Font(name=FUENTE, size=18, bold=True, color=OSCURO)
f_sub = Font(name=FUENTE, size=10, italic=True, color="666666")
f_head = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
f_norm = Font(name=FUENTE, size=10)
f_bold = Font(name=FUENTE, size=10, bold=True)
f_auto = Font(name=FUENTE, size=10, color="555555")
fill_head = PatternFill("solid", fgColor=OSCURO)
fill_head_auto = PatternFill("solid", fgColor="7A9490")
fill_input = PatternFill("solid", fgColor=CREMA)
fill_auto = PatternFill("solid", fgColor=GRIS)
fill_rosa = PatternFill("solid", fgColor=ROSA)
linea = Side(style="thin", color="DDDDDD")
borde = Border(bottom=linea)

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
RUBROS = ["Cosmetología", "Nutrición", "Otro"]
CATEGORIAS = [
    ("Insumos y productos", "Variable"), ("Descartables", "Variable"),
    ("Alquiler del consultorio", "Fijo"), ("Luz, agua, gas e internet", "Fijo"),
    ("Celular", "Fijo"), ("Monotributo e impuestos", "Fijo"),
    ("Matrícula y seguro", "Fijo"), ("Software y apps", "Fijo"),
    ("Aparatología y equipos", "Variable"), ("Mantenimiento de equipos", "Variable"),
    ("Capacitación y cursos", "Variable"), ("Marketing y redes", "Variable"),
    ("Transporte", "Variable"), ("Lavandería y limpieza", "Variable"),
    ("Comisiones bancarias", "Variable"), ("Otros", "Variable"),
]
MEDIOS = [("Efectivo", 0), ("Transferencia", 0), ("Mercado Pago", 0.03),
          ("Tarjeta de débito", 0.02), ("Tarjeta de crédito", 0.05)]
SERVICIOS = [  # servicio, rubro, precio, minutos, insumos por sesión
    ("Limpieza facial profunda", "Cosmetología", 30000, 60, 4000),
    ("Peeling químico", "Cosmetología", 40000, 45, 7000),
    ("Dermaplaning", "Cosmetología", 32000, 45, 3000),
    ("Hidratación facial", "Cosmetología", 25000, 50, 3500),
    ("Radiofrecuencia facial", "Cosmetología", 35000, 40, 1500),
    ("Microneedling", "Cosmetología", 45000, 60, 8000),
    ("Drenaje linfático", "Cosmetología", 28000, 60, 1000),
    ("Consulta nutricional inicial", "Nutrición", 35000, 60, 0),
    ("Control nutricional", "Nutrición", 25000, 30, 0),
]

wb = Workbook()


def encabezado(ws, fila, textos, autos=()):
    for i, t in enumerate(textos, start=1):
        c = ws.cell(row=fila, column=i, value=t)
        c.font = f_head
        c.fill = fill_head_auto if i in autos else fill_head
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 30


def anchos(ws, valores):
    for letra, w in valores.items():
        ws.column_dimensions[letra].width = w


def lista(ws, rango_celdas, formula, msg=None):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error = "Elegí una opción de la lista (o agregala en la hoja correspondiente)."
    dv.errorStyle = "warning"
    if msg:
        dv.prompt = msg
    ws.add_data_validation(dv)
    dv.add(rango_celdas)


# ------------------------------------------------------------------ Hojas (orden)
ayuda = wb.active
ayuda.title = "Cómo usar"
panel = wb.create_sheet("Panel")
ses = wb.create_sheet("Sesiones")
gas = wb.create_sheet("Gastos")
srv = wb.create_sheet("Servicios")
lis = wb.create_sheet("Listas")

# ------------------------------------------------------------------ Listas
encabezado(lis, 1, ["Categoría de gasto", "Tipo", "", "Medio de pago", "Comisión %", "", "", "Meses", "", "Rubros"])
for c in ("C", "F", "G", "I"):
    lis[f"{c}1"].fill = PatternFill(fill_type=None)
for i, (cat, tipo) in enumerate(CATEGORIAS, start=2):
    lis.cell(row=i, column=1, value=cat).fill = fill_input
    lis.cell(row=i, column=2, value=tipo).fill = fill_input
for i, (medio, com) in enumerate(MEDIOS, start=2):
    lis.cell(row=i, column=4, value=medio).fill = fill_input
    c = lis.cell(row=i, column=5, value=com)
    c.fill = fill_input
    c.number_format = "0.0%"
lis["E2"].comment = Comment(
    "Porcentajes de EJEMPLO. Poné lo que realmente te descuenta cada medio "
    "(Mercado Pago, posnet, banco) incluyendo IVA.", "Planilla")
for i, m in enumerate(MESES, start=2):
    lis.cell(row=i, column=8, value=m)
for i, r in enumerate(RUBROS, start=2):
    lis.cell(row=i, column=10, value=r)
lis["L1"] = "Podés agregar categorías o medios de pago abajo de los existentes: aparecen solos en las listas desplegables."
lis["L1"].font = f_sub
lista(lis, "B2:B40", '"Fijo,Variable"')
anchos(lis, {"A": 30, "B": 11, "C": 3, "D": 22, "E": 12, "F": 3, "G": 3, "H": 13, "I": 3, "J": 15})
lis.freeze_panes = "A2"

# ------------------------------------------------------------------ Servicios
srv["A1"] = "Mis servicios y precios"
srv["A1"].font = f_titulo
srv["A2"] = ("Lo cremita lo completás vos; lo gris se calcula solo. Los precios y costos son de EJEMPLO: cambialos por los tuyos. "
             "El costo de insumos es una estimación de cuánto producto gastás por sesión.")
srv["A2"].font = f_sub
srv.merge_cells("A2:H2")
srv["A2"].alignment = Alignment(wrap_text=True, vertical="top")
srv.row_dimensions[2].height = 30
encabezado(srv, 4, ["Servicio", "Rubro", "Precio de lista", "Duración (min)", "Costo insumos por sesión",
                    "Ganancia por sesión", "Margen", "Ganancia por hora"], autos=(6, 7, 8))
SRV_INI, SRV_FIN = 5, 34
for r in range(SRV_INI, SRV_FIN + 1):
    for col in range(1, 6):
        srv.cell(row=r, column=col).fill = fill_input
        srv.cell(row=r, column=col).font = f_norm
    srv[f"F{r}"] = f'=IF(C{r}="","",C{r}-E{r})'
    srv[f"G{r}"] = f'=IF(OR(C{r}="",C{r}=0),"",F{r}/C{r})'
    srv[f"H{r}"] = f'=IF(OR(C{r}="",D{r}="",D{r}=0),"",F{r}/(D{r}/60))'
    for col, fmt in (("C", PESOS), ("E", PESOS), ("F", PESOS), ("G", PORC), ("H", PESOS)):
        srv[f"{col}{r}"].number_format = fmt
    for col in "FGH":
        srv[f"{col}{r}"].fill = fill_auto
        srv[f"{col}{r}"].font = f_auto
for i, s in enumerate(SERVICIOS):
    for j, v in enumerate(s):
        srv.cell(row=SRV_INI + i, column=j + 1, value=v)
lista(srv, f"B{SRV_INI}:B{SRV_FIN}", "=Listas!$J$2:$J$10")
anchos(srv, {"A": 32, "B": 15, "C": 15, "D": 12, "E": 15, "F": 15, "G": 10, "H": 15})
srv.freeze_panes = "A5"

# ------------------------------------------------------------------ Sesiones (ingresos)
encabezado(ses, 1, ["Fecha", "Clienta / paciente", "Servicio", "Rubro", "Cobrado", "Medio de pago",
                    "Comisión de cobro", "Neto que te queda", "Insumos estimados", "Ganancia estimada", "Nota"],
           autos=(4, 7, 8, 9, 10))
ses["E1"].comment = Comment("Lo que te pagó realmente (con descuento, seña, etc.).", "Planilla")
ses["G1"].comment = Comment("Se calcula con el % de la hoja Listas según el medio de pago.", "Planilla")
ses["J1"].comment = Comment(
    "Neto menos el costo estimado de insumos del servicio. Sirve para ver qué servicios te dejan más. "
    "El Panel NO resta estos insumos: resta lo que cargás en Gastos (para no contarlos dos veces).", "Planilla")
for r in range(2, FILAS + 2):
    ses[f"D{r}"] = f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r},Servicios!$A:$B,2,FALSE),"Otro"))'
    ses[f"G{r}"] = f'=IF(E{r}="","",E{r}*IFERROR(VLOOKUP(F{r},Listas!$D:$E,2,FALSE),0))'
    ses[f"H{r}"] = f'=IF(E{r}="","",E{r}-G{r})'
    ses[f"I{r}"] = f'=IF(OR(E{r}="",C{r}=""),"",IFERROR(VLOOKUP(C{r},Servicios!$A:$E,5,FALSE),0))'
    ses[f"J{r}"] = f'=IF(E{r}="","",H{r}-N(I{r}))'
    ses[f"A{r}"].number_format = FECHA
    for col in "EGHIJ":
        ses[f"{col}{r}"].number_format = PESOS
    for col in "DGHIJ":
        ses[f"{col}{r}"].fill = fill_auto
        ses[f"{col}{r}"].font = f_auto
    for col in "ABCEFK":
        ses[f"{col}{r}"].font = f_norm
ejemplos_ses = [
    (date(ANIO, 11, 3), "Lucía G.", "Limpieza facial profunda", 30000, "Transferencia", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 4), "Sofía R.", "Peeling químico", 40000, "Mercado Pago", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 5), "Carla M.", "Consulta nutricional inicial", 35000, "Efectivo", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 5), "Ana P.", "Microneedling", 45000, "Tarjeta de crédito", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 6), "Lucía G.", "Hidratación facial", 25000, "Transferencia", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 6), "Valen T.", "Dermaplaning", 32000, "Mercado Pago", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 7), "Mica S.", "Radiofrecuencia facial", 35000, "Transferencia", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 7), "Carla M.", "Control nutricional", 25000, "Efectivo", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 7), "Juli F.", "Limpieza facial profunda", 28000, "Mercado Pago", "Ejemplo: borrá esta fila (con descuento)"),
]
for i, (f, cli, s, monto, medio, nota) in enumerate(ejemplos_ses, start=2):
    ses[f"A{i}"], ses[f"B{i}"], ses[f"C{i}"], ses[f"E{i}"], ses[f"F{i}"], ses[f"K{i}"] = f, cli, s, monto, medio, nota
fin = FILAS + 1
lista(ses, f"C2:C{fin}", f"=Servicios!$A${SRV_INI}:$A${SRV_FIN}")
lista(ses, f"F2:F{fin}", "=Listas!$D$2:$D$20")
dv_fecha = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
dv_fecha.error = "Poné una fecha (ej. 07/11/2026)."
ses.add_data_validation(dv_fecha)
dv_fecha.add(f"A2:A{fin}")
ses.conditional_formatting.add(f"J2:J{fin}", CellIsRule(operator="lessThan", formula=["0"],
                                                        font=Font(name=FUENTE, color="C0392B", bold=True)))
anchos(ses, {"A": 12, "B": 22, "C": 30, "D": 14, "E": 13, "F": 18, "G": 13, "H": 14, "I": 13, "J": 14, "K": 30})
ses.freeze_panes = "B2"
ses.auto_filter.ref = f"A1:K{fin}"

# ------------------------------------------------------------------ Gastos
encabezado(gas, 1, ["Fecha", "Concepto", "Categoría", "Tipo", "Monto", "Medio de pago", "Proveedor", "Nota"], autos=(4,))
gas["D1"].comment = Comment("Fijo o Variable, según la categoría (se cambia en la hoja Listas).", "Planilla")
for r in range(2, FILAS + 2):
    gas[f"D{r}"] = f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r},Listas!$A:$B,2,FALSE),""))'
    gas[f"A{r}"].number_format = FECHA
    gas[f"E{r}"].number_format = PESOS
    gas[f"D{r}"].fill = fill_auto
    gas[f"D{r}"].font = f_auto
    for col in "ABCEFGH":
        gas[f"{col}{r}"].font = f_norm
ejemplos_gas = [
    (date(ANIO, 11, 1), "Alquiler noviembre", "Alquiler del consultorio", 180000, "Transferencia", "", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 2), "Ácido glicólico y máscaras", "Insumos y productos", 42000, "Tarjeta de débito", "Distribuidora", "Ejemplo: borrá esta fila"),
    (date(ANIO, 11, 6), "Monotributo", "Monotributo e impuestos", 38000, "Transferencia", "ARCA", "Ejemplo: borrá esta fila"),
]
for i, fila in enumerate(ejemplos_gas, start=2):
    for col, v in zip("ABCEFGH", fila):
        gas[f"{col}{i}"] = v
lista(gas, f"C2:C{fin}", "=Listas!$A$2:$A$40")
lista(gas, f"F2:F{fin}", "=Listas!$D$2:$D$20")
dv_fecha2 = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
dv_fecha2.error = "Poné una fecha (ej. 07/11/2026)."
gas.add_data_validation(dv_fecha2)
dv_fecha2.add(f"A2:A{fin}")
anchos(gas, {"A": 12, "B": 32, "C": 26, "D": 10, "E": 13, "F": 18, "G": 20, "H": 30})
gas.freeze_panes = "B2"
gas.auto_filter.ref = f"A1:H{fin}"

# ------------------------------------------------------------------ Panel
P = panel
P["A1"] = "Finanzas del consultorio"
P["A1"].font = f_titulo
P["A2"] = "Elegí el mes y el año. Todo se calcula solo con lo que cargás en Sesiones y Gastos."
P["A2"].font = f_sub
P["A4"], P["C4"] = "Mes", "Año"
P["B4"], P["D4"] = MESES[10], ANIO
for c in ("A4", "C4"):
    P[c].font = f_bold
    P[c].alignment = Alignment(horizontal="right")
for c in ("B4", "D4"):
    P[c].fill = fill_input
    P[c].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
    P[c].alignment = Alignment(horizontal="center")
lista(P, "B4", "=Listas!$H$2:$H$13")
# Fechas auxiliares (columna Z, oculta)
P["Z1"] = "=DATE($D$4,MATCH($B$4,Listas!$H$2:$H$13,0),1)"
P["Z2"] = "=EDATE($Z$1,1)"
P.column_dimensions["Z"].hidden = True

MES_S = 'Sesiones!$A:$A,">="&$Z$1,Sesiones!$A:$A,"<"&$Z$2'
MES_G = 'Gastos!$A:$A,">="&$Z$1,Gastos!$A:$A,"<"&$Z$2'

kpis = [
    (6, "Cobrado en el mes", f"=SUMIFS(Sesiones!$E:$E,{MES_S})", PESOS),
    (7, "Comisiones de cobro (MP, tarjetas)", f"=SUMIFS(Sesiones!$G:$G,{MES_S})", PESOS),
    (8, "Gastos del mes", f"=SUMIFS(Gastos!$E:$E,{MES_G})", PESOS),
    (9, "GANANCIA DEL MES", "=B6-B7-B8", PESOS),
    (10, "Margen (de cada $100 cobrados te quedan)", '=IF(B6=0,"",B9/B6)', PORC),
    (11, "Sesiones atendidas", f"=COUNTIFS({MES_S})", "0"),
    (12, "Cobro promedio por sesión", '=IF(B11=0,"",B6/B11)', PESOS),
    (13, "Meta de ganancia mensual", 600000, PESOS),
    (14, "Avance de la meta", '=IF(OR(B13="",B13=0),"",B9/B13)', PORC),
]
P["A5"] = "Resumen del mes"
P["A5"].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
for fila, etiqueta, formula, fmt in kpis:
    P[f"A{fila}"] = etiqueta
    P[f"A{fila}"].font = f_norm
    P[f"B{fila}"] = formula
    P[f"B{fila}"].number_format = fmt
    P[f"B{fila}"].font = f_bold
    P[f"A{fila}"].border = P[f"B{fila}"].border = borde
P["B13"].fill = fill_input
P["B13"].comment = Comment("Valor de EJEMPLO: poné cuánto querés que te quede por mes.", "Planilla")
for c in ("A9", "B9"):
    P[c].fill = fill_rosa
    P[c].font = Font(name=FUENTE, size=13, bold=True, color=OSCURO)
P.conditional_formatting.add("B9", CellIsRule(operator="lessThan", formula=["0"],
                                              font=Font(name=FUENTE, size=13, bold=True, color="C0392B")))

P["A16"] = "De dónde vino la plata"
P["A16"].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
P["A17"], P["B17"] = "Cosmetología", f'=SUMIFS(Sesiones!$E:$E,{MES_S},Sesiones!$D:$D,"Cosmetología")'
P["A18"], P["B18"] = "Nutrición", f'=SUMIFS(Sesiones!$E:$E,{MES_S},Sesiones!$D:$D,"Nutrición")'
P["A19"], P["B19"] = "Otros", "=B6-B17-B18"
P["A21"] = "En qué se fue"
P["A21"].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
P["A22"], P["B22"] = "Gastos fijos (alquiler, monotributo…)", f'=SUMIFS(Gastos!$E:$E,{MES_G},Gastos!$D:$D,"Fijo")'
P["A23"], P["B23"] = "Gastos variables (insumos, cursos…)", "=B8-B22"
for r in (17, 18, 19, 22, 23):
    P[f"B{r}"].number_format = PESOS
    P[f"A{r}"].font = f_norm
    P[f"B{r}"].font = f_bold
    P[f"A{r}"].border = P[f"B{r}"].border = borde

# Por servicio (mes elegido)
P["D5"] = "Por servicio (mes elegido)"
P["D5"].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
for i, t in enumerate(["Servicio", "Sesiones", "Cobrado", "Ganancia estimada"]):
    c = P.cell(row=6, column=4 + i, value=t)
    c.font, c.fill = f_head, fill_head
    c.alignment = Alignment(horizontal="center", wrap_text=True)
N_SRV = 20
for k in range(N_SRV):
    r, sr = 7 + k, SRV_INI + k
    P[f"D{r}"] = f'=IF(Servicios!A{sr}="","",Servicios!A{sr})'
    P[f"E{r}"] = f'=IF(D{r}="","",COUNTIFS({MES_S},Sesiones!$C:$C,D{r}))'
    P[f"F{r}"] = f'=IF(D{r}="","",SUMIFS(Sesiones!$E:$E,{MES_S},Sesiones!$C:$C,D{r}))'
    P[f"G{r}"] = f'=IF(D{r}="","",SUMIFS(Sesiones!$J:$J,{MES_S},Sesiones!$C:$C,D{r}))'
    P[f"F{r}"].number_format = P[f"G{r}"].number_format = PESOS
    for col in "DEFG":
        P[f"{col}{r}"].font = f_norm
        P[f"{col}{r}"].border = borde

# Gastos por categoría (mes elegido)
P["I5"] = "Gastos por categoría (mes elegido)"
P["I5"].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
for i, t in enumerate(["Categoría", "Gastado", "% del total"]):
    c = P.cell(row=6, column=9 + i, value=t)
    c.font, c.fill = f_head, fill_head
    c.alignment = Alignment(horizontal="center")
for k in range(N_SRV):
    r, lr = 7 + k, 2 + k
    P[f"I{r}"] = f'=IF(Listas!A{lr}="","",Listas!A{lr})'
    P[f"J{r}"] = f'=IF(I{r}="","",SUMIFS(Gastos!$E:$E,{MES_G},Gastos!$C:$C,I{r}))'
    P[f"K{r}"] = f'=IF(OR(I{r}="",$B$8=0),"",J{r}/$B$8)'
    P[f"J{r}"].number_format = PESOS
    P[f"K{r}"].number_format = PORC
    for col in "IJK":
        P[f"{col}{r}"].font = f_norm
        P[f"{col}{r}"].border = borde

# Año completo
TA = 30
P[f"A{TA - 1}"] = "El año mes a mes"
P[f"A{TA - 1}"].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
for i, t in enumerate(["Mes", "Cobrado", "Comisiones", "Gastos", "Ganancia", "Margen", "Sesiones"]):
    c = P.cell(row=TA, column=1 + i, value=t)
    c.font, c.fill = f_head, fill_head
    c.alignment = Alignment(horizontal="center")
for n in range(1, 13):
    r = TA + n
    d1, d2 = f"DATE($D$4,{n},1)", f"DATE($D$4,{n + 1},1)"
    ms = f'Sesiones!$A:$A,">="&{d1},Sesiones!$A:$A,"<"&{d2}'
    mg = f'Gastos!$A:$A,">="&{d1},Gastos!$A:$A,"<"&{d2}'
    P[f"A{r}"] = MESES[n - 1]
    P[f"B{r}"] = f"=SUMIFS(Sesiones!$E:$E,{ms})"
    P[f"C{r}"] = f"=SUMIFS(Sesiones!$G:$G,{ms})"
    P[f"D{r}"] = f"=SUMIFS(Gastos!$E:$E,{mg})"
    P[f"E{r}"] = f"=B{r}-C{r}-D{r}"
    P[f"F{r}"] = f'=IF(B{r}=0,"",E{r}/B{r})'
    P[f"G{r}"] = f"=COUNTIFS({ms})"
tot = TA + 13
P[f"A{tot}"] = "Total del año"
for col in "BCDEG":
    P[f"{col}{tot}"] = f"=SUM({col}{TA + 1}:{col}{TA + 12})"
P[f"F{tot}"] = f'=IF(B{tot}=0,"",E{tot}/B{tot})'
for r in range(TA + 1, tot + 1):
    for col in "BCDE":
        P[f"{col}{r}"].number_format = PESOS
    P[f"F{r}"].number_format = PORC
    for col in "ABCDEFG":
        P[f"{col}{r}"].font = f_bold if r == tot else f_norm
        P[f"{col}{r}"].border = borde
        if r == tot:
            P[f"{col}{r}"].fill = fill_rosa
P.conditional_formatting.add(f"E{TA + 1}:E{tot}", CellIsRule(operator="lessThan", formula=["0"],
                                                             font=Font(name=FUENTE, color="C0392B", bold=True)))

barras = BarChart()
barras.title = "Cobrado vs. gastos y ganancia"
barras.y_axis.title = "$"
barras.add_data(Reference(P, min_col=2, min_row=TA, max_row=TA + 12), titles_from_data=True)
barras.add_data(Reference(P, min_col=4, min_row=TA, max_row=TA + 12), titles_from_data=True)
barras.set_categories(Reference(P, min_col=1, min_row=TA + 1, max_row=TA + 12))
barras.series[0].graphicalProperties.solidFill = OSCURO
barras.series[1].graphicalProperties.solidFill = ROSA_FUERTE
lineal = LineChart()
lineal.add_data(Reference(P, min_col=5, min_row=TA, max_row=TA + 12), titles_from_data=True)
lineal.series[0].graphicalProperties.line.solidFill = "E0A800"
lineal.series[0].graphicalProperties.line.width = 28000
barras += lineal
barras.height, barras.width = 9, 18
P.add_chart(barras, f"I{TA - 1}")

anchos(P, {"A": 40, "B": 16, "C": 13, "D": 30, "E": 12, "F": 14, "G": 16, "H": 3, "I": 28, "J": 14, "K": 11})
P.sheet_view.showGridLines = False

# ------------------------------------------------------------------ Cómo usar
A = ayuda
A.sheet_view.showGridLines = False
A["A1"] = "Finanzas del consultorio"
A["A1"].font = f_titulo
A["A2"] = "Cosmetología + nutrición · cuánto entra, cuánto sale y cuánto te queda"
A["A2"].font = f_sub
pasos = [
    ("Antes de empezar (una sola vez)", None),
    ("1.", "En «Servicios» poné tus servicios con su precio, duración y cuánto gastás de insumos por sesión (los que hay son de ejemplo)."),
    ("2.", "En «Listas» revisá el % de comisión de cada medio de pago (Mercado Pago, tarjetas). Podés sumar categorías de gasto."),
    ("3.", "En «Panel» poné tu meta de ganancia mensual."),
    ("4.", "Borrá las filas de ejemplo de «Sesiones» y «Gastos» (dicen «Ejemplo» en la nota). Borrá solo las celdas cremitas, no las grises."),
    ("Todos los días", None),
    ("•", "Cada sesión que cobrás → una fila en «Sesiones»: fecha, clienta, servicio (de la lista), cuánto cobraste y cómo te pagó."),
    ("•", "Cada gasto → una fila en «Gastos»: fecha, qué fue, categoría y monto. Incluye insumos, alquiler, monotributo, cursos, todo lo del trabajo."),
    ("•", "Las columnas grises se completan solas. No escribas en ellas."),
    ("Para mirar", None),
    ("•", "«Panel»: elegí el mes y ves cuánto cobraste, cuánto gastaste y la GANANCIA real, qué servicios rinden más y el año mes a mes."),
    ("•", "Ganancia del mes = lo cobrado − comisiones de Mercado Pago/tarjetas − gastos cargados."),
    ("•", "La «Ganancia estimada» por servicio usa el costo de insumos que pusiste en «Servicios». Es para comparar servicios; "
          "el Panel usa los gastos reales, así los insumos no se cuentan dos veces."),
    ("Colores", None),
    ("", "Cremita = lo completás vos"),
    ("", "Gris = se calcula solo"),
]
r = 4
for a, b in pasos:
    if b is None:
        A[f"A{r}"] = a
        A[f"A{r}"].font = Font(name=FUENTE, size=12, bold=True, color=OSCURO)
        r += 1
        continue
    A[f"A{r}"] = a
    A[f"B{r}"] = b
    A[f"A{r}"].font = f_bold
    A[f"B{r}"].font = f_norm
    A[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    A[f"A{r}"].alignment = Alignment(horizontal="right", vertical="top")
    A.row_dimensions[r].height = 30 if len(b) > 95 else 16
    r += 1
A[f"A{r - 2}"].fill = fill_input
A[f"A{r - 1}"].fill = fill_auto
anchos(A, {"A": 6, "B": 110})

wb.active = 1  # abre en el Panel
wb.save(SALIDA)
print("OK ->", SALIDA)
