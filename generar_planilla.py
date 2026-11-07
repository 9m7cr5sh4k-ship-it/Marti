"""Genera la planilla de finanzas del consultorio.

Uso:  python3 generar_planilla.py   ->  Finanzas-Consultorio.xlsx
Funciona en Excel y en Google Sheets (subir a Drive y abrir con Hojas de cálculo).
"""
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

SALIDA = "Finanzas-Consultorio.xlsx"
FILAS = 1000
ANIO = 2026
MES_INICIAL = "Noviembre"

# ---------------------------------------------------------------- Estilo
TINTA = "2F3E46"      # títulos y encabezados
ACENTO = "B5838D"     # rosa viejo
ACENTO_CLARO = "F6EEEF"
BORDE = "E3E3E3"
GRIS_TXT = "7A7A7A"
CALC_BG = "F7F7F7"
VERDE = "3C7A5A"
ROJO = "B0413E"
FUENTE = "Arial"

PESOS = '"$" #,##0;-"$" #,##0;"–"'
PORC = '0%;-0%;"–"'
ENTERO = '#,##0;-#,##0;"–"'
FECHA = "dd/mm/yyyy"


def font(size=10, bold=False, color="222222", italic=False):
    return Font(name=FUENTE, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", fgColor=color)


fino = Side(style="thin", color=BORDE)
B_ABAJO = Border(bottom=fino)
B_CAJA = Border(left=fino, right=fino, top=fino, bottom=fino)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
IZQ = Alignment(horizontal="left", vertical="center", indent=1)
DER = Alignment(horizontal="right", vertical="center", indent=1)

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
AREAS = ["Cosmetología", "Nutrición", "Otro"]
CATEGORIAS = [
    ("Insumos y productos", "Variable"), ("Descartables", "Variable"),
    ("Alquiler", "Fijo"), ("Servicios (luz, agua, internet)", "Fijo"),
    ("Telefonía", "Fijo"), ("Monotributo e impuestos", "Fijo"),
    ("Matrícula y seguros", "Fijo"), ("Software y suscripciones", "Fijo"),
    ("Aparatología y equipamiento", "Variable"), ("Mantenimiento", "Variable"),
    ("Capacitación", "Variable"), ("Publicidad y redes", "Variable"),
    ("Movilidad", "Variable"), ("Lavandería y limpieza", "Variable"),
    ("Gastos bancarios", "Variable"), ("Otros", "Variable"),
]
MEDIOS = [("Efectivo", 0), ("Transferencia", 0), ("Mercado Pago", 0.03),
          ("Tarjeta de débito", 0.02), ("Tarjeta de crédito", 0.05)]
SERVICIOS = [
    ("Limpieza facial profunda", "Cosmetología", 30000, 60, 4000),
    ("Peeling químico", "Cosmetología", 40000, 45, 7000),
    ("Dermaplaning", "Cosmetología", 32000, 45, 3000),
    ("Hidratación facial", "Cosmetología", 25000, 50, 3500),
    ("Radiofrecuencia facial", "Cosmetología", 35000, 40, 1500),
    ("Microneedling", "Cosmetología", 45000, 60, 8000),
    ("Drenaje linfático", "Cosmetología", 28000, 60, 1000),
    ("Consulta nutricional", "Nutrición", 35000, 60, 0),
    ("Control nutricional", "Nutrición", 25000, 30, 0),
]

wb = Workbook()
res = wb.active
res.title = "Resumen"
ing = wb.create_sheet("Ingresos")
gas = wb.create_sheet("Gastos")
srv = wb.create_sheet("Servicios")
cfg = wb.create_sheet("Configuración")
guia = wb.create_sheet("Guía")
CFG = "'Configuración'"


def titulo_hoja(ws, texto, subtitulo, ultima_col):
    ws.sheet_view.showGridLines = False
    ws["A1"] = texto
    ws["A1"].font = font(16, True, TINTA)
    ws["A2"] = subtitulo
    ws["A2"].font = font(10, color=GRIS_TXT)
    ws.row_dimensions[1].height = 28
    for c in range(1, ultima_col + 1):
        ws.cell(row=2, column=c).border = Border(bottom=Side(style="medium", color=ACENTO))


def encabezados(ws, fila, textos, calculadas=()):
    for i, t in enumerate(textos, start=1):
        c = ws.cell(row=fila, column=i, value=t)
        c.font = font(10, True, "FFFFFF")
        c.fill = fill(TINTA if i not in calculadas else "5C6B73")
        c.alignment = CENTRO
    ws.row_dimensions[fila].height = 32


def lista(ws, rango, formula):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.errorStyle = "warning"
    dv.error = "El valor no está en la lista."
    ws.add_data_validation(dv)
    dv.add(rango)


def validar_fecha(ws, rango):
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
    dv.error = "Ingresá una fecha válida (dd/mm/aaaa)."
    ws.add_data_validation(dv)
    dv.add(rango)


def anchos(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


def col_mas(col, n):
    return chr(ord(col) + n)


# ================================================================ Configuración
titulo_hoja(cfg, "Configuración", "Datos generales, medios de pago y categorías", 7)
cfg["A4"], cfg["A5"] = "Nombre del consultorio", "Meta de ganancia mensual"
cfg["B4"] = "Mi consultorio"
cfg["B5"].number_format = PESOS
for r in (4, 5):
    cfg[f"A{r}"].font = font(10, True)
    cfg[f"B{r}"].font = font(10)
    cfg[f"B{r}"].border = B_CAJA

CFG_INI, CFG_FIN = 8, 37
for col, t in zip("ABDEG", ["Medio de pago", "Comisión", "Categoría de gasto", "Tipo", "Áreas"]):
    c = cfg[f"{col}7"]
    c.value = t
    c.font = font(10, True, "FFFFFF")
    c.fill = fill(TINTA)
    c.alignment = CENTRO
for r in range(CFG_INI, CFG_FIN + 1):
    for col in "DEG":
        cfg[f"{col}{r}"].border = B_ABAJO
        cfg[f"{col}{r}"].font = font(10)
for r in range(CFG_INI, CFG_INI + 10):
    for col in "AB":
        cfg[f"{col}{r}"].border = B_ABAJO
        cfg[f"{col}{r}"].font = font(10)
    cfg[f"B{r}"].number_format = "0.0%"
for i, (m, p) in enumerate(MEDIOS):
    cfg[f"A{CFG_INI + i}"], cfg[f"B{CFG_INI + i}"] = m, p
for i, (c, t) in enumerate(CATEGORIAS):
    cfg[f"D{CFG_INI + i}"], cfg[f"E{CFG_INI + i}"] = c, t
for i, a in enumerate(AREAS):
    cfg[f"G{CFG_INI + i}"] = a
nota = f"A{CFG_INI + 11}"
cfg[nota] = "Comisión: porcentaje que retiene cada medio de pago (con IVA)."
cfg[nota].font = font(9, color=GRIS_TXT, italic=True)
lista(cfg, f"E{CFG_INI}:E{CFG_FIN}", '"Fijo,Variable"')
for i, m in enumerate(MESES):  # columna oculta para el selector del resumen
    cfg[f"I{CFG_INI + i}"] = m
cfg.column_dimensions["I"].hidden = True
anchos(cfg, {"A": 28, "B": 18, "C": 4, "D": 32, "E": 12, "F": 4, "G": 18})

MEDIOS_RNG = f"{CFG}!$A${CFG_INI}:$A${CFG_INI + 9}"
COMIS_RNG = f"{CFG}!$A${CFG_INI}:$B${CFG_INI + 9}"
CATS_RNG = f"{CFG}!$D${CFG_INI}:$D${CFG_FIN}"
CATTIPO_RNG = f"{CFG}!$D${CFG_INI}:$E${CFG_FIN}"
AREAS_RNG = f"{CFG}!$G${CFG_INI}:$G${CFG_INI + 9}"
MESES_RNG = f"{CFG}!$I${CFG_INI}:$I${CFG_INI + 11}"

# ================================================================ Servicios
titulo_hoja(srv, "Servicios", "Precios de lista y costo estimado de insumos por sesión", 8)
encabezados(srv, 4, ["Servicio", "Área", "Precio", "Duración (min)", "Costo de insumos",
                     "Margen por sesión", "Margen %", "Margen por hora"], calculadas=(6, 7, 8))
S_INI, S_FIN = 5, 34
for r in range(S_INI, S_FIN + 1):
    srv[f"F{r}"] = f'=IF(C{r}="","",C{r}-E{r})'
    srv[f"G{r}"] = f'=IF(OR(C{r}="",C{r}=0),"",F{r}/C{r})'
    srv[f"H{r}"] = f'=IF(OR(C{r}="",N(D{r})=0),"",F{r}/(D{r}/60))'
    for col in "ABCDEFGH":
        c = srv[f"{col}{r}"]
        c.border = B_ABAJO
        c.font = font(10, color="555555" if col in "FGH" else "222222")
        if col in "FGH":
            c.fill = fill(CALC_BG)
    for col, fmt in (("C", PESOS), ("D", ENTERO), ("E", PESOS), ("F", PESOS), ("G", PORC), ("H", PESOS)):
        srv[f"{col}{r}"].number_format = fmt
for i, fila in enumerate(SERVICIOS):
    for j, v in enumerate(fila):
        srv.cell(row=S_INI + i, column=j + 1, value=v)
lista(srv, f"B{S_INI}:B{S_FIN}", f"={AREAS_RNG}")
anchos(srv, {"A": 34, "B": 16, "C": 14, "D": 14, "E": 16, "F": 16, "G": 12, "H": 16})
srv.freeze_panes = "A5"

# ================================================================ Ingresos
titulo_hoja(ing, "Ingresos", "Una fila por cada sesión cobrada", 11)
encabezados(ing, 4, ["Fecha", "Paciente", "Servicio", "Área", "Importe", "Medio de pago",
                     "Comisión", "Neto", "Costo de insumos", "Margen", "Observaciones"],
            calculadas=(4, 7, 8, 9, 10))
R0, R1 = 5, 5 + FILAS - 1
for r in range(R0, R1 + 1):
    ing[f"D{r}"] = f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r},Servicios!$A:$B,2,FALSE),"Otro"))'
    ing[f"G{r}"] = f'=IF(E{r}="","",E{r}*IFERROR(VLOOKUP(F{r},{COMIS_RNG},2,FALSE),0))'
    ing[f"H{r}"] = f'=IF(E{r}="","",E{r}-G{r})'
    ing[f"I{r}"] = f'=IF(OR(E{r}="",C{r}=""),"",IFERROR(VLOOKUP(C{r},Servicios!$A:$E,5,FALSE),0))'
    ing[f"J{r}"] = f'=IF(E{r}="","",H{r}-N(I{r}))'
    for col in "ABCDEFGHIJK":
        c = ing[f"{col}{r}"]
        c.border = B_ABAJO
        calc = col in "DGHIJ"
        c.font = font(10, color="555555" if calc else "222222")
        if calc:
            c.fill = fill(CALC_BG)
    ing[f"A{r}"].number_format = FECHA
    for col in "EGHIJ":
        ing[f"{col}{r}"].number_format = PESOS
lista(ing, f"C{R0}:C{R1}", f"=Servicios!$A${S_INI}:$A${S_FIN}")
lista(ing, f"F{R0}:F{R1}", f"={MEDIOS_RNG}")
validar_fecha(ing, f"A{R0}:A{R1}")
ing.conditional_formatting.add(f"J{R0}:J{R1}", CellIsRule(operator="lessThan", formula=["0"], font=font(10, color=ROJO)))
anchos(ing, {"A": 12, "B": 24, "C": 30, "D": 15, "E": 13, "F": 19, "G": 12, "H": 13, "I": 15, "J": 13, "K": 32})
ing.freeze_panes = "B5"
ing.auto_filter.ref = f"A4:K{R1}"

# ================================================================ Gastos
titulo_hoja(gas, "Gastos", "Una fila por cada gasto del consultorio", 8)
encabezados(gas, 4, ["Fecha", "Concepto", "Categoría", "Tipo", "Importe", "Medio de pago",
                     "Proveedor", "Observaciones"], calculadas=(4,))
for r in range(R0, R1 + 1):
    gas[f"D{r}"] = f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r},{CATTIPO_RNG},2,FALSE),""))'
    for col in "ABCDEFGH":
        c = gas[f"{col}{r}"]
        c.border = B_ABAJO
        c.font = font(10, color="555555" if col == "D" else "222222")
    gas[f"D{r}"].fill = fill(CALC_BG)
    gas[f"A{r}"].number_format = FECHA
    gas[f"E{r}"].number_format = PESOS
lista(gas, f"C{R0}:C{R1}", f"={CATS_RNG}")
lista(gas, f"F{R0}:F{R1}", f"={MEDIOS_RNG}")
validar_fecha(gas, f"A{R0}:A{R1}")
anchos(gas, {"A": 12, "B": 34, "C": 30, "D": 11, "E": 13, "F": 19, "G": 22, "H": 32})
gas.freeze_panes = "B5"
gas.auto_filter.ref = f"A4:H{R1}"

# ================================================================ Resumen
R = res
R.sheet_view.showGridLines = False
anchos(R, {"A": 2, **{col_mas("B", i): 13 for i in range(12)}, "N": 2})
for row in (1, 2, 3):  # banda de título
    for col in range(1, 15):
        R.cell(row=row, column=col).fill = fill(TINTA)
R.row_dimensions[1].height = 10
R.row_dimensions[2].height = 34
R.row_dimensions[3].height = 20
R["B2"] = f"={CFG}!B4"
R["B2"].font = font(18, True, "FFFFFF")
R["B3"] = "Resumen financiero"
R["B3"].font = font(10, color="D5DADD")

# Selector de período
R.row_dimensions[5].height = 24
R["B5"] = "Período"
R["B5"].font = font(10, True, TINTA)
R["B5"].alignment = Alignment(vertical="center")
R.merge_cells("C5:D5")
R["C5"] = MES_INICIAL
R["E5"] = ANIO
borde_acento = Side(style="thin", color=ACENTO)
for c in ("C5", "D5", "E5"):
    R[c].border = Border(left=borde_acento, right=borde_acento, top=borde_acento, bottom=borde_acento)
for c in ("C5", "E5"):
    R[c].font = font(11, True, TINTA)
    R[c].alignment = CENTRO
    R[c].fill = fill(ACENTO_CLARO)
lista(R, "C5", f"={MESES_RNG}")

# Fechas auxiliares (columna P, oculta): inicio del mes, inicio del siguiente, inicio del anterior
R["P1"] = f"=DATE($E$5,MATCH($C$5,{MESES_RNG},0),1)"
R["P2"] = "=EDATE($P$1,1)"
R["P3"] = "=EDATE($P$1,-1)"
R.column_dimensions["O"].hidden = True
R.column_dimensions["P"].hidden = True


def rango_mes(hoja, ini="$P$1", fin="$P$2"):
    return f'{hoja}!$A:$A,">="&{ini},{hoja}!$A:$A,"<"&{fin}'


M_I, M_G = rango_mes("Ingresos"), rango_mes("Gastos")
A_I, A_G = rango_mes("Ingresos", "$P$3", "$P$1"), rango_mes("Gastos", "$P$3", "$P$1")
ING, COM, GAS = (f"SUMIFS(Ingresos!$E:$E,{M_I})", f"SUMIFS(Ingresos!$G:$G,{M_I})", f"SUMIFS(Gastos!$E:$E,{M_G})")
ING_A, COM_A, GAS_A = (f"SUMIFS(Ingresos!$E:$E,{A_I})", f"SUMIFS(Ingresos!$G:$G,{A_I})", f"SUMIFS(Gastos!$E:$E,{A_G})")

# Tarjetas. Valor principal en {col}{fila+1}; dato inferior en {col+1}{fila+2}.
#   Fila 1: Ingresos B8 / C9 · Gastos E8 / F9 · Ganancia H8 / I9 · Margen K8 / L9
#   Fila 2: Sesiones B13 / C14 · Ticket E13 / F14 · Comisiones H13 / I14 · Meta K13 / L14
tarjetas = [
    (7, "B", "Ingresos", f"={ING}", PESOS, "Mes anterior", f"={ING_A}", PESOS),
    (7, "E", "Gastos", f"={GAS}", PESOS, "Mes anterior", f"={GAS_A}", PESOS),
    (7, "H", "Ganancia neta", f"={ING}-{COM}-{GAS}", PESOS, "Mes anterior", f"={ING_A}-{COM_A}-{GAS_A}", PESOS),
    (7, "K", "Margen", '=IF(B8=0,"",H8/B8)', PORC, "Mes anterior", '=IF(C9=0,"",I9/C9)', PORC),
    (12, "B", "Sesiones", f"=COUNTIFS({M_I})", ENTERO, "Mes anterior", f"=COUNTIFS({A_I})", ENTERO),
    (12, "E", "Ticket promedio", '=IF(B13=0,"",B8/B13)', PESOS, "Mes anterior", '=IF(C14=0,"",C9/C14)', PESOS),
    (12, "H", "Comisiones de cobro", f"={COM}", PESOS, "Sobre ingresos", '=IF(B8=0,"",H13/B8)', PORC),
    (12, "K", "Avance de la meta", f'=IF(N({CFG}!B5)=0,"",H8/{CFG}!B5)', PORC, "Meta", f"={CFG}!B5", PESOS),
]
for fila, c0, tit, val, fmt, sub, subval, subfmt in tarjetas:
    c1, c2 = col_mas(c0, 1), col_mas(c0, 2)
    R.merge_cells(f"{c0}{fila}:{c2}{fila}")
    R.merge_cells(f"{c0}{fila + 1}:{c2}{fila + 1}")
    R.merge_cells(f"{c1}{fila + 2}:{c2}{fila + 2}")
    R[f"{c0}{fila}"] = tit
    R[f"{c0}{fila}"].font = font(9, True, GRIS_TXT)
    R[f"{c0}{fila}"].alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    R[f"{c0}{fila + 1}"] = val
    R[f"{c0}{fila + 1}"].number_format = fmt
    R[f"{c0}{fila + 1}"].font = font(20, True, TINTA)
    R[f"{c0}{fila + 1}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    R[f"{c0}{fila + 2}"] = sub
    R[f"{c0}{fila + 2}"].font = font(8, color=GRIS_TXT)
    R[f"{c0}{fila + 2}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    R[f"{c1}{fila + 2}"] = subval
    R[f"{c1}{fila + 2}"].number_format = subfmt
    R[f"{c1}{fila + 2}"].font = font(8, True, GRIS_TXT)
    R[f"{c1}{fila + 2}"].alignment = DER
    for rr in range(fila, fila + 3):
        for cc in (c0, c1, c2):
            cell = R[f"{cc}{rr}"]
            cell.fill = fill("FFFFFF" if rr < fila + 2 else "FAFAFA")
            cell.border = Border(
                left=Side(style="thick", color=ACENTO) if cc == c0 else Side(),
                right=fino if cc == c2 else Side(),
                top=fino if rr == fila else Side(),
                bottom=fino if rr == fila + 2 else Side())
    R.row_dimensions[fila].height = 20
    R.row_dimensions[fila + 1].height = 36
    R.row_dimensions[fila + 2].height = 18
R.conditional_formatting.add("H8", CellIsRule(operator="lessThan", formula=["0"], font=font(20, True, ROJO)))


def seccion(ref, texto, hasta):
    R[ref] = texto
    R[ref].font = font(12, True, TINTA)
    fila = int(ref[1:])
    for cc in range(ord(ref[0]), ord(hasta) + 1):
        R[f"{chr(cc)}{fila}"].border = Border(bottom=Side(style="medium", color=ACENTO))
    R.row_dimensions[fila].height = 24


def encab_tabla(fila, defs):
    """defs: lista de (col_inicio, col_fin, texto)."""
    for a, b, t in defs:
        if a != b:
            R.merge_cells(f"{a}{fila}:{b}{fila}")
        c = R[f"{a}{fila}"]
        c.value = t
        c.font = font(9, True, "FFFFFF")
        c.alignment = CENTRO
        for cc in range(ord(a), ord(b) + 1):
            R[f"{chr(cc)}{fila}"].fill = fill(TINTA)
    R.row_dimensions[fila].height = 22


def celda(ref, valor, fmt=None, bold=False, alin=None):
    R[ref] = valor
    R[ref].font = font(10, bold)
    if fmt:
        R[ref].number_format = fmt
    if alin:
        R[ref].alignment = alin


# ---- Ingresos por servicio / Gastos por categoría
T0 = 17
seccion(f"B{T0}", "Ingresos por servicio", "G")
seccion(f"I{T0}", "Gastos por categoría", "M")
encab_tabla(T0 + 1, [("B", "D", "Servicio"), ("E", "E", "Sesiones"), ("F", "F", "Ingresos"), ("G", "G", "% del total"),
                     ("I", "J", "Categoría"), ("K", "K", "Tipo"), ("L", "L", "Importe"), ("M", "M", "% del total")])
N_FILAS = 20
for k in range(N_FILAS):
    r = T0 + 2 + k
    R.merge_cells(f"B{r}:D{r}")
    R.merge_cells(f"I{r}:J{r}")
    celda(f"B{r}", f'=IF(Servicios!A{S_INI + k}="","",Servicios!A{S_INI + k})', alin=IZQ)
    celda(f"E{r}", f'=IF(B{r}="","",COUNTIFS({M_I},Ingresos!$C:$C,B{r}))', ENTERO, alin=CENTRO)
    celda(f"F{r}", f'=IF(B{r}="","",SUMIFS(Ingresos!$E:$E,{M_I},Ingresos!$C:$C,B{r}))', PESOS)
    celda(f"G{r}", f'=IF(OR(B{r}="",$B$8=0),"",F{r}/$B$8)', PORC, alin=CENTRO)
    celda(f"I{r}", f'=IF({CFG}!D{CFG_INI + k}="","",{CFG}!D{CFG_INI + k})', alin=IZQ)
    celda(f"K{r}", f'=IF(I{r}="","",{CFG}!E{CFG_INI + k})', alin=CENTRO)
    celda(f"L{r}", f'=IF(I{r}="","",SUMIFS(Gastos!$E:$E,{M_G},Gastos!$C:$C,I{r}))', PESOS)
    celda(f"M{r}", f'=IF(OR(I{r}="",$E$8=0),"",L{r}/$E$8)', PORC, alin=CENTRO)
    for cc in "BCDEFGIJKLM":
        R[f"{cc}{r}"].border = B_ABAJO
    R[f"K{r}"].font = font(9, color=GRIS_TXT)

# ---- Por área / estructura de gastos
T1 = T0 + 2 + N_FILAS + 1
seccion(f"B{T1}", "Ingresos por área", "G")
seccion(f"I{T1}", "Estructura de gastos", "M")
encab_tabla(T1 + 1, [("B", "D", "Área"), ("E", "E", "Sesiones"), ("F", "F", "Ingresos"), ("G", "G", "% del total"),
                     ("I", "K", "Tipo de gasto"), ("L", "L", "Importe"), ("M", "M", "% del total")])
for k, area in enumerate(AREAS):
    r = T1 + 2 + k
    R.merge_cells(f"B{r}:D{r}")
    celda(f"B{r}", area, alin=IZQ)
    celda(f"E{r}", f'=COUNTIFS({M_I},Ingresos!$D:$D,B{r})', ENTERO, alin=CENTRO)
    celda(f"F{r}", f'=SUMIFS(Ingresos!$E:$E,{M_I},Ingresos!$D:$D,B{r})', PESOS)
    celda(f"G{r}", f'=IF($B$8=0,"",F{r}/$B$8)', PORC, alin=CENTRO)
    for cc in "BCDEFG":
        R[f"{cc}{r}"].border = B_ABAJO
for k, (tipo, etiqueta) in enumerate([("Fijo", "Gastos fijos"), ("Variable", "Gastos variables")]):
    r = T1 + 2 + k
    R.merge_cells(f"I{r}:K{r}")
    celda(f"I{r}", etiqueta, alin=IZQ)
    R[f"O{r}"] = tipo
    celda(f"L{r}", f'=SUMIFS(Gastos!$E:$E,{M_G},Gastos!$D:$D,$O{r})', PESOS)
    celda(f"M{r}", f'=IF($E$8=0,"",L{r}/$E$8)', PORC, alin=CENTRO)
    for cc in "IJKLM":
        R[f"{cc}{r}"].border = B_ABAJO

# ---- Evolución anual
T2 = T1 + 7
seccion(f"B{T2}", "Evolución anual", "M")
R[f"M{T2}"] = '="Año "&$E$5'
R[f"M{T2}"].font = font(9, color=GRIS_TXT)
R[f"M{T2}"].alignment = Alignment(horizontal="right")
encab_tabla(T2 + 1, [("B", "C", "Mes"), ("D", "E", "Ingresos"), ("F", "G", "Comisiones"), ("H", "I", "Gastos"),
                     ("J", "K", "Ganancia neta"), ("L", "L", "Margen"), ("M", "M", "Sesiones")])
PARES = (("B", "C"), ("D", "E"), ("F", "G"), ("H", "I"), ("J", "K"))
for n in range(1, 13):
    r = T2 + 1 + n
    d1, d2 = f"DATE($E$5,{n},1)", f"DATE($E$5,{n + 1},1)"
    mi, mg = rango_mes("Ingresos", d1, d2), rango_mes("Gastos", d1, d2)
    for a, b in PARES:
        R.merge_cells(f"{a}{r}:{b}{r}")
    celda(f"B{r}", MESES[n - 1], alin=IZQ)
    celda(f"D{r}", f"=SUMIFS(Ingresos!$E:$E,{mi})", PESOS)
    celda(f"F{r}", f"=SUMIFS(Ingresos!$G:$G,{mi})", PESOS)
    celda(f"H{r}", f"=SUMIFS(Gastos!$E:$E,{mg})", PESOS)
    celda(f"J{r}", f"=D{r}-F{r}-H{r}", PESOS, bold=True)
    celda(f"L{r}", f'=IF(D{r}=0,"",J{r}/D{r})', PORC, alin=CENTRO)
    celda(f"M{r}", f"=COUNTIFS({mi})", ENTERO, alin=CENTRO)
    for cc in "BCDEFGHIJKLM":
        R[f"{cc}{r}"].border = B_ABAJO
rt = T2 + 14
for a, b in PARES:
    R.merge_cells(f"{a}{rt}:{b}{rt}")
celda(f"B{rt}", "Total", bold=True, alin=IZQ)
for col in "DFHJ":
    celda(f"{col}{rt}", f"=SUM({col}{T2 + 2}:{col}{T2 + 13})", PESOS, bold=True)
celda(f"L{rt}", f'=IF(D{rt}=0,"",J{rt}/D{rt})', PORC, bold=True, alin=CENTRO)
celda(f"M{rt}", f"=SUM(M{T2 + 2}:M{T2 + 13})", ENTERO, bold=True, alin=CENTRO)
for cc in "BCDEFGHIJKLM":
    R[f"{cc}{rt}"].fill = fill(ACENTO_CLARO)
    R[f"{cc}{rt}"].border = Border(top=borde_acento, bottom=borde_acento)
R.conditional_formatting.add(f"J{T2 + 2}:J{rt}", CellIsRule(operator="lessThan", formula=["0"], font=font(10, True, ROJO)))

graf = BarChart()
graf.y_axis.numFmt = '"$" #,##0'
graf.y_axis.majorGridlines = None
graf.add_data(Reference(R, min_col=4, min_row=T2 + 1, max_row=T2 + 13), titles_from_data=True)
graf.add_data(Reference(R, min_col=8, min_row=T2 + 1, max_row=T2 + 13), titles_from_data=True)
graf.set_categories(Reference(R, min_col=2, min_row=T2 + 2, max_row=T2 + 13))
graf.series[0].graphicalProperties.solidFill = TINTA
graf.series[0].graphicalProperties.line.noFill = True
graf.series[1].graphicalProperties.solidFill = ACENTO
graf.series[1].graphicalProperties.line.noFill = True
graf.gapWidth = 60
linea = LineChart()
linea.add_data(Reference(R, min_col=10, min_row=T2 + 1, max_row=T2 + 13), titles_from_data=True)
linea.series[0].graphicalProperties.line.solidFill = VERDE
linea.series[0].graphicalProperties.line.width = 28000
linea.series[0].smooth = False
graf += linea
graf.legend.position = "b"
graf.height, graf.width = 8.5, 30
R.add_chart(graf, f"B{rt + 2}")

R.page_setup.orientation = "portrait"
R.page_setup.fitToWidth = 1
R.page_setup.fitToHeight = 0
R.sheet_properties.pageSetUpPr.fitToPage = True
R.print_area = f"A1:N{rt + 20}"

# ================================================================ Guía
titulo_hoja(guia, "Guía de uso", "Cómo cargar y consultar la información", 3)
pasos = [
    ("1", "Configuración", "Completá el nombre del consultorio, la meta de ganancia mensual y el porcentaje "
                           "que retiene cada medio de pago."),
    ("2", "Servicios", "Cargá cada servicio con su precio, duración y costo estimado de insumos por sesión."),
    ("3", "Ingresos", "Registrá cada sesión cobrada. El servicio y el medio de pago se eligen de la lista; "
                      "el área, la comisión y el neto se calculan solos."),
    ("4", "Gastos", "Registrá cada gasto del consultorio: insumos, alquiler, impuestos, cursos, etc."),
    ("5", "Resumen", "Elegí el mes y el año para ver ingresos, gastos, ganancia neta y su evolución."),
]
for col, t in zip("ABC", ["", "Hoja", "Qué hacer"]):
    guia[f"{col}4"] = t
    guia[f"{col}4"].font = font(10, True, "FFFFFF")
    guia[f"{col}4"].fill = fill(TINTA)
    guia[f"{col}4"].alignment = IZQ
for i, (n, hoja, txt) in enumerate(pasos):
    r = 5 + i
    guia[f"A{r}"], guia[f"B{r}"], guia[f"C{r}"] = n, hoja, txt
    guia[f"A{r}"].font = font(11, True, ACENTO)
    guia[f"A{r}"].alignment = CENTRO
    guia[f"B{r}"].font = font(10, True, TINTA)
    guia[f"B{r}"].alignment = Alignment(vertical="center", indent=1)
    guia[f"C{r}"].font = font(10)
    guia[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="center")
    guia.row_dimensions[r].height = 34
    for col in "ABC":
        guia[f"{col}{r}"].border = B_ABAJO
guia["B12"] = "A tener en cuenta"
guia["B12"].font = font(10, True, TINTA)
notas = [
    "Las columnas con fondo gris se calculan automáticamente: no hace falta completarlas.",
    "Ganancia neta = ingresos − comisiones de cobro − gastos registrados en el mes.",
    "El margen por sesión usa el costo estimado de insumos de cada servicio. Sirve para comparar "
    "servicios; la ganancia neta del resumen se calcula con los gastos reales.",
    "Para agregar un servicio, un medio de pago o una categoría, sumalo al final de su lista.",
]
for i, t in enumerate(notas):
    r = 13 + i
    guia[f"B{r}"] = "•"
    guia[f"B{r}"].alignment = Alignment(horizontal="right", vertical="top")
    guia[f"C{r}"] = t
    guia[f"C{r}"].font = font(10, color="444444")
    guia[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    guia.row_dimensions[r].height = 30 if len(t) > 95 else 18
anchos(guia, {"A": 5, "B": 18, "C": 95})

for ws, color in ((res, TINTA), (ing, ACENTO), (gas, ACENTO), (srv, "9AA5AB"), (cfg, "9AA5AB"), (guia, "9AA5AB")):
    ws.sheet_properties.tabColor = color

wb.active = 0
wb.save(SALIDA)
print("OK ->", SALIDA)
