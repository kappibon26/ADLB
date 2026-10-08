"""Reconstruye el cuadro de cantidades UNIQUE 76 a partir de la exportación Revit.

Uso: python scripts/build_cuadro.py <entrada.xlsx> <salida.xlsx>

- NIVEL: mismo cuadro, cantidades globales del cap. 7 calculadas con SUMIFS sobre BD_REVIT.
- BD_REVIT: una sola tabla con todos los elementos Revit (antes 6 hojas). La columna
  "Código NIVEL" es la que se edita para asignar un elemento a un ítem.
- CRUCE RCI: matriz ítem x nivel de piso.
- DASHBOARD: KPIs, selector de ítem y gráficos.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]

NIVELES = ["Nivel Cuarto de Bomba Sub", "Nivel Sótano 4", "Nivel Sótano 3", "Nivel Sótano 2",
           "Nivel Sótano 1", "Nivel Semisótano"] + [f"Nivel {i}" for i in range(1, 16)]

REVISION = "Revisión 2026-10-08"

# Ítems nuevos del cap. 7 (no existían en el cuadro y sí están modelados).
# (código, descripción, und., se inserta después de)
NUEVOS = [
    ("7.103", 'Tubería A.C. Sche 10 De ø1 1/4"', "Ml.", "7.07"),
    ("7.104", 'Tubería A.C. Sche 10 De ø2"', "Ml.", "7.08"),
    ("7.105", 'Soporte pera 2"', "Un.", "7.14"),
    ("7.106", 'Soporte Antisísmico Longitudinal ø3"', "Un.", "7.21"),
    ("7.107", 'Soporte Antisísmico Transversal ø3"', "Un.", "7.106"),
    ("7.108", 'Soporte Antisísmico Longitudinal ø4"', "Un.", "7.107"),
    ("7.109", 'Soporte Antisísmico Transversal ø4"', "Un.", "7.108"),
    ("7.110", 'Sello intumescente 1-1/2"', "Un.", "7.27"),
    ("7.111", 'Sello intumescente 4"', "Un.", "7.110"),
    ("7.112", 'Tee mecánica D=2-1/2x1-1/2"', "Un.", "7.54"),
    ("7.113", 'Tee mecánica D=6x2-1/2"', "Un.", "7.112"),
    ("7.114", 'Válvula de Compuerta De ø1"', "Un.", "7.69"),
    ("7.115", 'Válvula de Cheque Ranurado De ø4"', "Un.", "7.76"),
    ("7.116", 'Base en concreto para equipo de bombeo', "Un.", "7.101"),
    ("7.117", 'Tee A.G. SCH-40 D=3"', "Un.", "7.58"),
]

# Asignación de elementos Revit que venían sin código.
# (hoja, texto en familia, tamaño exacto o None, texto en descripción o None, código)
ASIGNACIONES = [
    ("Accesorios", "VKV EasyPac", None, None, "7.36"),
    ("Accesorios", "VALVULA DE COMPUERTA: 1 Inch", '2"ø-2"ø', None, "7.69"),
    ("Accesorios", "VALVULA DE COMPUERTA1", None, None, "7.114"),
    ("Accesorios", "Swing-Check", None, None, "7.115"),
    ("Accesorios", "PASE RUANA", None, None, "7.19"),
    ("Uniones", "IDC_PASE", None, None, "7.19"),
    ("Uniones", "SELLO CORTAFUEGOS", '1 1/2"ø-1 1/2"ø', None, "7.110"),
    ("Uniones", "SELLO CORTAFUEGOS", '4"ø-4"ø', None, "7.111"),
    ("Uniones", "Tee 2: A.G. - TEE", '3"ø-3"ø-3"ø', None, "7.117"),
    ("Uniones", "Tee 2: A.G. - TEE", None, None, "7.58"),
    ("Uniones", 'TEE MECÁNICA (A.C.) - 3"', None, None, "7.50"),
    ("Uniones", "TEE MECANICA", '1 1/2"ø-1"ø', None, "7.48"),
    ("Uniones", "TEE MECANICA", '2 1/2"ø-1 1/2"ø', None, "7.112"),
    ("Uniones", "TEE MECANICA", '2"ø-1"ø', None, "7.53"),
    ("Uniones", "TEE MECANICA", '2"ø-2"ø', None, "7.49"),
    ("Uniones", "TEE MECANICA", '6"ø-2"ø', None, "7.52"),
    ("Uniones", "TEE MECANICA", '6"ø-2 1/2"ø', None, "7.113"),
    ("Tuberías", "R.C.I. - ACERO", '2"ø', None, "7.104"),
    ("Tuberías", "R.C.I. - ACERO", '1 1/4"ø', None, "7.103"),
    ("Tuberías", "TUBASYS", None, None, "7.08"),
    ("Soportes", "SOPORTE TUBERIA", '2"', None, "7.105"),
    ("Soportes", "ANTISISMICOS", '0"', "LONGITUDINAL", "7.22"),
    ("Soportes", "ANTISISMICOS", '0"', "TRANSVERSAL", "7.23"),
    ("Soportes", "ANTISISMICOS", '3"', "LONGITUDINAL", "7.106"),
    ("Soportes", "ANTISISMICOS", '3"', "TRANSVERSAL", "7.107"),
    ("Soportes", "ANTISISMICOS", '4"', "LONGITUDINAL", "7.108"),
    ("Soportes", "ANTISISMICOS", '4"', "TRANSVERSAL", "7.109"),
    ("Equipos", "LS model", None, None, "7.100"),
    ("Equipos", "USE TYPE CATALOG", None, None, "7.101"),
    ("Equipos", "BASE DE CONCRETO", None, None, "7.116"),
]

ARIAL = "Arial"
F_BASE = Font(name=ARIAL, size=10)
F_BOLD = Font(name=ARIAL, size=10, bold=True)
F_HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=ARIAL, size=14, bold=True, color="1F3864")
F_SUB = Font(name=ARIAL, size=10, italic=True, color="595959")
F_INPUT = Font(name=ARIAL, size=10, color="0000FF", bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_CAP = PatternFill("solid", fgColor="D9E1F2")
FILL_GRP = PatternFill("solid", fgColor="F2F2F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_KPI = PatternFill("solid", fgColor="EEF3FA")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
NUM = '#,##0.00;-#,##0.00;"-"'
INT = '#,##0;-#,##0;"-"'
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row, i, lab)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CENTER, BORDER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w


def title(ws, text, sub, span):
    ws["A1"] = "PROYECTO: UNIQUE 76"
    ws["A1"].font = F_SUB
    ws["A2"] = text
    ws["A2"].font = F_TITLE
    ws["A3"] = sub
    ws["A3"].font = F_SUB


# ---------------------------------------------------------------- lectura
wb = openpyxl.load_workbook(SRC)
rows = []
for name in [s for s in wb.sheetnames if s.startswith("RCI - ")]:
    ws = wb[name]
    hdr = [c.value for c in ws[6]]
    cat = name.replace("RCI - ", "")
    for r in ws.iter_rows(min_row=7, values_only=True):
        if not any(v is not None for v in r):
            continue
        d = dict(zip(hdr, r))
        tam = d.get("Tamaño") or d.get("Tamaño nominal") or d.get("Diámetro nominal") or ""
        fam = d.get("Familia o tipo (origen)") or d.get("Tipo de origen") or ""
        cod, origen = d.get("Código de NIVEL"), "Cruce Revit"
        if not cod:
            origen = "SIN REGLA"
            for h, f, t, de, c in ASIGNACIONES:
                if h == cat and f in fam and (t is None or t == tam) and \
                        (de is None or de in (d.get("Descripción de origen") or "")):
                    cod, origen = c, REVISION
                    break
        rows.append(dict(cat=cat, cod=cod, nivel=d.get("Nivel de piso"),
                         cant=d.get("Cantidad de origen"), fam=fam,
                         desc=d.get("Descripción de origen") or "", tam=tam,
                         com=d.get("Comentarios") or "", sug=origen))
for s in [s for s in wb.sheetnames if s.startswith("RCI - ") or s == "Cruce RCI"]:
    del wb[s]

# ---------------------------------------------------------------- NIVEL
nv = wb["NIVEL"]
for mr in list(nv.merged_cells.ranges):
    nv.unmerge_cells(str(mr))
for cod, desc, und, after in NUEVOS:
    fila = next(r for r in range(7, nv.max_row + 1) if str(nv.cell(r, 1).value) == after)
    nv.insert_rows(fila + 1)
    nv.cell(fila + 1, 1, cod), nv.cell(fila + 1, 2, desc), nv.cell(fila + 1, 3, und)
for r in range(7, nv.max_row + 1):  # 7.29 venía sin unidad
    if str(nv.cell(r, 1).value) == "7.29" and not nv.cell(r, 3).value:
        nv.cell(r, 3, "Un.")
LAST = nv.max_row
PIE = next(r for r in range(7, LAST + 1) if str(nv.cell(r, 2).value or "").startswith("COSTO DIRECTO")) - 1
cap7 = []  # (codigo, descripcion, und, fila)
in_cap7 = False
for r in range(1, LAST + 1):
    a, b = nv.cell(r, 1).value, nv.cell(r, 2).value
    for col in range(1, 6):
        c = nv.cell(r, col)
        c.font, c.fill, c.border = F_BASE, PatternFill(), Border()
    if r < 6:
        continue
    if r == 6:
        continue
    is_cap = a is not None and str(a).strip().isdigit()
    if is_cap:
        in_cap7 = str(a).strip() == "7"
    if r > PIE - 1:
        if b:
            for col in range(1, 6):
                nv.cell(r, col).font = F_BOLD
        continue
    if a is None and b is None:
        continue
    if is_cap:
        nv.cell(r, 1).value = int(str(a).strip())
        for col in range(1, 6):
            nv.cell(r, col).font, nv.cell(r, col).fill = F_BOLD, FILL_CAP
    elif nv.cell(r, 3).value is None and not str(a or "").startswith("7."):  # subgrupo
        for col in range(1, 6):
            nv.cell(r, col).font, nv.cell(r, col).fill = F_BOLD, FILL_GRP
    for col in range(1, 6):
        nv.cell(r, col).border = BORDER
    nv.cell(r, 1).alignment = Alignment(horizontal="center", vertical="center")
    nv.cell(r, 2).alignment = WRAP
    nv.cell(r, 3).alignment = Alignment(horizontal="center", vertical="center")
    nv.cell(r, 4).number_format = NUM
    if in_cap7 and a and str(a).startswith("7."):
        cap7.append((str(a), b, nv.cell(r, 3).value or "", r))
        nv.cell(r, 4).value = f'=SUMIFS(BD_REVIT!$H:$H,BD_REVIT!$C:$C,$A{r})'
        nv.cell(r, 5).value = f'=IF(COUNTIF(BD_REVIT!$C:$C,$A{r})>0,"Revit","Manual")'
        nv.cell(r, 5).alignment = Alignment(horizontal="center")
        nv.cell(r, 5).font = Font(name=ARIAL, size=8, color="808080")

nv["B2"], nv["B3"], nv["B4"] = None, "PROYECTO: UNIQUE 76", "CUADRO DE CANTIDADES MONTAJE HIDROSANITARIO Y RCI"
nv["B3"].font = F_TITLE
nv["B4"].font = Font(name=ARIAL, size=11, bold=True, color="1F3864")
nv["B5"] = "Cap. 7 (RCI): cantidades globales calculadas desde BD_REVIT. Ítems 'Manual' se diligencian a mano."
nv["B5"].font = F_SUB
for i, lab in enumerate(["ITEM", "DESCRIPCIÓN", "UND.", "CANT.", "FUENTE"], 1):
    c = nv.cell(6, i, lab)
    c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CENTER, BORDER
for col, w in zip("ABCDE", [8, 70, 8, 13, 9]):
    nv.column_dimensions[col].width = w
nv.freeze_panes = "A7"
nv.sheet_view.showGridLines = False
nv.auto_filter.ref = None

# Las fórmulas cubren hasta MAXR filas de BD_REVIT para que el libro acepte exportaciones más grandes.
MAXR = 5000
CATS = ["Accesorios", "Equipos", "Rociadores", "Soportes", "Tuberías", "Uniones"]
SERIES = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "4A3AA7"]
CNIV = NIVELES + ["="]  # "=" en SUMIFS/COUNTIFS = elementos sin nivel de piso
corto = lambda n: n.replace("Nivel ", "").replace("Cuarto de Bomba Sub", "C. Bomba").replace("=", "Sin nivel")
BD = lambda col: f"BD_REVIT!${col}$6:${col}${MAXR}"

for dn in [n for n, d in wb.defined_names.items() if "#REF!" in str(d.attr_text)]:
    del wb.defined_names[dn]

# ---------------------------------------------------------------- LISTAS
ls = wb.create_sheet("LISTAS")
for col, lab in enumerate(["Código", "Descripción", "Und.", "Nivel de piso", "", "Filtro nivel", "Filtro categoría"], 1):
    ls.cell(1, col, lab).font = F_BOLD
for i, (cod, desc, und, _) in enumerate(cap7, 2):
    ls.cell(i, 1, cod), ls.cell(i, 2, desc), ls.cell(i, 3, und)
for i, n in enumerate(NIVELES, 2):
    ls.cell(i, 4, n)
for i, n in enumerate(["Todos"] + NIVELES, 2):
    ls.cell(i, 6, n)
for i, n in enumerate(["Todas"] + CATS, 2):
    ls.cell(i, 7, n)
ls.sheet_state = "hidden"
N_COD = len(cap7) + 1

# ---------------------------------------------------------------- BD_REVIT
bd = wb.create_sheet("BD_REVIT")
title(bd, "BASE DE DATOS REVIT – RCI",
      "Edite solo la columna amarilla 'Código NIVEL' para asignar elementos. Las demás columnas grises se calculan.", 15)
H = ["ID", "Categoría", "Código NIVEL", "Ítem NIVEL", "Und.", "Nivel de piso", "Orden nivel",
     "Cantidad", "Familia / tipo Revit", "Descripción Revit", "Tamaño / diámetro", "Comentarios",
     "Estado", "Origen asignación", "En filtro"]
header(bd, 5, H, [6, 12, 11, 42, 7, 24, 7, 10, 48, 30, 18, 30, 13, 22, 8])
F_CALC = Font(name=ARIAL, size=10, color="404040")
for i, d in enumerate(rows):
    r = 6 + i
    vals = [i + 1, d["cat"], d["cod"], None, None, d["nivel"], None, d["cant"], d["fam"], d["desc"],
            d["tam"], d["com"], None, d["sug"], None]
    for col, v in enumerate(vals, 1):
        c = bd.cell(r, col, v)
        c.font, c.border = F_BASE, BORDER
    bd.cell(r, 3).font, bd.cell(r, 3).fill = F_INPUT, FILL_INPUT
    bd.cell(r, 3).number_format = "@"
    bd.cell(r, 4).value = f'=IF(C{r}="","",IFERROR(INDEX(NIVEL!$B:$B,MATCH(C{r},NIVEL!$A:$A,0)),"¡Código no existe!"))'
    bd.cell(r, 5).value = f'=IF(C{r}="","",IFERROR(INDEX(NIVEL!$C:$C,MATCH(C{r},NIVEL!$A:$A,0)),""))'
    bd.cell(r, 7).value = f'=IFERROR(MATCH(F{r},LISTAS!$D$2:$D${len(NIVELES)+1},0),99)'
    bd.cell(r, 8).number_format = NUM
    bd.cell(r, 13).value = f'=IF(C{r}="","SIN ASIGNAR","ASIGNADO")'
    bd.cell(r, 15).value = (f'=IF(AND(OR(DASHBOARD!$C$6="Todos",F{r}=DASHBOARD!$C$6),'
                            f'OR(DASHBOARD!$C$7="Todas",B{r}=DASHBOARD!$C$7)),1,0)')
    for col in (4, 5, 7, 13, 15):
        bd.cell(r, col).font = F_CALC
    bd.cell(r, 15).alignment = Alignment(horizontal="center")
BD_LAST = 5 + len(rows)
tab = Table(displayName="TablaRevit", ref=f"A5:{L(len(H))}{BD_LAST}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
bd.add_table(tab)
dv = DataValidation(type="list", formula1=f"=LISTAS!$A$2:$A${N_COD}", allow_blank=True,
                    showErrorMessage=True, errorTitle="Código", error="Use un código del capítulo 7 de NIVEL")
bd.add_data_validation(dv)
dv.add(f"C6:C{MAXR}")
bd.conditional_formatting.add(f"M6:M{MAXR}", CellIsRule(operator="equal", formula=['"SIN ASIGNAR"'],
                              font=Font(color="C00000", bold=True), fill=PatternFill("solid", fgColor="FCE4E4")))
bd.conditional_formatting.add(f"M6:M{MAXR}", CellIsRule(operator="equal", formula=['"ASIGNADO"'],
                              font=Font(color="2E7D32")))
bd.freeze_panes = "D6"
bd.sheet_view.showGridLines = False

# ---------------------------------------------------------------- CRUCE RCI
cr = wb.create_sheet("CRUCE RCI")
title(cr, "CRUCE DE CANTIDADES RCI POR NIVEL DE PISO",
      "Suma de BD_REVIT por ítem y nivel. 'Dif.' compara contra la cantidad global de NIVEL (debe ser 0).", 0)
H = ["Código", "Descripción", "Und."] + [corto(n) for n in CNIV] + ["TOTAL", "NIVEL", "Dif."]
header(cr, 5, H, [8, 44, 6] + [8.5] * len(CNIV) + [11, 11, 7])
for i, n in enumerate(CNIV):  # nombre completo como criterio (fila oculta)
    cr.cell(4, 4 + i, n).font = Font(name=ARIAL, size=7, color="FFFFFF")
cr.row_dimensions[4].hidden = True
cT, cN, cD = 4 + len(CNIV), 5 + len(CNIV), 6 + len(CNIV)
for i, (cod, desc, und, nrow) in enumerate(cap7):
    r = 6 + i
    cr.cell(r, 1, cod), cr.cell(r, 2, desc), cr.cell(r, 3, und)
    for j in range(len(CNIV)):
        col = 4 + j
        cr.cell(r, col).value = f'=SUMIFS({BD("H")},{BD("C")},$A{r},{BD("F")},{L(col)}$4)'
    cr.cell(r, cT).value = f"=SUM(D{r}:{L(cT-1)}{r})"
    cr.cell(r, cN).value = f"=NIVEL!D{nrow}"
    cr.cell(r, cD).value = f"=ROUND({L(cT)}{r}-{L(cN)}{r},2)"
    for col in range(1, cD + 1):
        c = cr.cell(r, col)
        c.font, c.border = F_BASE, BORDER
        if col > 3:
            c.number_format = NUM
    cr.cell(r, cT).font = F_BOLD
CR_LAST = 5 + len(cap7)
r = CR_LAST + 1
cr.cell(r, 2, "Sin asignar (unidades mezcladas, ver BD_REVIT)").font = Font(name=ARIAL, size=9, italic=True, color="C00000")
for j in range(len(CNIV)):
    col = 4 + j
    cr.cell(r, col).value = f'=SUMIFS({BD("H")},{BD("M")},"SIN ASIGNAR",{BD("F")},{L(col)}$4)'
    cr.cell(r, col).number_format = NUM
    cr.cell(r, col).font = Font(name=ARIAL, size=9, color="C00000")
cr.conditional_formatting.add(f"D6:{L(cT-1)}{CR_LAST}",
                              ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                             end_type="max", end_color="5B9BD5"))
cr.conditional_formatting.add(f"{L(cD)}6:{L(cD)}{CR_LAST}",
                              CellIsRule(operator="notEqual", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
cr.freeze_panes = "D6"
cr.sheet_view.showGridLines = False

# ---------------------------------------------------------------- CALC (auxiliar oculta)
ca = wb.create_sheet("CALC")
for col, lab in enumerate(["Código", "Descripción", "Und.", "Cant. filtro", "Clave orden", "Etiqueta"], 1):
    ca.cell(1, col, lab).font = F_BOLD
for i, (cod, desc, und, _) in enumerate(cap7, 2):
    ca.cell(i, 1, cod), ca.cell(i, 2, desc), ca.cell(i, 3, und)
    ca.cell(i, 4).value = f'=SUMIFS({BD("H")},{BD("C")},A{i},{BD("O")},1)'
    ca.cell(i, 5).value = f"=IF(D{i}>0,D{i}+ROW()/10000000,0)"  # desempata valores iguales
    ca.cell(i, 6).value = f'=A{i}&"  "&LEFT(B{i},48)&" ("&C{i}&")"'
CA_LAST = len(cap7) + 1
ca.sheet_state = "hidden"

# ---------------------------------------------------------------- DASHBOARD
db = wb.create_sheet("DASHBOARD", 0)
db.sheet_view.showGridLines = False
db.sheet_view.zoomScale = 90
for col, w in zip("ABCDEFGHIJKLMNOPQRS", [2, 36, 13, 13, 13, 13, 13, 13, 13, 3] + [11] * 9):
    db.column_dimensions[col].width = w
F_SEC = Font(name=ARIAL, size=11, bold=True, color="1F3864")
F_LBL = Font(name=ARIAL, size=8, bold=True, color="595959")
F_KPI = Font(name=ARIAL, size=18, bold=True, color="1F3864")
F_NOTE = Font(name=ARIAL, size=8, italic=True, color="808080")
LINE = Border(bottom=Side(style="medium", color="AF2B1E"))

db["B2"] = "DASHBOARD DE CANTIDADES RCI – UNIQUE 76"
db["B2"].font = Font(name=ARIAL, size=18, bold=True, color="1F3864")
db["B3"] = "Cap. 7 Redes contraincendios · Datos: BD_REVIT · Cambie las celdas amarillas y todo el tablero se recalcula."
db["B3"].font = F_SUB
for c in "BCDEFGHI":
    db[f"{c}3"].border = LINE


def seccion(cell, texto, nota=None):
    db[cell] = texto
    db[cell].font = F_SEC
    if nota:
        r = int(cell[1:])
        db[f"B{r+1}"] = nota
        db[f"B{r+1}"].font = F_NOTE


def tabla_hdr(row, col0, labels):
    for j, lab in enumerate(labels):
        c = db.cell(row, col0 + j, lab)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, CENTER, BORDER


def celda(row, col, val, fmt=None, bold=False):
    c = db.cell(row, col, val)
    c.font, c.border = (F_BOLD if bold else F_BASE), BORDER
    if fmt:
        c.number_format = fmt
    return c


# Filtros
seccion("B5", "FILTROS")
filtros = [(6, "Nivel de piso", "Todos", f"=LISTAS!$F$2:$F${len(NIVELES)+2}"),
           (7, "Categoría Revit", "Todas", f"=LISTAS!$G$2:$G${len(CATS)+2}"),
           (8, "Ítem para detalle por nivel", "7.08", f"=LISTAS!$A$2:$A${N_COD}")]
for r, lab, val, src in filtros:
    db[f"B{r}"] = lab
    db[f"B{r}"].font = F_BOLD
    c = db[f"C{r}"]
    c.value, c.font, c.fill, c.border = val, F_INPUT, FILL_INPUT, BORDER
    c.number_format = "@"
    v = DataValidation(type="list", formula1=src, allow_blank=False)
    db.add_data_validation(v)
    v.add(f"C{r}")
db["D8"] = '=IFERROR(INDEX(LISTAS!$B:$B,MATCH($C$8,LISTAS!$A:$A,0))&" ("&INDEX(LISTAS!$C:$C,MATCH($C$8,LISTAS!$A:$A,0))&")","Código no válido")'
db["D8"].font = Font(name=ARIAL, size=10, italic=True, color="1F3864")
db["D6"] = "Los filtros de nivel y categoría afectan los indicadores, el top 10 y la asignación por categoría."
db["D6"].font = F_NOTE

# Indicadores
seccion("B10", "INDICADORES (con filtros)")
kpis = [
    ("B", "Elementos Revit", f"=SUM({BD('O')})", INT),
    ("C", "Asignados", f'=COUNTIFS({BD("O")},1,{BD("M")},"ASIGNADO")', INT),
    ("D", "Sin asignar", f'=COUNTIFS({BD("O")},1,{BD("M")},"SIN ASIGNAR")', INT),
    ("E", "% asignado", "=IFERROR(C12/B12,0)", "0.0%"),
    ("F", "Ítems con cantidad", f'=COUNTIF(CALC!$D$2:$D${CA_LAST},">0")', INT),
    ("G", "Niveles con elementos", f'=SUMPRODUCT(--(COUNTIFS({BD("F")},LISTAS!$D$2:$D${len(NIVELES)+1},{BD("O")},1)>0))', INT),
    ("H", "Cant. ítem detalle", f'=SUMIFS({BD("H")},{BD("C")},$C$8,{BD("O")},1)', NUM),
    ("I", "Elementos ítem detalle", f'=COUNTIFS({BD("C")},$C$8,{BD("O")},1)', INT),
]
for col, lab, f, fmt in kpis:
    a, b = db[f"{col}11"], db[f"{col}12"]
    a.value, b.value = lab, f
    a.font, b.font = F_LBL, F_KPI
    b.number_format = fmt
    for c in (a, b):
        c.fill, c.alignment = FILL_KPI, CENTER
db.row_dimensions[12].height = 30
db["D12"].font = Font(name=ARIAL, size=18, bold=True, color="C00000")

# 1. Ítem por nivel
R1 = 14
seccion(f"B{R1}", "1. CANTIDAD DEL ÍTEM DE DETALLE POR NIVEL",
        "Respeta el filtro de categoría. La fila del nivel filtrado se resalta.")
tabla_hdr(R1 + 2, 2, ["Nivel de piso", "Cantidad", "Elementos", "Nivel (Revit)"])
cat_ok = '$C$7="Todas"'
for i, n in enumerate(CNIV):
    r = R1 + 3 + i
    crit = '"="' if n == "=" else f"$E{r}"
    celda(r, 2, corto(n))
    celda(r, 3, f'=IF({cat_ok},SUMIFS({BD("H")},{BD("C")},$C$8,{BD("F")},{crit}),'
                f'SUMIFS({BD("H")},{BD("C")},$C$8,{BD("F")},{crit},{BD("B")},$C$7))', NUM)
    celda(r, 4, f'=IF({cat_ok},COUNTIFS({BD("F")},{crit},{BD("B")},"<>"),COUNTIFS({BD("F")},{crit},{BD("B")},$C$7))', INT)
    e = celda(r, 5, "" if n == "=" else n)
    e.font = Font(name=ARIAL, size=8, color="808080")
T1_END = R1 + 2 + len(CNIV)
db.conditional_formatting.add(f"B{R1+3}:E{T1_END}", FormulaRule(
    formula=[f'$E{R1+3}=$C$6'], fill=PatternFill("solid", fgColor="DBE8F9"), font=Font(bold=True, color="1C5CAB")))
db.conditional_formatting.add(f"C{R1+3}:C{T1_END}", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5"))
ch = BarChart()
ch.type, ch.style = "bar", 10
ch.title = "Ítem de detalle por nivel"
ch.add_data(Reference(db, min_col=3, min_row=R1 + 2, max_row=T1_END), titles_from_data=True)
ch.set_categories(Reference(db, min_col=2, min_row=R1 + 3, max_row=T1_END))
ch.x_axis.scaling.orientation = "maxMin"
ch.legend = None
ch.height, ch.width = 11.5, 17
ch.series[0].graphicalProperties.solidFill = "2E5AAC"
ch.series[0].graphicalProperties.line.noFill = True
db.add_chart(ch, f"G{R1+2}")

# 2. Top 10
R2 = T1_END + 3
seccion(f"B{R2}", "2. TOP 10 ÍTEMS POR CANTIDAD (con filtros)",
        "Ordenado por cantidad en su propia unidad (Ml. y Un. se comparan solo como referencia).")
tabla_hdr(R2 + 2, 2, ["Ítem", "Und.", "Cantidad", "Elementos"])
for k in range(1, 11):
    r = R2 + 2 + k
    key = f"LARGE(CALC!$E$2:$E${CA_LAST},{k})"
    fila = f"MATCH({key},CALC!$E$2:$E${CA_LAST},0)"
    celda(r, 2, f'=IF({key}>0,INDEX(CALC!$F$2:$F${CA_LAST},{fila}),"")')
    celda(r, 3, f'=IF({key}>0,INDEX(CALC!$C$2:$C${CA_LAST},{fila}),"")')
    celda(r, 4, f'=IF({key}>0,INDEX(CALC!$D$2:$D${CA_LAST},{fila}),0)', NUM, bold=True)
    celda(r, 5, f'=IF({key}>0,COUNTIFS({BD("C")},INDEX(CALC!$A$2:$A${CA_LAST},{fila}),{BD("O")},1),0)', INT)
T2_END = R2 + 12
db.conditional_formatting.add(f"D{R2+3}:D{T2_END}", DataBarRule(start_type="num", start_value=0, end_type="max", color="AF2B1E"))
ch = BarChart()
ch.type, ch.style = "bar", 10
ch.title = "Top 10 ítems"
ch.add_data(Reference(db, min_col=4, min_row=R2 + 2, max_row=T2_END), titles_from_data=True)
ch.set_categories(Reference(db, min_col=2, min_row=R2 + 3, max_row=T2_END))
ch.x_axis.scaling.orientation = "maxMin"
ch.legend = None
ch.height, ch.width = 8, 17
ch.series[0].graphicalProperties.solidFill = "AF2B1E"
ch.series[0].graphicalProperties.line.noFill = True
db.add_chart(ch, f"G{R2+2}")

# 3. Asignación por categoría
R3 = T2_END + 4
seccion(f"B{R3}", "3. ASIGNACIÓN POR CATEGORÍA", "Respeta el filtro de nivel.")
tabla_hdr(R3 + 2, 2, ["Categoría", "Elementos", "Asignados", "Sin asignar", "% asignado"])
niv_ok = '$C$6="Todos"'
for i, cat in enumerate(CATS):
    r = R3 + 3 + i
    celda(r, 2, cat)
    celda(r, 3, f'=IF({niv_ok},COUNTIFS({BD("B")},$B{r}),COUNTIFS({BD("B")},$B{r},{BD("F")},$C$6))', INT)
    celda(r, 4, f'=IF({niv_ok},COUNTIFS({BD("B")},$B{r},{BD("M")},"ASIGNADO"),COUNTIFS({BD("B")},$B{r},{BD("F")},$C$6,{BD("M")},"ASIGNADO"))', INT)
    celda(r, 5, f"=C{r}-D{r}", INT).font = Font(name=ARIAL, size=10, bold=True, color="C00000")
    celda(r, 6, f"=IFERROR(D{r}/C{r},0)", "0.0%")
T3_END = R3 + 2 + len(CATS)
r = T3_END + 1
celda(r, 2, "TOTAL", bold=True)
for col in (3, 4, 5):
    celda(r, col, f"=SUM({L(col)}{R3+3}:{L(col)}{T3_END})", INT, bold=True)
celda(r, 6, f"=IFERROR(D{r}/C{r},0)", "0.0%", bold=True)
for col in range(2, 7):
    db.cell(r, col).fill = FILL_GRP
db.conditional_formatting.add(f"F{R3+3}:F{T3_END}", ColorScaleRule(
    start_type="num", start_value=0, start_color="F8CBAD", end_type="num", end_value=1, end_color="C6EFCE"))
ch = BarChart()
ch.type, ch.grouping, ch.overlap, ch.style = "bar", "stacked", 100, 10
ch.title = "Asignados vs sin asignar"
ch.add_data(Reference(db, min_col=4, max_col=5, min_row=R3 + 2, max_row=T3_END), titles_from_data=True)
ch.set_categories(Reference(db, min_col=2, min_row=R3 + 3, max_row=T3_END))
ch.x_axis.scaling.orientation = "maxMin"
ch.height, ch.width = 7, 17
ch.series[0].graphicalProperties.solidFill = "2E7D32"
ch.series[1].graphicalProperties.solidFill = "D9822B"
db.add_chart(ch, f"G{R3+2}")

# 4. Elementos por nivel y categoría
R4 = T3_END + 5
seccion(f"B{R4}", "4. ELEMENTOS POR NIVEL Y CATEGORÍA", "Conteo de elementos del modelo (sin filtros).")
tabla_hdr(R4 + 2, 2, ["Nivel de piso"] + CATS + ["Total"])
for i, n in enumerate(CNIV):
    r = R4 + 3 + i
    crit = '"="' if n == "=" else f'"{n}"'
    celda(r, 2, corto(n))
    for j, cat in enumerate(CATS):
        celda(r, 3 + j, f'=COUNTIFS({BD("F")},{crit},{BD("B")},{L(3+j)}${R4+2})', INT)
    celda(r, 9, f"=SUM(C{r}:H{r})", INT, bold=True)
T4_END = R4 + 2 + len(CNIV)
r = T4_END + 1
celda(r, 2, "TOTAL", bold=True)
for col in range(3, 10):
    celda(r, col, f"=SUM({L(col)}{R4+3}:{L(col)}{T4_END})", INT, bold=True)
    db.cell(r, col).fill = FILL_GRP
db.cell(r, 2).fill = FILL_GRP
db.conditional_formatting.add(f"C{R4+3}:H{T4_END}", ColorScaleRule(
    start_type="num", start_value=0, start_color="FFFFFF", end_type="max", end_color="5B9BD5"))
ch = BarChart()
ch.type, ch.grouping, ch.overlap, ch.style = "bar", "stacked", 100, 10
ch.title = "Elementos por nivel"
ch.add_data(Reference(db, min_col=3, max_col=8, min_row=R4 + 2, max_row=T4_END), titles_from_data=True)
ch.set_categories(Reference(db, min_col=2, min_row=R4 + 3, max_row=T4_END))
ch.x_axis.scaling.orientation = "maxMin"
ch.height, ch.width = 12, 17
for s, color in zip(ch.series, SERIES):
    s.graphicalProperties.solidFill = color
    s.graphicalProperties.line.noFill = True
db.add_chart(ch, f"K{R4+2}")
db.freeze_panes = "A5"

# Impresión: cada hoja a una página de ancho
for ws, orient in ((db, "landscape"), (nv, "portrait"), (bd, "landscape"), (cr, "landscape")):
    ws.page_setup.orientation = orient
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_options.horizontalCentered = True
nv.print_title_rows = "6:6"
bd.print_title_rows = "5:5"
cr.print_title_rows = "5:5"

# Ejes visibles en Excel reciente y todas las etiquetas de nivel
for chart in db._charts:
    chart.x_axis.delete = chart.y_axis.delete = False
    chart.x_axis.tickLblSkip = 1
from openpyxl.workbook.properties import CalcProperties
wb.calculation = CalcProperties(fullCalcOnLoad=True)

orden = ["DASHBOARD", "NIVEL", "BD_REVIT", "CRUCE RCI", "LISTAS", "CALC"]
wb._sheets.sort(key=lambda w: orden.index(w.title) if w.title in orden else 99)
wb.active = 0
for ws in wb:
    ws.sheet_view.tabSelected = ws.title == "DASHBOARD"
wb.save(OUT)
print("ok", len(rows), "filas revit,", len(cap7), "ítems cap7")
