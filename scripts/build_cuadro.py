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
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]

NIVELES = ["Nivel Cuarto de Bomba Sub", "Nivel Sótano 4", "Nivel Sótano 3", "Nivel Sótano 2",
           "Nivel Sótano 1", "Nivel Semisótano"] + [f"Nivel {i}" for i in range(1, 16)]

# Sugerencias para elementos sin asignar (no se aplican: solo orientan la revisión)
SUGERENCIAS = [
    # (hoja, texto en familia, texto en tamaño/diámetro, sugerencia)
    ("Accesorios", "VKV EasyPac", "", "7.36 (verificar)"),
    ("Accesorios", "VALVULA DE COMPUERTA: 1 Inch", "2\"", "7.69 (verificar)"),
    ("Accesorios", "VALVULA DE COMPUERTA1", "", "Crear ítem: Válvula compuerta ø1\""),
    ("Accesorios", "Swing-Check", "", "Crear ítem: Cheque ranurado ø4\""),
    ("Accesorios", "PASE RUANA", "", "7.19 (verificar)"),
    ("Uniones", "IDC_PASE", "", "7.19 (verificar)"),
    ("Uniones", "SELLO CORTAFUEGOS", "", "Crear ítem: Sello intumescente (diámetro modelado)"),
    ("Uniones", "Tee 2: A.G. - TEE", "", "7.58 (verificar)"),
    ("Uniones", "TEE MECÁNICA (A.C.) - 3\"", "", "7.50"),
    ("Uniones", "TEE MECÁNICA (A.G.) - 1 1/2\"", "1 1/2\"ø-1\"ø", "7.48"),
    ("Uniones", "TEE MECÁNICA (A.C.) - 2\"", "2\"ø-1\"ø", "7.53 (verificar)"),
    ("Uniones", "TEE MECÁNICA (A.C.) - 2\"", "2\"ø-2\"ø", "7.49"),
    ("Uniones", "TEE MECÁNICA (A.C.) - 6\"", "6\"ø-2\"ø", "7.52"),
    ("Uniones", "TEE MECÁNICA (A.C.) - 6\"", "6\"ø-2 1/2\"ø", "Crear ítem: Tee mecánica 6x2½\""),
    ("Uniones", "TEE MECÁNICA (A.G.) - 1 1/2\"", "2 1/2\"ø-1 1/2\"ø", "Crear ítem: Tee mecánica 2½x1½\""),
    ("Tuberías", "R.C.I. - ACERO", "2\"ø", "Crear ítem: Tubería A.C. Sch10 ø2\""),
    ("Tuberías", "R.C.I. - ACERO", "1 1/4\"ø", "Crear ítem: Tubería A.C. Sch10 ø1¼\""),
    ("Tuberías", "TUBASYS", "", "Crear ítem: Tubería Tubasys RAL3000 ø1½\""),
    ("Soportes", "SOPORTE TUBERIA", "2\"", "Crear ítem: Soporte pera 2\""),
    ("Soportes", "ANTISISMICOS", "3\"", "Crear ítem: Soporte antisísmico ø3\""),
    ("Soportes", "ANTISISMICOS", "4\"", "Crear ítem: Soporte antisísmico ø4\""),
    ("Soportes", "ANTISISMICOS", "0\"", "7.29 (verificar: red vertical/cabezal)"),
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
        sug = ""
        if not d.get("Código de NIVEL"):
            for h, f, t, s in SUGERENCIAS:
                if h == cat and f in fam and t in f'{tam} {d.get("Diámetro nominal") or ""}':
                    sug = s
                    break
        rows.append(dict(cat=cat, cod=d.get("Código de NIVEL"), nivel=d.get("Nivel de piso"),
                         cant=d.get("Cantidad de origen"), fam=fam,
                         desc=d.get("Descripción de origen") or "", tam=tam,
                         com=d.get("Comentarios") or "", sug=sug))
for s in [s for s in wb.sheetnames if s.startswith("RCI - ") or s == "Cruce RCI"]:
    del wb[s]

# ---------------------------------------------------------------- NIVEL
nv = wb["NIVEL"]
LAST = 535
for mr in list(nv.merged_cells.ranges):
    nv.unmerge_cells(str(mr))
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
    if r > 526:
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

# ---------------------------------------------------------------- LISTAS
ls = wb.create_sheet("LISTAS")
ls["A1"], ls["B1"], ls["C1"], ls["D1"] = "Código", "Descripción", "Und.", "Nivel de piso"
for i, (cod, desc, und, _) in enumerate(cap7, 2):
    ls.cell(i, 1, cod), ls.cell(i, 2, desc), ls.cell(i, 3, und)
for i, n in enumerate(NIVELES, 2):
    ls.cell(i, 4, n)
for c in ls[1]:
    c.font = F_BOLD
ls.sheet_state = "hidden"
N_COD = len(cap7) + 1

# ---------------------------------------------------------------- BD_REVIT
bd = wb.create_sheet("BD_REVIT")
title(bd, "BASE DE DATOS REVIT – RCI",
      "Edite solo la columna amarilla 'Código NIVEL' para asignar elementos. Ítem, Und. y Estado se calculan.", 14)
H = ["ID", "Categoría", "Código NIVEL", "Ítem NIVEL", "Und.", "Nivel de piso", "Orden nivel",
     "Cantidad", "Familia / tipo Revit", "Descripción Revit", "Tamaño / diámetro", "Comentarios",
     "Estado", "Código sugerido"]
header(bd, 5, H, [6, 12, 11, 42, 7, 24, 7, 10, 48, 30, 18, 30, 12, 30])
for i, d in enumerate(rows):
    r = 6 + i
    vals = [i + 1, d["cat"], d["cod"], None, None, d["nivel"], None, d["cant"], d["fam"], d["desc"],
            d["tam"], d["com"], None, d["sug"]]
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
BD_LAST = 5 + len(rows)
tab = Table(displayName="TablaRevit", ref=f"A5:{L(len(H))}{BD_LAST}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
bd.add_table(tab)
dv = DataValidation(type="list", formula1=f"=LISTAS!$A$2:$A${N_COD}", allow_blank=True,
                    showErrorMessage=True, errorTitle="Código", error="Use un código del capítulo 7 de NIVEL")
bd.add_data_validation(dv)
dv.add(f"C6:C{BD_LAST + 300}")
bd.conditional_formatting.add(f"M6:M{BD_LAST}", CellIsRule(operator="equal", formula=['"SIN ASIGNAR"'],
                              font=Font(color="C00000", bold=True), fill=PatternFill("solid", fgColor="FCE4E4")))
bd.conditional_formatting.add(f"M6:M{BD_LAST}", CellIsRule(operator="equal", formula=['"ASIGNADO"'],
                              font=Font(color="2E7D32")))
bd.freeze_panes = "D6"
bd.sheet_view.showGridLines = False

# ---------------------------------------------------------------- CRUCE RCI
cr = wb.create_sheet("CRUCE RCI")
title(cr, "CRUCE DE CANTIDADES RCI POR NIVEL DE PISO",
      "Suma de BD_REVIT por ítem y nivel. 'Dif.' compara contra la cantidad global de NIVEL (debe ser 0).", 0)
short = [n.replace("Nivel ", "").replace("Cuarto de Bomba Sub", "C. Bomba") for n in NIVELES]
H = ["Código", "Descripción", "Und."] + short + ["TOTAL", "NIVEL", "Dif."]
header(cr, 5, H, [8, 44, 6] + [8.5] * len(NIVELES) + [11, 11, 7])
for i, n in enumerate(NIVELES):  # nombre completo como criterio (fila oculta)
    cr.cell(4, 4 + i, n).font = Font(name=ARIAL, size=7, color="FFFFFF")
cr.row_dimensions[4].hidden = True
cT, cN, cD = 4 + len(NIVELES), 5 + len(NIVELES), 6 + len(NIVELES)
for i, (cod, desc, und, nrow) in enumerate(cap7):
    r = 6 + i
    cr.cell(r, 1, cod), cr.cell(r, 2, desc), cr.cell(r, 3, und)
    for j in range(len(NIVELES)):
        col = 4 + j
        cr.cell(r, col).value = f'=SUMIFS(BD_REVIT!$H:$H,BD_REVIT!$C:$C,$A{r},BD_REVIT!$F:$F,{L(col)}$4)'
        cr.cell(r, col).number_format = NUM
    cr.cell(r, cT).value = f"=SUM(D{r}:{L(cT-1)}{r})"
    cr.cell(r, cN).value = f"=NIVEL!D{nrow}"
    cr.cell(r, cD).value = f"=ROUND({L(cT)}{r}-{L(cN)}{r},2)"
    for col in range(1, cD + 1):
        c = cr.cell(r, col)
        c.font, c.border = F_BASE, BORDER
        if col > 3:
            c.number_format = NUM
    cr.cell(r, 2).alignment = Alignment(wrap_text=False, vertical="center")
    cr.cell(r, cT).font = F_BOLD
CR_LAST = 5 + len(cap7)
r = CR_LAST + 1
cr.cell(r, 2, "Sin asignar (unidades mezcladas, ver BD_REVIT)").font = Font(name=ARIAL, size=9, italic=True, color="C00000")
for j in range(len(NIVELES)):
    col = 4 + j
    cr.cell(r, col).value = f'=SUMIFS(BD_REVIT!$H:$H,BD_REVIT!$M:$M,"SIN ASIGNAR",BD_REVIT!$F:$F,{L(col)}$4)'
    cr.cell(r, col).number_format = NUM
    cr.cell(r, col).font = Font(name=ARIAL, size=9, color="C00000")
cr.conditional_formatting.add(f"D6:{L(cT-1)}{CR_LAST}",
                              ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                             end_type="max", end_color="5B9BD5"))
cr.conditional_formatting.add(f"{L(cD)}6:{L(cD)}{CR_LAST}",
                              CellIsRule(operator="notEqual", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
cr.freeze_panes = "D6"
cr.sheet_view.showGridLines = False

# ---------------------------------------------------------------- DASHBOARD
db = wb.create_sheet("DASHBOARD", 0)
db.sheet_view.showGridLines = False
db["B2"] = "DASHBOARD DE CANTIDADES RCI – UNIQUE 76"
db["B2"].font = Font(name=ARIAL, size=16, bold=True, color="1F3864")
db["B3"] = "Todo se recalcula desde BD_REVIT. Cambie el código en la celda amarilla para ver otro ítem."
db["B3"].font = F_SUB
for col, w in zip("ABCDEFGHIJKLMN", [2, 26, 14, 14, 14, 3, 26, 14, 14, 14, 3, 14, 14, 14]):
    db.column_dimensions[col].width = w
BDR = f"BD_REVIT!$A$6:$A${BD_LAST}"
kpis = [("Elementos Revit", f"=COUNTA({BDR})", INT),
        ("Asignados", f'=COUNTIF(BD_REVIT!$M$6:$M${BD_LAST},"ASIGNADO")', INT),
        ("Sin asignar", f'=COUNTIF(BD_REVIT!$M$6:$M${BD_LAST},"SIN ASIGNAR")', INT),
        ("% asignado", "=IFERROR(C6/B6,0)", "0.0%"),
        ("Ítems cap. 7 con cantidad", f'=COUNTIF(\'CRUCE RCI\'!{L(cT)}6:{L(cT)}{CR_LAST},">0")', INT),
        ("Niveles con elementos", f'=SUMPRODUCT(--(COUNTIF(BD_REVIT!$F$6:$F${BD_LAST},LISTAS!$D$2:$D${len(NIVELES)+1})>0))', INT)]
# fila 5 etiquetas / fila 6 valores, en B..G (dos bloques de 3)
pos = [("B", 5), ("C", 5), ("D", 5), ("E", 5), ("G", 5), ("H", 5)]
for (lab, f, fmt), (col, row) in zip(kpis, pos):
    a, b = db[f"{col}{row}"], db[f"{col}{row+1}"]
    a.value, b.value = lab, f
    a.font = Font(name=ARIAL, size=8, bold=True, color="595959")
    b.font = Font(name=ARIAL, size=18, bold=True, color="1F3864")
    b.number_format = fmt
    for c in (a, b):
        c.fill, c.alignment = FILL_KPI, CENTER
# corregir referencias de % asignado (B6=elementos, C6=asignados)
db["E6"] = "=IFERROR(C6/B6,0)"

# Selector de ítem
db["B9"] = "ÍTEM SELECCIONADO"
db["B9"].font = F_BOLD
db["B10"] = "Código"
db["C10"] = "7.08"
db["C10"].font, db["C10"].fill, db["C10"].border = F_INPUT, FILL_INPUT, BORDER
db["C10"].number_format = "@"
dv2 = DataValidation(type="list", formula1=f"=LISTAS!$A$2:$A${N_COD}", allow_blank=False)
db.add_data_validation(dv2)
dv2.add("C10")
db["B11"], db["C11"] = "Descripción", '=IFERROR(INDEX(LISTAS!$B:$B,MATCH($C$10,LISTAS!$A:$A,0)),"")'
db["B12"], db["C12"] = "Unidad", '=IFERROR(INDEX(LISTAS!$C:$C,MATCH($C$10,LISTAS!$A:$A,0)),"")'
db["B13"], db["C13"] = "Cantidad global", f'=SUMIFS(BD_REVIT!$H:$H,BD_REVIT!$C:$C,$C$10)'
db["B14"], db["C14"] = "Elementos Revit", f'=COUNTIF(BD_REVIT!$C:$C,$C$10)'
for r in range(10, 15):
    db[f"B{r}"].font = F_BASE
    db[f"C{r}"].font = F_BOLD if r != 10 else F_INPUT
db["C13"].number_format = NUM
db["C13"].font = Font(name=ARIAL, size=12, bold=True, color="1F3864")

# Tabla por nivel del ítem seleccionado
db["B16"], db["C16"], db["D16"] = "Nivel de piso", "Cantidad ítem", "Elementos (todos)"
for c in ("B16", "C16", "D16"):
    db[c].font, db[c].fill, db[c].alignment = F_HDR, FILL_HDR, CENTER
for i, n in enumerate(NIVELES):
    r = 17 + i
    db[f"B{r}"] = n.replace("Nivel ", "")
    db[f"C{r}"] = f'=SUMIFS(BD_REVIT!$H:$H,BD_REVIT!$C:$C,$C$10,BD_REVIT!$F:$F,LISTAS!$D${i+2})'
    db[f"D{r}"] = f'=COUNTIF(BD_REVIT!$F:$F,LISTAS!$D${i+2})'
    db[f"C{r}"].number_format, db[f"D{r}"].number_format = NUM, INT
    for c in ("B", "C", "D"):
        db[f"{c}{r}"].font, db[f"{c}{r}"].border = F_BASE, BORDER
T_END = 16 + len(NIVELES)

# Resumen por categoría
db["G9"] = "RESUMEN POR CATEGORÍA"
db["G9"].font = F_BOLD
for c, lab in zip("GHIJ", ["Categoría", "Elementos", "Asignados", "Sin asignar"]):
    db[f"{c}10"] = lab
    db[f"{c}10"].font, db[f"{c}10"].fill, db[f"{c}10"].alignment = F_HDR, FILL_HDR, CENTER
cats = sorted({d["cat"] for d in rows})
for i, cat in enumerate(cats):
    r = 11 + i
    db[f"G{r}"] = cat
    db[f"H{r}"] = f'=COUNTIF(BD_REVIT!$B$6:$B${BD_LAST},G{r})'
    db[f"I{r}"] = f'=COUNTIFS(BD_REVIT!$B$6:$B${BD_LAST},G{r},BD_REVIT!$M$6:$M${BD_LAST},"ASIGNADO")'
    db[f"J{r}"] = f'=H{r}-I{r}'
    for c in "GHIJ":
        db[f"{c}{r}"].font, db[f"{c}{r}"].border = F_BASE, BORDER
    db[f"J{r}"].font = Font(name=ARIAL, size=10, bold=True, color="C00000")
C_END = 10 + len(cats)
r = C_END + 1
db[f"G{r}"] = "TOTAL"
for c in "HIJ":
    db[f"{c}{r}"] = f"=SUM({c}11:{c}{C_END})"
for c in "GHIJ":
    db[f"{c}{r}"].font, db[f"{c}{r}"].border, db[f"{c}{r}"].fill = F_BOLD, BORDER, FILL_GRP

ch = BarChart()
ch.type = "bar"
ch.title = "Cantidad del ítem seleccionado por nivel"
ch.y_axis.title = None
ch.add_data(Reference(db, min_col=3, min_row=16, max_row=T_END), titles_from_data=True)
ch.set_categories(Reference(db, min_col=2, min_row=17, max_row=T_END))
ch.x_axis.scaling.orientation = "maxMin"
ch.legend = None
ch.height, ch.width = 11, 15
ch.series[0].graphicalProperties.solidFill = "2E5AAC"
db.add_chart(ch, "G20")

ch2 = BarChart()
ch2.type = "col"
ch2.grouping = "stacked"
ch2.overlap = 100
ch2.title = "Elementos por categoría"
ch2.add_data(Reference(db, min_col=9, min_row=10, max_col=10, max_row=C_END), titles_from_data=True)
ch2.set_categories(Reference(db, min_col=7, min_row=11, max_row=C_END))
ch2.height, ch2.width = 7.5, 15
ch2.series[0].graphicalProperties.solidFill = "2E5AAC"
ch2.series[1].graphicalProperties.solidFill = "D9534F"
db.add_chart(ch2, "G43")

orden = ["DASHBOARD", "NIVEL", "BD_REVIT", "CRUCE RCI", "LISTAS"]
wb._sheets.sort(key=lambda w: orden.index(w.title) if w.title in orden else 99)
wb.active = 0
wb.save(OUT)
print("ok", len(rows), "filas revit,", len(cap7), "ítems cap7")
