"""Genera los escenarios vectoriales de Biblioteca adlb. (SVG 1080x1920).

Uso:  python3 escenarios/generar.py
Cada SVG trae capas con id (cielo, fondo, medio, primer-plano) para animar con parallax.
Los maniquíes se paran en y = 1362 (mismo piso del animatic Nº01).
"""
import os
import random

W, H = 1080, 1920
OUT = os.path.dirname(os.path.abspath(__file__))

# paleta: naturaleza en plano + acentos de marca
NEGRO, CREMA, NAR, NAR2 = "#1C1F22", "#F5F2EE", "#F04E14", "#FF8A2A"
VERDES = ["#2E6533", "#3B7A3E", "#4E9147", "#6AAA52"]
EDIF = ["#C9D3DF", "#AEBBCB", "#E1D8CB", "#9AAEC3", "#D4DCE6"]
MADERA, MADERA_OSC = "#B5713A", "#8A5226"


def f(n):
    return f"{n:.1f}".rstrip("0").rstrip(".")


def rect(x, y, w, h, fill, extra=""):
    return f'<rect x="{f(x)}" y="{f(y)}" width="{f(w)}" height="{f(h)}" fill="{fill}" {extra}/>'


def poly(pts, fill, extra=""):
    return f'<polygon points="{" ".join(f"{f(a)},{f(b)}" for a, b in pts)}" fill="{fill}" {extra}/>'


def circ(cx, cy, r, fill, extra=""):
    return f'<circle cx="{f(cx)}" cy="{f(cy)}" r="{f(r)}" fill="{fill}" {extra}/>'


def line(x1, y1, x2, y2, stroke, w, extra=""):
    return f'<line x1="{f(x1)}" y1="{f(y1)}" x2="{f(x2)}" y2="{f(y2)}" stroke="{stroke}" stroke-width="{f(w)}" {extra}/>'


def g(id_, *parts):
    return f'<g id="{id_}">' + "".join(parts) + "</g>"


def grad(id_, stops, vertical=True):
    xy = 'x1="0" y1="0" x2="0" y2="1"' if vertical else 'x1="0" y1="0" x2="1" y2="0"'
    s = "".join(f'<stop offset="{o}" stop-color="{c}"/>' for o, c in stops)
    return f'<linearGradient id="{id_}" {xy}>{s}</linearGradient>'


def doc(name, titulo, defs, *capas):
    body = "".join(capas)
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}">'
           f"<title>{titulo}</title><defs>{defs}</defs>{body}</svg>\n")
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as fh:
        fh.write(svg)


# ---------- piezas reutilizables ----------
def cielo(y0, y1, top="#3E97DC", bot="#CDE7F7", id_="gCielo"):
    return grad(id_, [(0, top), (1, bot)]), rect(0, y0, W, y1 - y0, f"url(#{id_})")


def nube(cx, cy, s, c="#FFFFFF", sombra="#E2EDF6"):
    p = [(-60, 0, 38), (-15, -22, 50), (40, -8, 42), (80, 6, 28), (-95, 10, 24)]
    out = rect(cx - 110 * s, cy + 6 * s, 210 * s, 30 * s, sombra, f'rx="{f(15 * s)}"')
    out += "".join(circ(cx + dx * s, cy + dy * s, r * s, c) for dx, dy, r in p)
    out += rect(cx - 110 * s, cy, 210 * s, 30 * s, c, f'rx="{f(15 * s)}"')
    return out


def montanas(base, picos, color):
    pts = [(0, base)] + picos + [(W, base)]
    return poly(pts, color)


def ciudad(rng, x0, x1, base, hmin, hmax, colores=EDIF, vent="#8EA4BB", wmin=50, wmax=110):
    out, x = [], x0
    while x < x1:
        w = rng.uniform(wmin, wmax)
        h = rng.uniform(hmin, hmax)
        c = rng.choice(colores)
        out.append(rect(x, base - h, w, h, c))
        if vent:
            cols = max(2, int(w // 18))
            for i in range(cols):
                for j in range(int((h - 20) // 26)):
                    if rng.random() < 0.8:
                        out.append(rect(x + 8 + i * (w - 16) / cols, base - h + 14 + j * 26, (w - 16) / cols - 6, 12, vent, 'opacity=".55"'))
        x += w + rng.uniform(-8, 10)
    return "".join(out)


def arbol(x, base, s, verdes=VERDES[1:], tronco="#6B4426"):
    out = rect(x - 7 * s, base - 70 * s, 14 * s, 70 * s, tronco)
    for dx, dy, r, k in [(-28, -90, 34, 0), (26, -92, 32, 1), (0, -125, 38, 2), (-6, -80, 30, 1)]:
        out += circ(x + dx * s, base + dy * s, r * s, verdes[k % len(verdes)])
    return out


def hoja(x, y, s, rot, c="#2E6533", nervio="#21512A"):
    d = f"M0 0C{f(30*s)} {f(-20*s)} {f(70*s)} {f(-20*s)} {f(110*s)} 0C{f(70*s)} {f(20*s)} {f(30*s)} {f(20*s)} 0 0Z"
    return (f'<g transform="translate({f(x)} {f(y)}) rotate({f(rot)})"><path d="{d}" fill="{c}"/>'
            f'<path d="M0 0L{f(108*s)} 0" stroke="{nervio}" stroke-width="{f(3*s)}" fill="none"/></g>')


def mata(x, y, s, rots, cols=("#3B7A3E", "#2E6533", "#4E9147")):
    return "".join(hoja(x, y, s, r, cols[i % len(cols)]) for i, r in enumerate(rots))


def banca(x, y, w, madera=MADERA, osc=MADERA_OSC, hierro=NEGRO):
    """y = altura del asiento; patas hasta y+105."""
    out = ""
    for i in range(3):
        out += rect(x + 10, y - 120 + i * 28, w - 20, 20, madera if i % 2 == 0 else "#C17E45", 'rx="3"')
    out += rect(x, y, w, 22, madera, 'rx="3"') + rect(x, y + 22, w, 8, osc)
    for px in (x + 30, x + w - 50):
        out += rect(px, y - 125, 16, 125, hierro) + rect(px - 6, y + 28, 28, 77, hierro, 'rx="4"')
    return out


# ---------- 1. parque ----------
def parque():
    rng = random.Random(1)
    d, sky = cielo(0, 1120)
    fondo = g("fondo",
              nube(640, 640, 1.0), nube(920, 560, .8), nube(330, 780, .7),
              montanas(1080, [(120, 960), (300, 900), (470, 980), (650, 880), (840, 960), (1000, 910)], "#8FA6C6"),
              ciudad(rng, 360, W, 1090, 120, 330, wmin=45, wmax=85),
              "".join(arbol(x, 1110, rng.uniform(.6, .9)) for x in range(20, W + 40, 70)))
    piso = (rect(0, 1100, W, 300, "#86BC5E") + rect(0, 1100, W, 40, "#6FA84F")
            + '<path d="M560 1100C640 1160 520 1220 700 1270S980 1330 1080 1330V1390H0V1380C300 1380 500 1300 430 1230S480 1140 560 1100Z" fill="#E7DCC8"/>'
            + rect(0, 1380, W, 540, "#E6DCCB")
            + "".join(line(0, y, W, y, "#D5C8B3", 3) for y in range(1470, 1920, 110))
            + "".join(line(x, 1380, x - 160, 1920, "#D5C8B3", 3) for x in range(120, 1300, 220))
            + "".join(f'<ellipse cx="{x}" cy="{y}" rx="{rx}" ry="34" fill="#CFC1AA" opacity=".8"/>' for x, y, rx in [(330, 1560, 210), (180, 1700, 160), (560, 1760, 190), (90, 1480, 120)]))
    farol = (rect(852, 760, 16, 630, NEGRO) + rect(838, 1370, 44, 22, NEGRO)
             + poly([(830, 690), (890, 690), (880, 760), (840, 760)], NEGRO)
             + rect(840, 700, 40, 48, "#FFE7A8") + poly([(825, 690), (895, 690), (860, 660)], NEGRO))
    caneca = rect(780, 1250, 70, 135, "#24433A", 'rx="6"') + rect(774, 1240, 82, 16, "#1B342C", 'rx="4"')
    medio = g("medio", piso, banca(80, 1290, 540), caneca, farol)
    copa = "".join(circ(x, y, r, c) for x, y, r, c in [
        (120, 120, 200, "#3B7A3E"), (380, 60, 170, "#4E9147"), (560, 200, 120, "#3B7A3E"),
        (300, 260, 140, "#2E6533"), (60, 360, 150, "#4E9147"), (230, 430, 90, "#3B7A3E")])
    tronco = '<path d="M20 1500C60 1200 40 900 110 600L170 420L200 440L160 620C120 900 140 1200 120 1500Z" fill="#6B4426"/><path d="M130 640L330 360L350 380L170 690Z" fill="#6B4426"/>'
    frente = g("primer-plano", tronco, copa, mata(-20, 1880, 2.2, [-70, -40, -15, -95]), mata(1100, 1900, 1.6, [-120, -150, -100]))
    doc("01-parque.svg", "Parque", d, g("cielo", sky), fondo, medio, frente)


# ---------- 2. habitación ----------
def habitacion():
    rng = random.Random(2)
    d, sky = cielo(0, 1080, id_="gCielo")
    ventana = g("fondo", sky, nube(930, 300, .7),
                montanas(820, [(780, 700), (900, 640), (1040, 700)], "#8FA6C6"),
                ciudad(rng, 740, W, 860, 100, 260, wmin=40, wmax=70),
                "".join(arbol(x, 900, .6) for x in range(760, W, 60)))
    muro = (rect(0, 0, 740, 1180, "#E7E0D6") + rect(740, 1080, 340, 100, "#DCD3C6")
            + poly([(420, 0), (620, 0), (300, 1180), (60, 1180)], "#F0EAE1", 'opacity=".8"')
            + rect(728, 0, 24, 1100, NEGRO) + rect(728, 1070, 352, 26, NEGRO) + rect(905, 0, 12, 1080, NEGRO))
    estante = (rect(40, 420, 360, 20, "#2A2A2A")
               + rect(70, 330, 26, 90, NAR) + rect(98, 345, 22, 75, "#D8CFC2") + rect(122, 340, 24, 80, "#C24A1A")
               + rect(230, 350, 80, 70, "#C95C2E", 'rx="6"')
               + mata(270, 360, .6, [-160, -120, -60, -20, 40, 80, 120]))
    cuadro = (rect(470, 440, 180, 240, NEGRO) + rect(484, 454, 152, 212, CREMA)
              + rect(510, 560, 40, 106, "#3C3C3C") + rect(552, 520, 50, 146, "#555") + circ(600, 500, 13, NAR))
    cama = (rect(0, 700, 430, 420, "#2B2B2B", 'rx="10"')
            + rect(0, 960, 560, 180, "#F4F2EE", 'rx="14"')
            + rect(60, 900, 220, 110, NAR, 'rx="26"') + rect(250, 880, 200, 120, "#F7F5F1", 'rx="26"')
            + '<path d="M0 1010H520Q570 1010 575 1060L590 1330H0Z" fill="#8E8E8E"/>'
            + '<path d="M0 1010H520Q560 1012 566 1050L420 1060Q200 1080 0 1060Z" fill="#B3B3B3"/>'
            + rect(0, 1300, 600, 40, "#2B2B2B"))
    mesa = (rect(580, 960, 170, 230, "#3A3A3A") + line(590, 1040, 740, 1040, "#262626", 4) + line(590, 1115, 740, 1115, "#262626", 4)
            + rect(655, 1000, 30, 8, "#222") + rect(655, 1075, 30, 8, "#222")
            + rect(690, 940, 40, 20, NEGRO) + line(710, 940, 690, 850, NEGRO, 8) + poly([(670, 860), (720, 830), (735, 860)], NEGRO)
            + rect(600, 920, 40, 40, "#F4F2EE") + mata(620, 925, .35, [-150, -110, -70, -30]))
    piso = (rect(0, 1180, W, 740, "#B07A47")
            + "".join(line(0, y, W, y, "#99683A", 3) for y in range(1240, 1920, 90))
            + "".join(line(x, 1180 + (i % 2) * 60, x, 1920, "#99683A", 2) for i, x in enumerate(range(140, W, 260)))
            + poly([(460, 1250), (1000, 1250), (1080, 1520), (380, 1520)], "#3C3C3C")
            + poly([(700, 1190), (1080, 1190), (1080, 1420), (560, 1420)], "#C79361", 'opacity=".55"'))
    skate = ('<g transform="translate(900 1040) rotate(-6)">' + rect(-34, -250, 68, 500, NEGRO, 'rx="34"')
             + line(-24, -230, -24, 230, NAR2, 4) + circ(-38, -180, 12, NAR2) + circ(-38, 180, 12, NAR2) + "</g>")
    mochila = (rect(950, 1040, 125, 220, NEGRO, 'rx="30"') + rect(965, 1150, 95, 70, "#2D3135", 'rx="12"')
               + '<path d="M985 1045Q1012 990 1040 1045" stroke="#2D3135" stroke-width="10" fill="none"/>')
    doc("02-habitacion.svg", "Habitación", d, g("cielo", ventana), g("medio", muro, estante, cuadro, piso, cama, mesa), g("primer-plano", skate, mochila))


# ---------- 3. oficina ----------
def oficina():
    rng = random.Random(3)
    d, sky = cielo(0, 1180, id_="gCielo")
    vista = g("fondo", sky, nube(380, 420, .8), nube(120, 560, .5),
              montanas(900, [(0, 820), (160, 740), (330, 820), (480, 760), (560, 800)], "#8FA6C6"),
              ciudad(rng, 0, 570, 960, 120, 340, wmin=40, wmax=80),
              "".join(arbol(x, 1010, .7) for x in range(0, 580, 55)))
    muro = (rect(560, 0, 520, 1240, "#D9D6D1") + rect(560, 0, 520, 1240, "#CFCBC5", 'opacity=".4"')
            + rect(0, 0, 32, 1240, NEGRO) + rect(270, 0, 22, 1180, NEGRO) + rect(540, 0, 30, 1240, NEGRO) + rect(0, 1160, 570, 30, NEGRO)
            + rect(0, 1190, 570, 50, "#BDB8B1"))
    luz = rect(520, 40, 520, 26, NEGRO) + rect(540, 66, 480, 10, "#FFFFFF") + poly([(540, 76), (1020, 76), (1080, 260), (480, 260)], "#FFFFFF", 'opacity=".18"')
    afiche = (rect(780, 170, 260, 400, "#151515")
              + "".join(f'<text x="804" y="{230 + i * 58}" font-family="Archivo, \'Arial Narrow\', sans-serif" font-stretch="condensed" font-weight="600" font-size="46" fill="{CREMA}" textLength="{min(len(t) * 22, 212)}" lengthAdjust="spacingAndGlyphs">{t}</text>'
                        for i, t in enumerate(["BIM", "DISEÑO", "TECNOLOGÍA", "MEJORES", "ESPACIOS"]))
              + rect(804, 528, 40, 7, NAR))
    piso = rect(0, 1240, W, 680, "#8F8B86") + poly([(0, 1240), (W, 1240), (W, 1300), (0, 1320)], "#7F7B76")
    cubo = ('<g transform="translate(730 905)" stroke="#D4DBE3" stroke-width="2" fill="none">'
            '<path d="M-60 -20L0 -50L60 -20L0 10Z M-60 -20V40L0 70V10 M60 -20V40L0 70 M-30 -35V25L30 55 M30 -35V25L-30 55 M-60 10L0 40L60 10"/></g>')
    monitor = (rect(560, 760, 340, 230, NEGRO, 'rx="8"') + rect(574, 774, 312, 196, "#2B2F36")
               + rect(574, 774, 312, 16, "#3B4049") + rect(574, 790, 40, 180, "#343941")
               + "".join(rect(580, 800 + i * 18, 26, 6, "#5A616C") for i in range(8)) + cubo
               + rect(712, 990, 36, 70, "#2A2A2A") + rect(670, 1056, 120, 14, "#2A2A2A", 'rx="4"'))
    escritorio = (poly([(330, 1070), (W, 1070), (W, 1110), (300, 1110)], "#D9A86C")
                  + rect(300, 1110, 780, 22, "#B98A50") + rect(330, 1132, 26, 250, NEGRO) + rect(1020, 1132, 26, 230, NEGRO)
                  + rect(560, 1080, 250, 22, "#222", 'rx="4"') + rect(840, 1084, 40, 18, "#222", 'rx="9"')
                  + rect(930, 1010, 60, 66, NEGRO, 'rx="8"') + '<path d="M990 1025Q1015 1040 990 1060" stroke="#1C1F22" stroke-width="8" fill="none"/>'
                  + rect(1015, 990, 50, 80, "#2A2A2A") + line(1025, 990, 1018, 945, NAR2, 6) + line(1040, 990, 1045, 940, "#F7C948", 6) + line(1055, 990, 1062, 950, "#222", 6)
                  + rect(420, 1010, 90, 66, "#F4F2EE", 'rx="6"') + mata(465, 1015, .55, [-150, -120, -90, -60, -30]))
    silla = (rect(250, 830, 230, 330, "#202326", 'rx="40"') + rect(270, 860, 190, 270, "#2C3034", 'rx="30"')
             + rect(230, 1150, 280, 50, "#202326", 'rx="18"') + rect(355, 1200, 30, 120, NEGRO)
             + line(260, 1350, 480, 1350, NEGRO, 16, 'stroke-linecap="round"') + circ(265, 1365, 14, NEGRO) + circ(475, 1365, 14, NEGRO)
             + rect(220, 1040, 40, 14, NEGRO, 'rx="7"') + rect(470, 1040, 40, 14, NEGRO, 'rx="7"'))
    doc("03-oficina.svg", "Oficina", d, g("cielo", vista), g("medio", muro, luz, afiche, piso, monitor, escritorio), g("primer-plano", silla))


# ---------- 4. obra ----------
def obra():
    rng = random.Random(4)
    d, sky = cielo(0, 1300)
    fondo = g("fondo", nube(820, 760, .9), nube(560, 560, .6), nube(220, 300, .5),
              ciudad(rng, 600, W, 1250, 200, 460, wmin=60, wmax=120))
    losas, cols = "", ""
    for y in (520, 720, 920, 1120):
        losas += rect(0, y - 190, 640, 190, "#8E8982", 'opacity=".55"')
    for y in (520, 720, 920, 1120):
        losas += rect(0, y, 650, 30, "#BDB7AE") + line(0, y - 40, 640, y - 40, NAR2, 5) + "".join(line(x, y - 40, x, y, NAR2, 4) for x in range(20, 650, 70))
    for x in (40, 200, 360, 520):
        cols += rect(x, 300, 40, 950, "#A9A39A") + "".join(line(x + 8 + i * 12, 300, x + 8 + i * 12, 260, "#6E6A65", 4) for i in range(3))
    edificio = cols + losas + rect(0, 1250, 660, 10, "#8E8982")
    torre = rect(760, 280, 46, 1000, "none", f'stroke="#E9A72A" stroke-width="6"')
    torre += "".join(line(760, y, 806, y + 46, "#E9A72A", 4) + line(806, y, 760, y + 46, "#E9A72A", 4) for y in range(280, 1280, 46))
    pluma = rect(300, 250, 780, 34, "none", 'stroke="#E9A72A" stroke-width="6"')
    pluma += "".join(line(x, 250, x + 34, 284, "#E9A72A", 4) for x in range(300, 1080, 34))
    grua = (torre + pluma + poly([(745, 250), (820, 250), (783, 170)], "#E9A72A") + rect(810, 286, 60, 50, "#E9A72A")
            + line(783, 170, 320, 250, "#6E6A65", 3) + line(783, 170, 1060, 250, "#6E6A65", 3)
            + rect(960, 284, 90, 70, "#7A756E") + line(470, 284, 470, 760, "#333", 3) + rect(456, 760, 28, 18, NAR) + '<path d="M470 778Q470 800 456 800" stroke="#333" stroke-width="5" fill="none"/>')
    suelo = rect(0, 1260, W, 660, "#D8CDB9") + rect(0, 1260, W, 30, "#C7BBA4") + "".join(line(x, 1300, x - 200, 1920, "#CBBEA6", 3) for x in range(200, 1400, 260))
    caseta = (rect(30, 1150, 360, 240, "#F2F0EB") + rect(20, 1135, 380, 24, "#9C978F")
              + rect(70, 1200, 110, 80, "#3A4048") + rect(230, 1200, 110, 80, "#3A4048") + rect(30, 1380, 360, 14, "#9C978F"))
    contenedor = rect(800, 1150, 280, 140, "#2F67B3") + "".join(line(x, 1150, x, 1290, "#24538F", 5) for x in range(820, 1080, 26))
    malla = rect(390, 1270, 690, 100, NAR, 'opacity=".9"') + "".join(line(x, 1270, x + 25, 1370, "#C23A0C", 3) + line(x + 25, 1270, x, 1370, "#C23A0C", 3) for x in range(390, 1080, 25)) + "".join(rect(x, 1255, 12, 135, "#555") for x in range(390, 1080, 170))
    material = "".join(rect(440, 1330 + i * 16, 330, 12, "#7E7B77", 'rx="6"') for i in range(4)) + rect(460, 1394, 20, 12, "#5A3A22") + rect(730, 1394, 20, 12, "#5A3A22")
    cono = poly([(280, 1480), (320, 1480), (348, 1600), (252, 1600)], NAR) + rect(262, 1530, 76, 18, CREMA) + rect(236, 1598, 128, 16, NAR)
    doc("04-obra.svg", "Obra", d, g("cielo", sky), fondo,
        g("medio", edificio, grua, suelo, caseta, contenedor, malla, material),
        g("primer-plano", cono, mata(1110, 1900, 2.0, [-110, -140, -165, -95]), mata(-30, 1300, .9, [-60, -30, -5])))


# ---------- 5. paradero ----------
def paradero():
    rng = random.Random(5)
    d, sky = cielo(0, 1240)
    fondo = g("fondo", nube(780, 640, .9), nube(980, 820, .6),
              ciudad(rng, 340, W, 1160, 160, 400, wmin=60, wmax=120),
              "".join(arbol(x, 1220, rng.uniform(.8, 1.1)) for x in range(320, W + 60, 80)))
    calle = (rect(0, 1200, W, 60, "#BFC7A0") + rect(0, 1240, W, 340, "#DAD4C9")
             + "".join(line(x, 1240, x - 90, 1580, "#C9C2B5", 3) for x in range(80, 1200, 180))
             + poly([(0, 1500), (W, 1450), (W, 1490), (0, 1545)], "#F2C230")
             + rect(0, 1580, W, 30, "#9C978F") + rect(0, 1610, W, 310, "#3A3D42")
             + "".join(rect(x, 1760, 140, 16, "#E9E6E0") for x in range(40, W, 260)))
    banca_ = banca(330, 1270, 470)
    techo = (poly([(0, 0), (960, 0), (1080, 200), (1080, 300), (0, 300)], "#2B2E33")
             + poly([(0, 300), (1080, 300), (1080, 330), (0, 340)], "#1C1F22")
             + rect(300, 330, 26, 1050, NEGRO) + rect(830, 330, 26, 1050, NEGRO)
             + rect(326, 400, 504, 6, "#3A3D42") + rect(326, 760, 504, 6, "#3A3D42", 'opacity=".6"')
             + rect(326, 330, 504, 830, "#9FC4E4", 'opacity=".18"'))
    texto = "".join(f'<text x="60" y="{700 + i * 74}" font-family="Archivo, \'Arial Narrow\', sans-serif" font-stretch="condensed" font-weight="600" font-size="66" fill="{CREMA}" textLength="{min(len(t) * 32, 210)}" lengthAdjust="spacingAndGlyphs">{t}</text>'
                    for i, t in enumerate(["LA", "DISCIPLINA", "TAMBIÉN", "ES UN", "PROYECTO"]))
    valla = rect(0, 480, 310, 820, "#2B2E33") + rect(20, 500, 270, 780, "#141618") + texto + rect(60, 1060, 50, 9, NAR)
    doc("05-paradero.svg", "Paradero", d, g("cielo", sky), fondo, g("medio", calle, banca_), g("primer-plano", techo, valla))


# ---------- 6. atardecer ----------
def atardecer():
    rng = random.Random(6)
    d = grad("gAtar", [(0, "#9C5A72"), (.45, "#E0703A"), (.8, "#F49A45"), (1, "#F8C177")])
    d += '<radialGradient id="gSol"><stop offset="0" stop-color="#FFF1C9"/><stop offset=".35" stop-color="#FFD58A" stop-opacity=".7"/><stop offset="1" stop-color="#FFD58A" stop-opacity="0"/></radialGradient>'
    sky = rect(0, 0, W, 1240, "url(#gAtar)") + circ(800, 1000, 260, "url(#gSol)") + circ(800, 1000, 70, "#FFEFC2")
    nubes = "".join(rect(x, y, w, 26, c, 'rx="13"') for x, y, w, c in [
        (40, 300, 380, "#8E5068"), (600, 220, 420, "#A55E6C"), (160, 520, 300, "#C2614A"), (700, 600, 340, "#D06A44"), (-40, 760, 260, "#D9753F")])
    lejos = (montanas(1120, [(0, 1010), (180, 950), (360, 1030), (560, 960), (760, 1040), (940, 980), (1080, 1020)], "#7E4E6A")
             + ciudad(rng, 0, W, 1180, 80, 360, colores=["#5E3A55", "#6A4258", "#53344D"], vent="#F6B25E", wmin=50, wmax=100))
    baranda = rect(0, 1150, W, 20, "#1E1A1C") + "".join(rect(x, 1170, 12, 140, "#1E1A1C") for x in range(60, W, 240)) + rect(0, 1300, W, 20, "#1E1A1C")
    mesa = (poly([(0, 1320), (W, 1320), (W, 1920), (0, 1920)], "#7A3F22")
            + "".join(line(0, y, W, y, "#5E2F18", 3) for y in (1400, 1500, 1620, 1760))
            + poly([(0, 1320), (W, 1320), (W, 1340), (0, 1345)], "#F39A55", 'opacity=".5"'))
    termo = (rect(300, 1080, 110, 340, "#1A1A1A", 'rx="22"') + rect(312, 1040, 86, 50, "#262626", 'rx="10"')
             + '<path d="M410 1090Q450 1110 440 1150" stroke="#262626" stroke-width="10" fill="none"/>'
             + rect(318, 1110, 14, 280, "#F39A55", 'opacity=".35" rx="7"'))
    portatil = poly([(560, 1400), (1000, 1380), (1060, 1480), (600, 1510)], "#2A2A2C") + poly([(600, 1510), (1060, 1480), (1060, 1494), (600, 1526)], "#151517") + circ(810, 1446, 12, "#3A3A3D")
    libreta = (poly([(100, 1600), (460, 1580), (480, 1720), (110, 1745)], "#1A1A1A") + line(120, 1610, 130, 1740, "#333", 6)
               + line(260, 1620, 430, 1690, "#111", 9, 'stroke-linecap="round"') + line(400, 1678, 432, 1691, NAR, 9, 'stroke-linecap="round"'))
    planta = rect(900, 1080, 150, 150, "#1E1A1C", 'rx="12"') + mata(975, 1090, 1.3, [-170, -140, -110, -80, -50, -20], ("#2F5A2E", "#3D6E35", "#264A26"))
    doc("06-atardecer.svg", "Atardecer", d, g("cielo", sky, nubes), g("fondo", lejos), g("medio", baranda, mesa), g("primer-plano", termo, portatil, libreta, planta))


if __name__ == "__main__":
    for fn in (parque, habitacion, oficina, obra, paradero, atardecer):
        fn()
    print("ok")
