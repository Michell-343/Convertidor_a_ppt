"""
Utilidades de bajo nivel para escribir DrawingML (el XML de PowerPoint):
rellenos sólidos con transparencia, degradados, trazos, geometrías libres
(curvas Bézier) y cajas de texto sin márgenes.
"""
from __future__ import annotations

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu

EMU_POR_PX = 9525  # 96 px = 1 pulgada = 914400 EMU


def emu(px: float) -> int:
    return int(round(px * EMU_POR_PX))


def _el(tag: str, **attrs) -> etree._Element:
    e = etree.Element(qn(tag))
    for k, v in attrs.items():
        e.set(k, str(v))
    return e


def _sub(padre, tag: str, **attrs):
    e = etree.SubElement(padre, qn(tag))
    for k, v in attrs.items():
        e.set(k, str(v))
    return e


def hex_rgb(rgb) -> str:
    return "%02X%02X%02X" % tuple(int(max(0, min(255, round(c)))) for c in rgb[:3])


def _color(padre, rgb, alfa: float = 1.0):
    c = _sub(padre, "a:srgbClr", val=hex_rgb(rgb))
    if alfa < 0.999:
        _sub(c, "a:alpha", val=int(max(0.0, alfa) * 100000))
    return c


# ---------------------------------------------------------------- rellenos
def relleno_solido(rgb, alfa: float = 1.0):
    e = _el("a:solidFill")
    _color(e, rgb, alfa)
    return e


def sin_relleno():
    return _el("a:noFill")


def relleno_degradado(paradas, tipo: str = "lineal", angulo: float = 0.0,
                      centro=(0.5, 0.5)):
    """paradas: lista de (posicion 0..1, rgb, alfa)."""
    e = _el("a:gradFill", rotWithShape="1")
    gs_lst = _sub(e, "a:gsLst")
    for pos, rgb, alfa in paradas:
        gs = _sub(gs_lst, "a:gs", pos=int(round(max(0, min(1, pos)) * 100000)))
        _color(gs, rgb, alfa)
    if tipo == "lineal":
        _sub(e, "a:lin", ang=int(round((angulo % 360) * 60000)), scaled="0")
    else:
        p = _sub(e, "a:path", path="circle")
        cx, cy = centro
        _sub(p, "a:fillToRect", l=int(cx * 100000), t=int(cy * 100000),
             r=int((1 - cx) * 100000), b=int((1 - cy) * 100000))
    return e


def linea(ancho_emu: int, relleno, extremo: str = "butt", union: str = "miter",
          guiones: list[float] | None = None):
    caps = {"butt": "flat", "round": "rnd", "square": "sq"}
    e = _el("a:ln", w=max(0, int(ancho_emu)), cap=caps.get(extremo, "flat"))
    e.append(relleno)
    if guiones and ancho_emu > 0:
        cd = _sub(e, "a:custDash")
        ancho_px = ancho_emu / EMU_POR_PX
        pares = guiones if len(guiones) % 2 == 0 else guiones * 2
        for i in range(0, len(pares), 2):
            _sub(cd, "a:ds", d=int(pares[i] / ancho_px * 100000),
                 sp=int(pares[i + 1] / ancho_px * 100000))
    if union == "round":
        _sub(e, "a:round")
    elif union == "bevel":
        _sub(e, "a:bevel")
    else:
        _sub(e, "a:miter", lim="800000")
    return e


def sin_linea():
    e = _el("a:ln")
    _sub(e, "a:noFill")
    return e


# --------------------------------------------------------------- geometría
def geometria_libre(subtrazos, izq: float, arr: float, ancho: float, alto: float):
    """subtrazos: lista de listas de comandos en px de diapositiva:
    ('M',x,y) ('L',x,y) ('C',x1,y1,x2,y2,x,y) ('Z',)"""
    w, h = max(1, emu(ancho)), max(1, emu(alto))
    g = _el("a:custGeom")
    for t in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
        _sub(g, t)
    _sub(g, "a:rect", l="0", t="0", r="r", b="b")
    lst = _sub(g, "a:pathLst")
    p = _sub(lst, "a:path", w=w, h=h)

    def pt(padre, x, y):
        _sub(padre, "a:pt", x=emu(x - izq), y=emu(y - arr))

    for sub in subtrazos:
        for cmd in sub:
            c = cmd[0]
            if c == "M":
                pt(_sub(p, "a:moveTo"), cmd[1], cmd[2])
            elif c == "L":
                pt(_sub(p, "a:lnTo"), cmd[1], cmd[2])
            elif c == "C":
                b = _sub(p, "a:cubicBezTo")
                pt(b, cmd[1], cmd[2]); pt(b, cmd[3], cmd[4]); pt(b, cmd[5], cmd[6])
            elif c == "Z":
                _sub(p, "a:close")
    return g


def geometria_preset(nombre: str, ajustes: dict | None = None):
    g = _el("a:prstGeom", prst=nombre)
    av = _sub(g, "a:avLst")
    for k, v in (ajustes or {}).items():
        _sub(av, "a:gd", name=k, fmla=f"val {int(v)}")
    return g


def aplicar_propiedades(forma, geometria=None, relleno=None, ln=None, quitar_estilo=True):
    """Reemplaza geometría/relleno/línea respetando el orden del esquema."""
    spPr = forma._element.spPr
    for hijo in list(spPr):
        if hijo.tag != qn("a:xfrm"):
            spPr.remove(hijo)
    if geometria is not None:
        spPr.append(geometria)
    if relleno is not None:
        spPr.append(relleno)
    if ln is not None:
        spPr.append(ln)
    if quitar_estilo:
        estilo = forma._element.find(qn("p:style"))
        if estilo is not None:
            forma._element.remove(estilo)


def nueva_forma(contenedor, izq, arr, ancho, alto, nombre=None):
    f = contenedor.add_shape(MSO_SHAPE.RECTANGLE, Emu(emu(izq)), Emu(emu(arr)),
                             Emu(max(1, emu(ancho))), Emu(max(1, emu(alto))))
    if nombre:
        f.name = nombre
    return f


def caja_texto(contenedor, izq, arr, ancho, alto, nombre=None):
    tb = contenedor.add_textbox(Emu(emu(izq)), Emu(emu(arr)),
                                Emu(max(1, emu(ancho))), Emu(max(1, emu(alto))))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if nombre:
        tb.name = nombre
    return tb


def estilo_run(run, nombre_fuente: str, tamano_pt: float, negrita: bool, cursiva: bool,
               rgb, alfa: float = 1.0, espaciado_pt: float = 0.0,
               subrayado: bool = False, tachado: bool = False, idioma: str = "es-ES",
               trazo=None):
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", idioma)
    rPr.set("sz", str(int(round(tamano_pt * 100))))
    rPr.set("b", "1" if negrita else "0")
    rPr.set("i", "1" if cursiva else "0")
    if abs(espaciado_pt) > 0.01:
        rPr.set("spc", str(int(round(espaciado_pt * 100))))
    if subrayado:
        rPr.set("u", "sng")
    if tachado:
        rPr.set("strike", "sngStrike")
    for hijo in list(rPr):
        rPr.remove(hijo)
    if trazo is not None:
        rPr.append(trazo)
    rPr.append(relleno_solido(rgb, alfa))
    _sub(rPr, "a:latin", typeface=nombre_fuente)
    _sub(rPr, "a:ea", typeface=nombre_fuente)
    _sub(rPr, "a:cs", typeface=nombre_fuente)


def desagrupar_si_trivial(grupo):
    """Elimina grupos vacíos y saca de su grupo a los hijos únicos."""
    el = grupo._element
    hijos = [h for h in el if h.tag in (qn("p:sp"), qn("p:pic"), qn("p:grpSp"), qn("p:cxnSp"))]
    padre = el.getparent()
    if not hijos:
        padre.remove(el)
    elif len(hijos) == 1:
        # python-pptx mantiene chOff/chExt == off/ext, así que las coordenadas
        # del hijo ya son coordenadas de la diapositiva.
        padre.replace(el, hijos[0])
