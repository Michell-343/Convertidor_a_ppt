"""
Conversión SVG -> elementos NATIVOS y editables de PowerPoint.

  <rect>, <circle>, <ellipse>      -> rectángulos / rectángulos redondeados / elipses
  <path>, <polygon>, <polyline>,
  <line>                            -> formas libres (curvas Bézier editables con
                                       "Modificar puntos")
  <text> / <tspan>                  -> cuadros de texto con la fuente, tamaño, peso,
                                       color, interletrado y alineación originales
  <image>                           -> imágenes (recortadas si tienen máscara rectangular)
  <g>                               -> grupos (con el nombre de la capa)
  <linearGradient>/<radialGradient> -> degradados nativos
  <style> (CSS de Illustrator/Figma), transform, opacity, <use>, viewBox -> soportados
"""
from __future__ import annotations

import base64
import io
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote

from lxml import etree
from PIL import Image, ImageColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

from . import pptx_utils as pu
from .fuentes import CatalogoFuentes, ancho_texto, linea_base_ppt, parsear_peso

XLINK = "{http://www.w3.org/1999/xlink}href"
PT_POR_PX = 0.75

HEREDADAS = {
    "fill", "fill-opacity", "fill-rule", "stroke", "stroke-width", "stroke-opacity",
    "stroke-linecap", "stroke-linejoin", "stroke-dasharray", "font-family", "font-size",
    "font-weight", "font-style", "text-anchor", "letter-spacing", "visibility", "color",
    "text-decoration", "xml:space", "white-space", "clip-rule",
}
PROPIEDADES = HEREDADAS | {"opacity", "display", "clip-path", "mask", "filter",
                           "stop-color", "stop-opacity", "transform"}
IGNORAR = {"defs", "clipPath", "mask", "symbol", "linearGradient", "radialGradient",
           "pattern", "style", "title", "desc", "metadata", "marker", "filter", "script",
           "namedview", "font", "font-face", "cursor", "view", "animate", "set"}


# =================================================================== informe
@dataclass
class Informe:
    avisos: list[str] = field(default_factory=list)
    fuentes: dict[str, bool] = field(default_factory=dict)  # nombre -> instalada
    conteo: dict[str, int] = field(default_factory=dict)

    def aviso(self, texto: str):
        if texto not in self.avisos:
            self.avisos.append(texto)

    def contar(self, tipo: str):
        self.conteo[tipo] = self.conteo.get(tipo, 0) + 1


# =================================================================== matrices
class M:
    __slots__ = ("a", "b", "c", "d", "e", "f")

    def __init__(self, a=1.0, b=0.0, c=0.0, d=1.0, e=0.0, f=0.0):
        self.a, self.b, self.c, self.d, self.e, self.f = a, b, c, d, e, f

    def __mul__(self, o: "M") -> "M":
        return M(self.a * o.a + self.c * o.b, self.b * o.a + self.d * o.b,
                 self.a * o.c + self.c * o.d, self.b * o.c + self.d * o.d,
                 self.a * o.e + self.c * o.f + self.e, self.b * o.e + self.d * o.f + self.f)

    def ap(self, x, y):
        return self.a * x + self.c * y + self.e, self.b * x + self.d * y + self.f

    @property
    def escala(self) -> float:
        return math.sqrt(abs(self.a * self.d - self.b * self.c)) or 1.0

    def descomponer(self):
        """-> (sx, sy, rotación°) si es rotación+escala sin sesgo ni espejo; si no, None."""
        sx = math.hypot(self.a, self.b)
        sy2 = math.hypot(self.c, self.d)
        det = self.a * self.d - self.b * self.c
        if sx < 1e-12 or sy2 < 1e-12 or det <= 0:
            return None
        if abs(self.a * self.c + self.b * self.d) > 1e-4 * sx * sy2:
            return None
        return sx, det / sx, math.degrees(math.atan2(self.b, self.a))


def traslacion(x, y):
    return M(1, 0, 0, 1, x, y)


def escala(sx, sy=None):
    return M(sx, 0, 0, sx if sy is None else sy, 0, 0)


NUM = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"
NUM_RE = re.compile(NUM)


def numeros(s) -> list[float]:
    return [float(x) for x in NUM_RE.findall(s or "")]


def parsear_transform(t: str | None) -> M:
    m = M()
    if not t:
        return m
    for nombre, args in re.findall(r"(\w+)\s*\(([^)]*)\)", t):
        v = numeros(args)
        if nombre == "matrix" and len(v) == 6:
            m = m * M(*v)
        elif nombre == "translate" and v:
            m = m * traslacion(v[0], v[1] if len(v) > 1 else 0)
        elif nombre == "scale" and v:
            m = m * escala(v[0], v[1] if len(v) > 1 else v[0])
        elif nombre == "rotate" and v:
            r = math.radians(v[0])
            rot = M(math.cos(r), math.sin(r), -math.sin(r), math.cos(r), 0, 0)
            if len(v) >= 3:
                m = m * traslacion(v[1], v[2]) * rot * traslacion(-v[1], -v[2])
            else:
                m = m * rot
        elif nombre == "skewX" and v:
            m = m * M(1, 0, math.tan(math.radians(v[0])), 1, 0, 0)
        elif nombre == "skewY" and v:
            m = m * M(1, math.tan(math.radians(v[0])), 0, 1, 0, 0)
    return m


# =================================================================== trazados
class _Lector:
    def __init__(self, s):
        self.s, self.i, self.n = s, 0, len(s)

    def _saltar(self):
        while self.i < self.n and self.s[self.i] in " \t\r\n,":
            self.i += 1

    def comando(self):
        self._saltar()
        if self.i < self.n and self.s[self.i].isalpha() and self.s[self.i] not in "eE":
            c = self.s[self.i]
            self.i += 1
            return c
        return None

    def hay_numero(self):
        self._saltar()
        return self.i < self.n and (self.s[self.i].isdigit() or self.s[self.i] in "+-.")

    def numero(self):
        self._saltar()
        m = NUM_RE.match(self.s, self.i)
        if not m:
            raise ValueError("número esperado")
        self.i = m.end()
        return float(m.group())

    def bandera(self):
        self._saltar()
        c = self.s[self.i]
        self.i += 1
        return c == "1"


def _arco_a_bezier(x1, y1, rx, ry, phi, fa, fs, x2, y2):
    if (x1, y1) == (x2, y2):
        return []
    rx, ry = abs(rx), abs(ry)
    if rx < 1e-9 or ry < 1e-9:
        return [("L", x2, y2)]
    cp, sp = math.cos(math.radians(phi)), math.sin(math.radians(phi))
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    x1p, y1p = cp * dx + sp * dy, -sp * dx + cp * dy
    lam = (x1p ** 2) / rx ** 2 + (y1p ** 2) / ry ** 2
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx ** 2 * ry ** 2 - rx ** 2 * y1p ** 2 - ry ** 2 * x1p ** 2
    den = rx ** 2 * y1p ** 2 + ry ** 2 * x1p ** 2
    coef = math.sqrt(max(0.0, num / den)) if den else 0.0
    if fa == fs:
        coef = -coef
    cxp, cyp = coef * rx * y1p / ry, -coef * ry * x1p / rx
    cx = cp * cxp - sp * cyp + (x1 + x2) / 2
    cy = sp * cxp + cp * cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a

    t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not fs and dt > 0:
        dt -= 2 * math.pi
    elif fs and dt < 0:
        dt += 2 * math.pi
    segs = max(1, int(math.ceil(abs(dt) / (math.pi / 2) - 1e-7)))
    d = dt / segs
    k = 4 / 3 * math.tan(d / 4)
    out = []
    for s in range(segs):
        a1, a2 = t1 + s * d, t1 + (s + 1) * d
        c1, s1, c2, s2 = math.cos(a1), math.sin(a1), math.cos(a2), math.sin(a2)
        pts = [(c1 - k * s1, s1 + k * c1), (c2 + k * s2, s2 - k * c2), (c2, s2)]
        res = []
        for ux, uy in pts:
            px, py = ux * rx, uy * ry
            res += [cp * px - sp * py + cx, sp * px + cp * py + cy]
        out.append(("C", *res))
    return out


def parsear_trazado(d: str) -> list[list[tuple]]:
    """Devuelve subtrazos con comandos absolutos M/L/C/Z."""
    lec = _Lector(d or "")
    subtrazos: list[list[tuple]] = []
    actual: list[tuple] = []
    cx = cy = sx = sy = 0.0
    ultimo_ctrl = None
    ultimo_cmd = ""
    cmd = None
    try:
        while True:
            nuevo = lec.comando()
            if nuevo:
                cmd = nuevo
            elif not lec.hay_numero():
                break
            elif cmd is None:
                break
            rel = cmd.islower()
            C = cmd.upper()
            if C == "Z":
                if actual:
                    actual.append(("Z",))
                cx, cy = sx, sy
                ultimo_cmd, ultimo_ctrl = "Z", None
                if not nuevo:
                    break
                continue
            ox, oy = (cx, cy) if rel else (0.0, 0.0)
            if C == "M":
                x, y = lec.numero() + ox, lec.numero() + oy
                if actual:
                    subtrazos.append(actual)
                actual = [("M", x, y)]
                cx, cy, sx, sy = x, y, x, y
                cmd = "l" if rel else "L"
                ultimo_ctrl = None
            elif C == "L":
                cx, cy = lec.numero() + ox, lec.numero() + oy
                actual.append(("L", cx, cy)); ultimo_ctrl = None
            elif C == "H":
                cx = lec.numero() + ox
                actual.append(("L", cx, cy)); ultimo_ctrl = None
            elif C == "V":
                cy = lec.numero() + oy
                actual.append(("L", cx, cy)); ultimo_ctrl = None
            elif C == "C":
                x1, y1 = lec.numero() + ox, lec.numero() + oy
                x2, y2 = lec.numero() + ox, lec.numero() + oy
                x, y = lec.numero() + ox, lec.numero() + oy
                actual.append(("C", x1, y1, x2, y2, x, y))
                ultimo_ctrl, cx, cy = (x2, y2), x, y
            elif C == "S":
                if ultimo_ctrl and ultimo_cmd in "CS":
                    x1, y1 = 2 * cx - ultimo_ctrl[0], 2 * cy - ultimo_ctrl[1]
                else:
                    x1, y1 = cx, cy
                x2, y2 = lec.numero() + ox, lec.numero() + oy
                x, y = lec.numero() + ox, lec.numero() + oy
                actual.append(("C", x1, y1, x2, y2, x, y))
                ultimo_ctrl, cx, cy = (x2, y2), x, y
            elif C in "QT":
                if C == "Q":
                    qx, qy = lec.numero() + ox, lec.numero() + oy
                else:
                    if ultimo_ctrl and ultimo_cmd in "QT":
                        qx, qy = 2 * cx - ultimo_ctrl[0], 2 * cy - ultimo_ctrl[1]
                    else:
                        qx, qy = cx, cy
                x, y = lec.numero() + ox, lec.numero() + oy
                actual.append(("C", cx + 2 / 3 * (qx - cx), cy + 2 / 3 * (qy - cy),
                               x + 2 / 3 * (qx - x), y + 2 / 3 * (qy - y), x, y))
                ultimo_ctrl, cx, cy = (qx, qy), x, y
            elif C == "A":
                rx, ry, phi = lec.numero(), lec.numero(), lec.numero()
                fa, fs = lec.bandera(), lec.bandera()
                x, y = lec.numero() + ox, lec.numero() + oy
                actual += _arco_a_bezier(cx, cy, rx, ry, phi, fa, fs, x, y)
                cx, cy, ultimo_ctrl = x, y, None
            else:
                break
            ultimo_cmd = C
            if not actual:
                actual = [("M", cx, cy)]
    except (ValueError, IndexError):
        pass
    if actual:
        subtrazos.append(actual)
    return [s for s in subtrazos if len(s) > 1]


K = 0.5522847498


def trazado_elipse(cx, cy, rx, ry):
    return [[("M", cx + rx, cy),
             ("C", cx + rx, cy + K * ry, cx + K * rx, cy + ry, cx, cy + ry),
             ("C", cx - K * rx, cy + ry, cx - rx, cy + K * ry, cx - rx, cy),
             ("C", cx - rx, cy - K * ry, cx - K * rx, cy - ry, cx, cy - ry),
             ("C", cx + K * rx, cy - ry, cx + rx, cy - K * ry, cx + rx, cy),
             ("Z",)]]


def trazado_rect(x, y, w, h, rx=0.0, ry=0.0):
    if rx <= 0 and ry <= 0:
        return [[("M", x, y), ("L", x + w, y), ("L", x + w, y + h), ("L", x, y + h), ("Z",)]]
    rx, ry = min(rx, w / 2), min(ry, h / 2)
    kx, ky = K * rx, K * ry
    return [[("M", x + rx, y), ("L", x + w - rx, y),
             ("C", x + w - rx + kx, y, x + w, y + ry - ky, x + w, y + ry),
             ("L", x + w, y + h - ry),
             ("C", x + w, y + h - ry + ky, x + w - rx + kx, y + h, x + w - rx, y + h),
             ("L", x + rx, y + h),
             ("C", x + rx - kx, y + h, x, y + h - ry + ky, x, y + h - ry),
             ("L", x, y + ry),
             ("C", x, y + ry - ky, x + rx - kx, y, x + rx, y),
             ("Z",)]]


def transformar(subtrazos, m: M):
    out = []
    for sub in subtrazos:
        ns = []
        for c in sub:
            if c[0] == "Z":
                ns.append(c)
            elif c[0] == "C":
                p1, p2, p3 = m.ap(c[1], c[2]), m.ap(c[3], c[4]), m.ap(c[5], c[6])
                ns.append(("C", *p1, *p2, *p3))
            else:
                ns.append((c[0], *m.ap(c[1], c[2])))
        out.append(ns)
    return out


def caja(subtrazos):
    """Caja envolvente exacta (muestrea las curvas)."""
    xs, ys = [], []
    for sub in subtrazos:
        px = py = None
        for c in sub:
            if c[0] in "ML":
                px, py = c[1], c[2]
                xs.append(px); ys.append(py)
            elif c[0] == "C":
                x0, y0 = (px, py) if px is not None else (c[5], c[6])
                for i in range(1, 17):
                    t = i / 16
                    u = 1 - t
                    xs.append(u**3 * x0 + 3 * u * u * t * c[1] + 3 * u * t * t * c[3] + t**3 * c[5])
                    ys.append(u**3 * y0 + 3 * u * u * t * c[2] + 3 * u * t * t * c[4] + t**3 * c[6])
                px, py = c[5], c[6]
    if not xs:
        return None
    return min(xs), min(ys), max(xs), max(ys)


# =================================================================== colores
def parsear_color(v: str, color_actual=(0, 0, 0)):
    """-> (rgb, alfa) o None."""
    v = (v or "").strip()
    if not v or v == "none" or v == "transparent":
        return None
    if v == "currentColor":
        return color_actual, 1.0
    try:
        if v.startswith("#") and len(v) in (5, 9):  # #rgba / #rrggbbaa
            h = v[1:]
            if len(h) == 4:
                h = "".join(c * 2 for c in h)
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)), int(h[6:8], 16) / 255
        m = re.match(r"rgba?\(([^)]*)\)", v)
        if m:
            partes = [p.strip() for p in re.split(r"[,\s/]+", m.group(1)) if p.strip()]
            rgb = tuple(int(float(p[:-1]) * 2.55) if p.endswith("%") else int(float(p)) for p in partes[:3])
            a = 1.0
            if len(partes) > 3:
                a = float(partes[3][:-1]) / 100 if partes[3].endswith("%") else float(partes[3])
            return rgb, a
        r = ImageColor.getcolor(v, "RGBA")
        return r[:3], r[3] / 255
    except (ValueError, IndexError):
        return None


def _mezclar(c1, c2, t):
    return tuple(c1[i] + (c2[i] - c1[i]) * t for i in range(3))


# =================================================================== CSS
class HojaCSS:
    def __init__(self):
        self.reglas: list[tuple[tuple, int, str, dict]] = []  # (especificidad, orden, selector, decl)

    def agregar(self, css: str):
        css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
        css = re.sub(r"@[\w-]+[^{;]*\{(?:[^{}]*\{[^{}]*\})*[^{}]*\}", "", css)  # @media/@font-face
        css = re.sub(r"@[^;{]*;", "", css)
        for selectores, cuerpo in re.findall(r"([^{}]+)\{([^{}]*)\}", css):
            decl = parsear_declaraciones(cuerpo)
            for sel in selectores.split(","):
                sel = sel.strip()
                if not sel or " " in sel or ">" in sel or ":" in sel:
                    continue
                esp = (sel.count("#"), sel.count(".") + sel.count("["), 0 if sel[0] in ".#*" else 1)
                self.reglas.append((esp, len(self.reglas), sel, decl))
        self.reglas.sort(key=lambda r: (r[0], r[1]))

    @staticmethod
    def _coincide(sel: str, tag: str, clases: set, ident: str | None) -> bool:
        if sel == "*":
            return True
        partes = re.findall(r"([.#]?)([\w-]+)", sel)
        for pref, nombre in partes:
            if pref == "." and nombre not in clases:
                return False
            if pref == "#" and nombre != ident:
                return False
            if pref == "" and nombre != tag:
                return False
        return bool(partes)

    def estilos(self, tag, clases, ident) -> dict:
        res = {}
        for _, _, sel, decl in self.reglas:
            if self._coincide(sel, tag, clases, ident):
                res.update(decl)
        return res


def parsear_declaraciones(s: str) -> dict:
    res = {}
    for parte in (s or "").split(";"):
        if ":" in parte:
            k, v = parte.split(":", 1)
            res[k.strip().lower()] = v.replace("!important", "").strip()
    return res


def local(tag) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.split("}", 1)[-1]


# =================================================================== conversor
@dataclass
class Contexto:
    m: M
    estilo: dict
    opacidad: float
    contenedor: object
    recorte: tuple | None = None  # (x0,y0,x1,y1) en px de diapositiva


class ConversorSVG:
    def __init__(self, catalogo: CatalogoFuentes, informe: Informe | None = None):
        self.cat = catalogo
        self.inf = informe or Informe()

    # ----------------------------------------------------------- utilidades
    @staticmethod
    def cargar(svg: bytes | str):
        if isinstance(svg, str):
            svg = svg.encode("utf-8")
        parser = etree.XMLParser(remove_comments=True, recover=True, huge_tree=True,
                                 resolve_entities=False, no_network=True)
        raiz = etree.fromstring(svg, parser)
        if raiz is None or local(raiz.tag) != "svg":
            raise ValueError("El archivo no parece ser un SVG válido.")
        return raiz

    @staticmethod
    def longitud(v, ref=100.0, fs=16.0, defecto=0.0) -> float:
        if v is None:
            return defecto
        v = str(v).strip()
        m = re.match(r"(" + NUM + r")\s*([a-z%]*)", v)
        if not m:
            return defecto
        n, u = float(m.group(1)), m.group(2)
        return n * {"": 1, "px": 1, "pt": 4 / 3, "pc": 16, "mm": 96 / 25.4, "cm": 96 / 2.54,
                    "in": 96, "em": fs, "ex": fs / 2, "rem": 16, "%": ref / 100}.get(u, 1)

    @classmethod
    def dimensiones(cls, raiz) -> tuple[float, float, M]:
        """Tamaño del lienzo en px y matriz viewBox -> lienzo."""
        vb = numeros(raiz.get("viewBox"))
        w_attr, h_attr = raiz.get("width"), raiz.get("height")
        if len(vb) == 4 and vb[2] > 0 and vb[3] > 0:
            w = cls.longitud(w_attr, vb[2]) if w_attr and "%" not in w_attr else vb[2]
            h = cls.longitud(h_attr, vb[3]) if h_attr and "%" not in h_attr else vb[3]
            m = cls._matriz_viewbox(vb, w, h, raiz.get("preserveAspectRatio"))
            return w, h, m
        w = cls.longitud(w_attr, 0, defecto=0) or 800
        h = cls.longitud(h_attr, 0, defecto=0) or 600
        return w, h, M()

    @staticmethod
    def _matriz_viewbox(vb, w, h, par=None) -> M:
        sx, sy = w / vb[2], h / vb[3]
        par = (par or "xMidYMid meet").split()
        alin = par[0] if par else "xMidYMid"
        if alin == "none":
            return escala(sx, sy) * traslacion(-vb[0], -vb[1])
        s = max(sx, sy) if len(par) > 1 and par[1] == "slice" else min(sx, sy)
        tx = {"xMin": 0, "xMid": 0.5, "xMax": 1}.get(alin[:4], 0.5) * (w - vb[2] * s)
        ty = {"YMin": 0, "YMid": 0.5, "YMax": 1}.get(alin[4:], 0.5) * (h - vb[3] * s)
        return traslacion(tx, ty) * escala(s) * traslacion(-vb[0], -vb[1])

    # ----------------------------------------------------------- entrada
    def convertir(self, svg: bytes | str, diapositiva, matriz_lienzo: M,
                  carpeta_base: str | Path | None = None):
        """Dibuja el SVG en `diapositiva`. `matriz_lienzo` lleva px del lienzo SVG
        a px de diapositiva (escala y centrado)."""
        raiz = self.cargar(svg)
        self.base = Path(carpeta_base) if carpeta_base else None
        self.ids = {e.get("id"): e for e in raiz.iter() if isinstance(e.tag, str) and e.get("id")}
        self.css = HojaCSS()
        for e in raiz.iter():
            if local(e.tag) == "style":
                self.css.agregar("".join(e.itertext()))
        _, _, m_vb = self.dimensiones(raiz)
        estilo_raiz = {"fill": "#000000", "stroke": "none", "stroke-width": "1",
                       "font-size": "16", "font-family": "sans-serif", "color": "#000000"}
        ctx = Contexto(matriz_lienzo * m_vb, estilo_raiz, 1.0, diapositiva.shapes)
        self._hijos(raiz, ctx)

    # ----------------------------------------------------------- estilos
    def _estilo(self, el, padre: dict) -> dict:
        est = {k: v for k, v in padre.items() if k in HEREDADAS}
        for k in PROPIEDADES:
            v = el.get(k)
            if v is not None:
                est[k] = v.strip()
        if el.get("{http://www.w3.org/XML/1998/namespace}space"):
            est["xml:space"] = el.get("{http://www.w3.org/XML/1998/namespace}space")
        clases = set((el.get("class") or "").split())
        est.update({k: v for k, v in self.css.estilos(local(el.tag), clases, el.get("id")).items()
                    if k in PROPIEDADES})
        est.update({k: v for k, v in parsear_declaraciones(el.get("style")).items() if k in PROPIEDADES})
        # tamaños relativos de fuente
        if "font-size" in est:
            fs_padre = self.longitud(padre.get("font-size"), defecto=16)
            v = est["font-size"]
            if v.endswith("%") or v.endswith("em"):
                est["font-size"] = str(self.longitud(v, fs_padre, fs_padre, fs_padre))
        for k, v in list(est.items()):
            if v == "inherit":
                if k in padre:
                    est[k] = padre[k]
                else:
                    del est[k]
        return est

    def _num(self, est, k, defecto):
        try:
            v = est.get(k)
            if v is None:
                return defecto
            v = v.strip()
            return float(v[:-1]) / 100 if v.endswith("%") else float(v)
        except ValueError:
            return defecto

    def _pintura(self, est: dict, clave: str, opacidad: float):
        """-> None | ('solido', rgb, alfa) | ('degradado', elemento, alfa)"""
        v = (est.get(clave) or "").strip()
        alfa = opacidad * self._num(est, f"{clave}-opacity", 1.0)
        if v.startswith("url("):
            m = re.match(r"url\(\s*['\"]?#([^)'\"]+)['\"]?\s*\)\s*(.*)", v)
            if m and m.group(1) in self.ids:
                ref = self.ids[m.group(1)]
                if local(ref.tag) in ("linearGradient", "radialGradient"):
                    return ("degradado", ref, alfa)
                self.inf.aviso("Los rellenos de tipo patrón (pattern) se reemplazan por un color plano.")
                return ("solido", (128, 128, 128), alfa)
            v = m.group(2) if m else ""
        color_actual = (parsear_color(est.get("color", "#000")) or ((0, 0, 0), 1))[0]
        c = parsear_color(v, color_actual)
        if c is None:
            return None
        return ("solido", c[0], alfa * c[1])

    # ----------------------------------------------------------- degradados
    def _degradado_datos(self, el):
        """Recopila atributos y paradas siguiendo la herencia por href."""
        cadena, vistos = [], set()
        e = el
        while e is not None and id(e) not in vistos:
            vistos.add(id(e))
            cadena.append(e)
            h = e.get(XLINK) or e.get("href")
            e = self.ids.get(h[1:]) if h and h.startswith("#") else None
        attrs = {}
        for e in reversed(cadena):
            attrs.update({k: v for k, v in e.attrib.items() if isinstance(k, str)})
        paradas = []
        for e in cadena:
            ps = [s for s in e if local(s.tag) == "stop"]
            if ps:
                for s in ps:
                    est = self._estilo(s, {})
                    off = s.get("offset", "0").strip()
                    o = float(off[:-1]) / 100 if off.endswith("%") else float(numeros(off)[0] if numeros(off) else 0)
                    c = parsear_color(est.get("stop-color", "#000")) or ((0, 0, 0), 1.0)
                    a = self._num(est, "stop-opacity", 1.0) * c[1]
                    o = max(paradas[-1][0] if paradas else 0.0, min(1.0, max(0.0, o)))
                    paradas.append((o, c[0], a))
                break
        return local(el.tag), attrs, paradas

    def _relleno_degradado(self, ref, alfa, caja_usuario, m: M, caja_diapo, rotacion=0.0):
        tipo, at, paradas = self._degradado_datos(ref)
        if not paradas:
            return pu.sin_relleno()
        if len(paradas) == 1:
            return pu.relleno_solido(paradas[0][1], paradas[0][2] * alfa)
        bx, by, bx1, by1 = caja_usuario
        bw, bh = max(bx1 - bx, 1e-9), max(by1 - by, 1e-9)
        bbox = at.get("gradientUnits", "objectBoundingBox") != "userSpaceOnUse"
        mt = m * (M(bw, 0, 0, bh, bx, by) if bbox else M()) * parsear_transform(at.get("gradientTransform"))
        ref_w, ref_h = (1, 1) if bbox else (bw, bh)

        def L(k, d):
            return self.longitud(at.get(k), ref_w if k in ("x1", "x2", "cx", "fx") else ref_h, defecto=d)

        X0, Y0, X1, Y1 = caja_diapo
        cxc, cyc = (X0 + X1) / 2, (Y0 + Y1) / 2
        rad = math.radians(-rotacion)

        def a_local(p):  # deshace la rotación de la forma
            dx, dy = p[0] - cxc, p[1] - cyc
            return (cxc + dx * math.cos(rad) - dy * math.sin(rad),
                    cyc + dx * math.sin(rad) + dy * math.cos(rad))

        def color_en(t):  # t en 0..1 a lo largo del degradado SVG
            if t <= paradas[0][0]:
                return paradas[0][1], paradas[0][2]
            for (o1, c1, a1), (o2, c2, a2) in zip(paradas, paradas[1:]):
                if t <= o2:
                    f = (t - o1) / (o2 - o1) if o2 > o1 else 0
                    return _mezclar(c1, c2, f), a1 + (a2 - a1) * f
            return paradas[-1][1], paradas[-1][2]

        if tipo == "linearGradient":
            p1 = a_local(mt.ap(L("x1", 0), L("y1", 0)))
            p2 = a_local(mt.ap(L("x2", ref_w if bbox else 1), L("y2", 0)))
            if at.get("x2") is None and bbox:
                p2 = a_local(mt.ap(1, 0))
            vx, vy = p2[0] - p1[0], p2[1] - p1[1]
            largo = math.hypot(vx, vy) or 1e-9
            ux, uy = vx / largo, vy / largo
            esquinas = [(X0, Y0), (X1, Y0), (X0, Y1), (X1, Y1)]
            proy = [(x - p1[0]) * ux + (y - p1[1]) * uy for x, y in esquinas]
            tmin, tmax = min(proy), max(proy)
            span = (tmax - tmin) or 1e-9
            angulo = math.degrees(math.atan2(uy, ux))
            out = []
            c, a = color_en(tmin / largo)
            out.append((0.0, c, a * alfa))
            for o, c, a in paradas:
                pos = (o * largo - tmin) / span
                if 0 < pos < 1:
                    out.append((pos, c, a * alfa))
            c, a = color_en(tmax / largo)
            out.append((1.0, c, a * alfa))
            return pu.relleno_degradado(out, "lineal", angulo)

        # radial
        cx, cy = L("cx", 0.5 if bbox else 0), L("cy", 0.5 if bbox else 0)
        r = L("r", 0.5 if bbox else 0)
        c = a_local(mt.ap(cx, cy))
        rx = math.dist(c, a_local(mt.ap(cx + r, cy)))
        ry = math.dist(c, a_local(mt.ap(cx, cy + r)))
        R = (rx + ry) / 2 or 1e-9
        W, H = max(X1 - X0, 1e-9), max(Y1 - Y0, 1e-9)
        centro = (min(1, max(0, (c[0] - X0) / W)), min(1, max(0, (c[1] - Y0) / H)))
        # PowerPoint: el círculo final pasa por la esquina más lejana del rectángulo
        R_ppt = max(math.dist(c, e) for e in [(X0, Y0), (X1, Y0), (X0, Y1), (X1, Y1)]) or 1e-9
        out = []
        for o, col, a in paradas:
            pos = o * R / R_ppt
            if pos <= 1:
                out.append((pos, col, a * alfa))
        if not out or out[-1][0] < 1:
            col, a = color_en(R_ppt / R)
            out.append((1.0, col, a * alfa))
        if out[0][0] > 0:
            out.insert(0, (0.0, paradas[0][1], paradas[0][2] * alfa))
        return pu.relleno_degradado(out, "radial", centro=centro)

    def _relleno(self, pintura, caja_usuario, m, caja_diapo, rotacion=0.0):
        if pintura is None:
            return pu.sin_relleno()
        if pintura[0] == "solido":
            return pu.relleno_solido(pintura[1], pintura[2])
        return self._relleno_degradado(pintura[1], pintura[2], caja_usuario, m, caja_diapo, rotacion)

    def _linea(self, est, ctx: Contexto, caja_usuario, caja_diapo, rotacion=0.0):
        p = self._pintura(est, "stroke", ctx.opacidad)
        ancho = self.longitud(est.get("stroke-width", "1"), defecto=1) * ctx.m.escala
        if p is None or ancho <= 0:
            return pu.sin_linea()
        if p[0] == "degradado":
            rel = self._relleno_degradado(p[1], p[2], caja_usuario, ctx.m, caja_diapo, rotacion)
        else:
            rel = pu.relleno_solido(p[1], p[2])
        guiones = None
        da = est.get("stroke-dasharray")
        if da and da != "none":
            guiones = [g * ctx.m.escala for g in numeros(da)] or None
        return pu.linea(pu.emu(ancho), rel, est.get("stroke-linecap", "butt"),
                        est.get("stroke-linejoin", "miter"), guiones)

    # ----------------------------------------------------------- recorrido
    def _hijos(self, el, ctx: Contexto):
        for h in el:
            if isinstance(h.tag, str):
                self._elemento(h, ctx)

    def _recorte_de(self, est, m: M, recorte):
        cp = est.get("clip-path")
        if not cp or cp == "none":
            return recorte
        mm = re.match(r"url\(\s*['\"]?#([^)'\"]+)", cp)
        clip = self.ids.get(mm.group(1)) if mm else None
        if clip is None:
            return recorte
        formas = [c for c in clip if isinstance(c.tag, str) and local(c.tag) in ("rect", "path", "polygon", "circle", "ellipse", "use")]
        if len(formas) == 1 and local(formas[0].tag) == "rect":
            r = formas[0]
            mc = m * parsear_transform(clip.get("transform")) * parsear_transform(r.get("transform"))
            x, y = self.longitud(r.get("x")), self.longitud(r.get("y"))
            w, h = self.longitud(r.get("width")), self.longitud(r.get("height"))
            pts = [mc.ap(x, y), mc.ap(x + w, y), mc.ap(x, y + h), mc.ap(x + w, y + h)]
            nuevo = (min(p[0] for p in pts), min(p[1] for p in pts),
                     max(p[0] for p in pts), max(p[1] for p in pts))
            if recorte:
                nuevo = (max(nuevo[0], recorte[0]), max(nuevo[1], recorte[1]),
                         min(nuevo[2], recorte[2]), min(nuevo[3], recorte[3]))
            return nuevo
        self.inf.aviso("Algunas máscaras de recorte no rectangulares se ignoraron "
                       "(los elementos se muestran completos).")
        return recorte

    def _elemento(self, el, ctx: Contexto, forzar=False):
        tag = local(el.tag)
        if tag in IGNORAR and not forzar:
            return
        est = self._estilo(el, ctx.estilo)
        if est.get("display") == "none" or est.get("visibility") in ("hidden", "collapse"):
            return
        if est.get("filter") and est["filter"] != "none":
            self.inf.aviso("Los efectos SVG (sombras, desenfoques) no se convierten; "
                           "puedes añadirlos en PowerPoint con Formato de forma > Efectos.")
        if est.get("mask") and est["mask"] != "none":
            self.inf.aviso("Las máscaras de opacidad (mask) se ignoraron.")
        m = ctx.m * parsear_transform(el.get("transform") if el.get("transform") is not None
                                      else est.get("transform"))
        op = ctx.opacidad * self._num(est, "opacity", 1.0)
        nuevo = Contexto(m, est, op, ctx.contenedor, self._recorte_de(est, m, ctx.recorte))

        if tag in ("g", "a", "switch", "symbol"):
            self._grupo(el, nuevo)
        elif tag == "svg":
            x, y = self.longitud(el.get("x")), self.longitud(el.get("y"))
            vb = numeros(el.get("viewBox"))
            w = self.longitud(el.get("width"), defecto=vb[2] if len(vb) == 4 else 0)
            h = self.longitud(el.get("height"), defecto=vb[3] if len(vb) == 4 else 0)
            nuevo.m = m * traslacion(x, y)
            if len(vb) == 4 and vb[2] and vb[3] and w and h:
                nuevo.m = nuevo.m * self._matriz_viewbox(vb, w, h, el.get("preserveAspectRatio"))
            self._grupo(el, nuevo)
        elif tag == "use":
            self._use(el, nuevo)
        elif tag in ("rect", "circle", "ellipse", "line", "polyline", "polygon", "path"):
            self._forma(el, tag, nuevo)
        elif tag == "text":
            self._texto(el, nuevo)
        elif tag == "image":
            self._imagen(el, nuevo)
        elif tag == "foreignObject":
            self.inf.aviso("El contenido <foreignObject> (HTML incrustado) no se puede convertir.")

    def _nombre(self, el, defecto):
        return (el.get("{http://www.inkscape.org/namespaces/inkscape}label") or el.get("data-name")
                or el.get("id") or defecto)

    def _grupo(self, el, ctx: Contexto):
        hijos = [h for h in el if isinstance(h.tag, str) and local(h.tag) not in IGNORAR]
        if len(hijos) > 1:
            grupo = ctx.contenedor.add_group_shape()
            grupo.name = self._nombre(el, "Grupo")
            ctx.contenedor = grupo.shapes
            self._hijos(el, ctx)
            pu.desagrupar_si_trivial(grupo)
        else:
            self._hijos(el, ctx)

    def _use(self, el, ctx: Contexto):
        h = el.get(XLINK) or el.get("href") or ""
        ref = self.ids.get(h[1:]) if h.startswith("#") else None
        if ref is None:
            return
        x, y = self.longitud(el.get("x")), self.longitud(el.get("y"))
        ctx.m = ctx.m * traslacion(x, y)
        if local(ref.tag) == "symbol":
            vb = numeros(ref.get("viewBox"))
            w = self.longitud(el.get("width"), defecto=vb[2] if len(vb) == 4 else 0)
            hh = self.longitud(el.get("height"), defecto=vb[3] if len(vb) == 4 else 0)
            if len(vb) == 4 and w and hh:
                ctx.m = ctx.m * self._matriz_viewbox(vb, w, hh, ref.get("preserveAspectRatio"))
        self._elemento(ref, ctx, forzar=True)

    # ----------------------------------------------------------- formas
    def _forma(self, el, tag, ctx: Contexto):
        L = self.longitud
        est = ctx.estilo
        nativa = None   # (prst, x, y, w, h, ajustes)
        if tag == "rect":
            x, y, w, h = L(el.get("x")), L(el.get("y")), L(el.get("width")), L(el.get("height"))
            if w <= 0 or h <= 0:
                return
            rx_a, ry_a = el.get("rx", est.get("rx")), el.get("ry", est.get("ry"))
            rx = L(rx_a if rx_a not in (None, "auto") else ry_a)
            ry = L(ry_a if ry_a not in (None, "auto") else rx_a)
            sub = trazado_rect(x, y, w, h, rx, ry)
            if rx <= 0 and ry <= 0:
                nativa = ("rect", x, y, w, h, None)
            elif abs(min(rx, w / 2) - min(ry, h / 2)) < 0.01:
                r = min(rx, w / 2, h / 2)
                nativa = ("roundRect", x, y, w, h, {"adj": r / min(w, h) * 100000})
        elif tag in ("circle", "ellipse"):
            cx, cy = L(el.get("cx")), L(el.get("cy"))
            if tag == "circle":
                rx = ry = L(el.get("r"))
            else:
                rx, ry = L(el.get("rx")), L(el.get("ry"))
            if rx <= 0 or ry <= 0:
                return
            sub = trazado_elipse(cx, cy, rx, ry)
            nativa = ("ellipse", cx - rx, cy - ry, 2 * rx, 2 * ry, None)
        elif tag == "line":
            sub = [[("M", L(el.get("x1")), L(el.get("y1"))), ("L", L(el.get("x2")), L(el.get("y2")))]]
        elif tag in ("polyline", "polygon"):
            v = numeros(el.get("points"))
            pts = list(zip(v[0::2], v[1::2]))
            if len(pts) < 2:
                return
            sub = [[("M", *pts[0])] + [("L", *p) for p in pts[1:]] + ([("Z",)] if tag == "polygon" else [])]
        else:
            sub = parsear_trazado(el.get("d"))
        if not sub:
            return

        caja_u = caja(sub)
        pintura = None if tag in ("line",) else self._pintura(est, "fill", ctx.opacidad)
        trazo_p = self._pintura(est, "stroke", ctx.opacidad)
        if pintura is None and trazo_p is None:
            return
        nombre = self._nombre(el, {"rect": "Rectángulo", "circle": "Círculo", "ellipse": "Elipse",
                                   "line": "Línea", "text": "Texto"}.get(tag, "Forma"))

        dec = ctx.m.descomponer()
        if nativa and dec:
            prst, x, y, w, h, aj = nativa
            sx, sy, rot = dec
            cx, cy = ctx.m.ap(x + w / 2, y + h / 2)
            W, H = w * sx, h * sy
            caja_d = (cx - W / 2, cy - H / 2, cx + W / 2, cy + H / 2)
            if prst == "roundRect" and abs(sx - sy) > 1e-3:
                prst = None  # radio deformado: mejor forma libre
            if prst:
                f = pu.nueva_forma(ctx.contenedor, caja_d[0], caja_d[1], W, H, nombre)
                if abs(rot) > 0.01:
                    f.rotation = rot % 360
                pu.aplicar_propiedades(
                    f, pu.geometria_preset(prst, aj),
                    self._relleno(pintura, caja_u, ctx.m, caja_d, rot),
                    self._linea(est, ctx, caja_u, caja_d, rot))
                self.inf.contar("formas")
                return

        subd = transformar(sub, ctx.m)
        cd = caja(subd)
        if not cd:
            return
        x0, y0, x1, y1 = cd
        f = pu.nueva_forma(ctx.contenedor, x0, y0, x1 - x0, y1 - y0, nombre)
        pu.aplicar_propiedades(
            f, pu.geometria_libre(subd, x0, y0, x1 - x0, y1 - y0),
            self._relleno(pintura, caja_u, ctx.m, cd),
            self._linea(est, ctx, caja_u, cd))
        self.inf.contar("formas")

    # ----------------------------------------------------------- imágenes
    def _imagen(self, el, ctx: Contexto):
        href = el.get(XLINK) or el.get("href") or ""
        datos = None
        try:
            if href.startswith("data:"):
                cab, _, carga = href.partition(",")
                if "svg" in cab:
                    self.inf.aviso("Hay imágenes SVG incrustadas dentro del SVG; se omitieron.")
                    return
                datos = base64.b64decode(carga) if ";base64" in cab else unquote(carga).encode("latin-1")
            elif href and self.base is not None:
                ruta = (self.base / unquote(href)).resolve()
                if ruta.is_file():
                    datos = ruta.read_bytes()
        except Exception:
            datos = None
        if not datos:
            self.inf.aviso(f"No se encontró una imagen enlazada ({href[:60]}). "
                           "Exporta el SVG con las imágenes incrustadas.")
            return
        try:
            img = Image.open(io.BytesIO(datos))
            iw, ih = img.size
            if img.format not in ("PNG", "JPEG", "GIF", "BMP", "TIFF"):
                buf = io.BytesIO()
                img.save(buf, "PNG")
                datos = buf.getvalue()
        except Exception:
            self.inf.aviso("Una imagen incrustada tiene un formato no reconocido y se omitió.")
            return
        L = self.longitud
        x, y = L(el.get("x")), L(el.get("y"))
        w, h = L(el.get("width"), defecto=iw), L(el.get("height"), defecto=ih)
        par = (el.get("preserveAspectRatio") or "xMidYMid meet").split()
        if par[0] != "none" and iw and ih:
            s = (max if len(par) > 1 and par[1] == "slice" else min)(w / iw, h / ih)
            nw, nh = iw * s, ih * s
            fx = {"xMin": 0, "xMid": .5, "xMax": 1}.get(par[0][:4], .5)
            fy = {"YMin": 0, "YMid": .5, "YMax": 1}.get(par[0][4:], .5)
            if len(par) > 1 and par[1] == "slice":
                # recorte implícito en la caja del <image>
                caja_img = ctx.m.ap(x, y) + ctx.m.ap(x + w, y + h)
                ctx = Contexto(ctx.m, ctx.estilo, ctx.opacidad, ctx.contenedor,
                               self._intersecar(ctx.recorte, (min(caja_img[0], caja_img[2]), min(caja_img[1], caja_img[3]),
                                                              max(caja_img[0], caja_img[2]), max(caja_img[1], caja_img[3]))))
            x, y, w, h = x + (w - nw) * fx, y + (h - nh) * fy, nw, nh
        dec = ctx.m.descomponer()
        if dec:
            sx, sy, rot = dec
        else:
            sx = sy = ctx.m.escala
            rot = 0.0
            self.inf.aviso("Algunas imágenes tenían sesgo o espejo; se colocaron sin esa deformación.")
        cx, cy = ctx.m.ap(x + w / 2, y + h / 2)
        W, H = w * sx, h * sy
        X0, Y0 = cx - W / 2, cy - H / 2
        pic = ctx.contenedor.add_picture(io.BytesIO(datos), Emu(pu.emu(X0)), Emu(pu.emu(Y0)),
                                         Emu(max(1, pu.emu(W))), Emu(max(1, pu.emu(H))))
        pic.name = self._nombre(el, "Imagen")
        if abs(rot) > 0.01:
            pic.rotation = rot % 360
        elif ctx.recorte:
            rx0, ry0, rx1, ry1 = ctx.recorte
            izq = max(0.0, (rx0 - X0) / W)
            arr = max(0.0, (ry0 - Y0) / H)
            der = max(0.0, (X0 + W - rx1) / W)
            aba = max(0.0, (Y0 + H - ry1) / H)
            if izq + der < 1 and arr + aba < 1 and (izq or arr or der or aba):
                pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = izq, arr, der, aba
                pic.left = Emu(pu.emu(X0 + izq * W))
                pic.top = Emu(pu.emu(Y0 + arr * H))
                pic.width = Emu(pu.emu(W * (1 - izq - der)))
                pic.height = Emu(pu.emu(H * (1 - arr - aba)))
        if ctx.opacidad < 0.999:
            blip = pic._element.find(".//" + "{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
            if blip is not None:
                a = etree.SubElement(blip, "{http://schemas.openxmlformats.org/drawingml/2006/main}alphaModFix")
                a.set("amt", str(int(ctx.opacidad * 100000)))
        self.inf.contar("imágenes")

    @staticmethod
    def _intersecar(a, b):
        if not a:
            return b
        return (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))

    # ----------------------------------------------------------- texto
    def _fuente(self, est):
        peso = parsear_peso(est.get("font-weight"))
        cursiva = est.get("font-style", "normal") in ("italic", "oblique")
        res = self.cat.resolver(est.get("font-family", "sans-serif"), peso, cursiva)
        clave = res.nombre + (" (negrita)" if res.negrita else "") + (" (cursiva)" if res.cursiva else "")
        self.inf.fuentes[clave] = res.instalada
        return res

    def _texto(self, el, ctx: Contexto):
        L = self.longitud
        trozos: list[dict] = []
        pos = {"x": 0.0, "y": 0.0}

        def recoger(nodo, est, preservar):
            xs, ys = numeros(nodo.get("x")), numeros(nodo.get("y"))
            dxs, dys = numeros(nodo.get("dx")), numeros(nodo.get("dy"))
            nuevo = bool(xs or ys)
            if xs:
                pos["x"] = xs[0]
            if ys:
                pos["y"] = ys[0]
            if dxs:
                pos["x"] += dxs[0]
                nuevo = nuevo or abs(dxs[0]) > 0
            if dys:
                pos["y"] += dys[0]
                nuevo = nuevo or abs(dys[0]) > 0
            if nuevo or not trozos:
                if trozos and not xs:
                    # continúa donde terminó el trozo anterior
                    t = trozos[-1]
                    pos["x"] = t["x"] + sum(self._medir(r) for r in t["runs"]) * (1 if not dxs else 1)
                    if dxs:
                        pos["x"] += dxs[0]
                trozos.append({"x": pos["x"], "y": pos["y"], "anchor": est.get("text-anchor", "start"),
                               "runs": []})

            def anadir(texto, e):
                if not texto:
                    return
                texto = texto.replace("\t", " ").replace("\r", "").replace("\n", " " if preservar else " ")
                if not preservar:
                    texto = re.sub(r" +", " ", texto)
                trozos[-1]["runs"].append([texto, e])

            anadir(nodo.text, est)
            for h in nodo:
                if not isinstance(h.tag, str):
                    continue
                if local(h.tag) in ("tspan", "a", "textPath", "tref"):
                    eh = self._estilo(h, est)
                    if eh.get("display") != "none":
                        recoger(h, eh, preservar or eh.get("xml:space") == "preserve"
                                or eh.get("white-space") in ("pre", "pre-wrap"))
                anadir(h.tail, est)

        preservar = ctx.estilo.get("xml:space") == "preserve" or ctx.estilo.get("white-space") in ("pre", "pre-wrap")
        recoger(el, ctx.estilo, preservar)

        # limpiar espacios al inicio/fin de cada trozo
        for t in trozos:
            runs = t["runs"]
            if runs and not preservar:
                runs[0][0] = runs[0][0].lstrip()
                runs[-1][0] = runs[-1][0].rstrip()
            t["runs"] = [r for r in runs if r[0]]
        trozos = [t for t in trozos if t["runs"]]
        if not trozos:
            return

        # trozos en la misma línea base -> misma línea
        lineas = []
        for t in trozos:
            if lineas and abs(t["y"] - lineas[-1]["y"]) < 0.01 * max(1, self._tam(t["runs"][0][1])):
                ant = lineas[-1]
                ultimo = ant["runs"][-1]
                hueco = t["x"] - (ant["x"] + sum(self._medir(r) for r in ant["runs"]))
                if hueco > 0.15 * self._tam(ultimo[1]) and not ultimo[0].endswith(" ") and not t["runs"][0][0].startswith(" "):
                    ultimo[0] += " "
                ant["runs"] += t["runs"]
            else:
                lineas.append(dict(t))

        # líneas con la misma x y separación constante -> un solo cuadro de texto
        bloques = []
        for ln in lineas:
            b = bloques[-1] if bloques else None
            if (b and abs(ln["x"] - b[-1]["x"]) < 0.5 and ln["anchor"] == b[-1]["anchor"]
                    and ln["y"] > b[-1]["y"]):
                dy = ln["y"] - b[-1]["y"]
                if len(b) == 1 or abs(dy - (b[1]["y"] - b[0]["y"])) < 0.05 * dy:
                    if dy < 3 * self._tam(b[-1]["runs"][0][1]):
                        b.append(ln)
                        continue
            bloques.append([ln])

        for b in bloques:
            self._bloque_texto(b, el, ctx)

    def _tam(self, est) -> float:
        return self.longitud(est.get("font-size"), defecto=16) or 16

    def _espaciado(self, est) -> float:
        v = est.get("letter-spacing")
        if not v or v == "normal":
            return 0.0
        return self.longitud(v, fs=self._tam(est))

    def _medir(self, run) -> float:
        texto, est = run
        f = self._fuente(est)
        return ancho_texto(f.info, texto, self._tam(est), self._espaciado(est))

    def _bloque_texto(self, lineas, el, ctx: Contexto):
        primero = lineas[0]["runs"][0][1]
        tam = max(self._tam(r[1]) for r in lineas[0]["runs"])
        info = self._fuente(primero).info
        anchos = [sum(self._medir(r) for r in ln["runs"]) for ln in lineas]
        ancho = max(anchos) * 1.03 + tam * 0.1
        anchor = lineas[0]["anchor"]
        x, y0 = lineas[0]["x"], lineas[0]["y"]

        interlineado = lineas[1]["y"] - lineas[0]["y"] if len(lineas) > 1 else None
        arriba = y0 - linea_base_ppt(info, tam, interlineado)
        alto = (interlineado or 1.2 * tam) * len(lineas)
        izq = {"middle": x - ancho / 2, "end": x - ancho}.get(anchor, x)

        dec = ctx.m.descomponer()
        if dec:
            sx, sy, rot = dec
        else:
            sx = sy = ctx.m.escala
            rot = 0.0
            self.inf.aviso("Algún texto tenía sesgo o espejo; se convirtió sin esa deformación.")
        cx, cy = ctx.m.ap(izq + ancho / 2, arriba + alto / 2)
        W, H = ancho * sx, alto * sy
        texto_plano = " / ".join("".join(r[0] for r in ln["runs"]) for ln in lineas)
        tb = pu.caja_texto(ctx.contenedor, cx - W / 2, cy - H / 2, W, H,
                           self._nombre(el, "Texto") if el.get("id") or el.get("data-name") else f"Texto: {texto_plano[:30]}")
        if abs(rot) > 0.01:
            tb.rotation = rot % 360
        tf = tb.text_frame
        alin = {"middle": PP_ALIGN.CENTER, "end": PP_ALIGN.RIGHT}.get(anchor, PP_ALIGN.LEFT)
        caja_d = (cx - W / 2, cy - H / 2, cx + W / 2, cy + H / 2)
        for i, ln in enumerate(lineas):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = alin
            if interlineado:
                p.line_spacing = Pt(interlineado * sy * PT_POR_PX)
            for texto, est in ln["runs"]:
                f = self._fuente(est)
                pint = self._pintura(est, "fill", ctx.opacidad)
                if pint is None:
                    rgb, alfa = (0, 0, 0), 0.0
                elif pint[0] == "degradado":
                    _, _, paradas = self._degradado_datos(pint[1])
                    if paradas:
                        rgb = tuple(sum(pp[1][k] for pp in paradas) / len(paradas) for k in range(3))
                    else:
                        rgb = (0, 0, 0)
                    alfa = pint[2]
                    self.inf.aviso("Los textos con degradado se convirtieron al color promedio.")
                else:
                    rgb, alfa = pint[1], pint[2]
                trazo = None
                tp = self._pintura(est, "stroke", ctx.opacidad)
                if tp is not None and tp[0] == "solido":
                    w_tr = self.longitud(est.get("stroke-width", "1"), defecto=1) * ctx.m.escala
                    if w_tr > 0:
                        trazo = pu.linea(pu.emu(w_tr), pu.relleno_solido(tp[1], tp[2]),
                                         est.get("stroke-linecap", "butt"), est.get("stroke-linejoin", "miter"))
                deco = est.get("text-decoration", "")
                r = p.add_run()
                r.text = texto
                pu.estilo_run(r, f.nombre, self._tam(est) * sy * PT_POR_PX, f.negrita, f.cursiva,
                              rgb, alfa, self._espaciado(est) * sx * PT_POR_PX,
                              "underline" in deco, "line-through" in deco, trazo=trazo)
        self.inf.contar("textos")
