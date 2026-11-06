"""
Catálogo de fuentes instaladas en el equipo.

Sirve para tres cosas:
  1. Traducir el nombre de fuente que viene en el SVG (p. ej. "Montserrat-SemiBold",
     o "Montserrat" + font-weight 600) al nombre EXACTO que PowerPoint reconoce
     ("Montserrat SemiBold"), más los indicadores de negrita/cursiva.
  2. Saber si la fuente está instalada (para avisar a la diseñadora).
  3. Medir textos (ancho y ascendente) con las métricas reales de la fuente, para
     colocar las cajas de texto en el mismo sitio que en el diseño.
"""
from __future__ import annotations

import json
import os
import re
import sys
from dataclasses import asdict, dataclass
from functools import lru_cache
from pathlib import Path

import logging

from fontTools.ttLib import TTCollection, TTFont

logging.getLogger("fontTools").setLevel(logging.ERROR)

EXTENSIONES = (".ttf", ".otf", ".ttc", ".otc")

GENERICAS = {
    "sans-serif": "Arial",
    "serif": "Times New Roman",
    "monospace": "Consolas",
    "cursive": "Comic Sans MS",
    "fantasy": "Impact",
    "system-ui": "Segoe UI",
}

EQUIVALENTES_METRICOS = {
    "arial": "Liberation Sans", "helvetica": "Liberation Sans", "helveticaneue": "Liberation Sans",
    "timesnewroman": "Liberation Serif", "times": "Liberation Serif",
    "couriernew": "Liberation Mono", "courier": "Liberation Mono",
}

PESOS_NOMBRE = {
    "thin": 100, "hairline": 100, "extralight": 200, "ultralight": 200,
    "light": 300, "book": 400, "regular": 400, "normal": 400, "roman": 400,
    "medium": 500, "semibold": 600, "demibold": 600, "demi": 600,
    "bold": 700, "extrabold": 800, "ultrabold": 800, "heavy": 900, "black": 900,
}


def _directorios_fuentes() -> list[str]:
    dirs: list[str] = []
    if sys.platform.startswith("win"):
        windir = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(os.path.join(windir, "Fonts"))
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
    elif sys.platform == "darwin":
        dirs += ["/System/Library/Fonts", "/Library/Fonts", os.path.expanduser("~/Library/Fonts")]
    else:
        dirs += ["/usr/share/fonts", "/usr/local/share/fonts",
                 os.path.expanduser("~/.fonts"), os.path.expanduser("~/.local/share/fonts")]
    return [d for d in dirs if os.path.isdir(d)]


def normalizar(nombre: str) -> str:
    return re.sub(r"[\s\-_'\".,]+", "", nombre or "").lower()


@dataclass
class InfoFuente:
    ruta: str
    indice: int
    familia: str          # name ID 1 -> el nombre que usa PowerPoint en Windows
    subfamilia: str       # name ID 2 -> Regular / Bold / Italic / Bold Italic
    familia_tipo: str     # name ID 16 (familia tipográfica) o ID 1
    subfamilia_tipo: str  # name ID 17 o ID 2
    completo: str         # name ID 4
    postscript: str       # name ID 6
    peso: int
    cursiva: bool
    upm: int
    ascendente: int
    descendente: int

    @property
    def negrita_flag(self) -> bool:
        return "bold" in self.subfamilia.lower()

    @property
    def cursiva_flag(self) -> bool:
        s = self.subfamilia.lower()
        return "italic" in s or "oblique" in s


@dataclass
class FuenteResuelta:
    nombre: str            # nombre para PowerPoint
    negrita: bool
    cursiva: bool
    info: InfoFuente | None
    instalada: bool
    original: str          # lo que pedía el diseño


def _leer_fuente(f: TTFont, ruta: str, indice: int) -> InfoFuente | None:
    try:
        nombre = f["name"]

        def n(i):
            v = nombre.getDebugName(i)
            return v.strip() if v else ""

        familia = n(1)
        if not familia:
            return None
        sub = n(2) or "Regular"
        os2 = f["OS/2"] if "OS/2" in f else None
        head = f["head"]
        peso = int(getattr(os2, "usWeightClass", 400) or 400) if os2 else 400
        cursiva = bool(os2.fsSelection & 1) if os2 else bool(head.macStyle & 2)
        if os2 and getattr(os2, "usWinAscent", 0):
            asc, desc = int(os2.usWinAscent), int(os2.usWinDescent)
        else:
            hhea = f["hhea"]
            asc, desc = int(hhea.ascent), abs(int(hhea.descent))
        return InfoFuente(
            ruta=ruta, indice=indice, familia=familia, subfamilia=sub,
            familia_tipo=n(16) or familia, subfamilia_tipo=n(17) or sub,
            completo=n(4) or familia, postscript=n(6),
            peso=peso, cursiva=cursiva, upm=int(head.unitsPerEm),
            ascendente=asc, descendente=desc,
        )
    except Exception:
        return None


def _escanear_archivo(ruta: str) -> list[InfoFuente]:
    salida = []
    try:
        if ruta.lower().endswith((".ttc", ".otc")):
            col = TTCollection(ruta, lazy=True)
            for i, f in enumerate(col.fonts):
                info = _leer_fuente(f, ruta, i)
                if info:
                    salida.append(info)
        else:
            f = TTFont(ruta, lazy=True)
            info = _leer_fuente(f, ruta, 0)
            if info:
                salida.append(info)
    except Exception:
        pass
    return salida


class CatalogoFuentes:
    """Escanea una sola vez las fuentes del sistema (con caché en disco)."""

    def __init__(self, directorios_extra: list[str] | None = None, incluir_sistema: bool = True):
        self.fuentes: list[InfoFuente] = []
        self._por_ps: dict[str, InfoFuente] = {}
        self._por_completo: dict[str, InfoFuente] = {}
        self._por_familia: dict[str, list[InfoFuente]] = {}
        # las fuentes subidas por la usuaria van primero: tienen prioridad
        self._cargar((directorios_extra or []) + (_directorios_fuentes() if incluir_sistema else []))

    # ------------------------------------------------------------------ carga
    def _cargar(self, directorios: list[str]):
        cache_path = Path.home() / ".cache" / "diseno_a_pptx" / "fuentes.json"
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}
        nueva_cache = {}
        for d in directorios:
            for raiz, _, archivos in os.walk(d):
                for a in archivos:
                    if not a.lower().endswith(EXTENSIONES):
                        continue
                    ruta = os.path.join(raiz, a)
                    try:
                        mtime = os.path.getmtime(ruta)
                    except OSError:
                        continue
                    previo = cache.get(ruta)
                    if previo and previo[0] == mtime:
                        infos = [InfoFuente(**x) for x in previo[1]]
                    else:
                        infos = _escanear_archivo(ruta)
                    nueva_cache[ruta] = [mtime, [asdict(x) for x in infos]]
                    self.fuentes.extend(infos)
        try:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            cache_path.write_text(json.dumps(nueva_cache), encoding="utf-8")
        except Exception:
            pass

        for f in self.fuentes:
            if f.postscript:
                self._por_ps.setdefault(normalizar(f.postscript), f)
            self._por_completo.setdefault(normalizar(f.completo), f)
            for fam in {normalizar(f.familia), normalizar(f.familia_tipo)}:
                self._por_familia.setdefault(fam, []).append(f)

    # -------------------------------------------------------------- consultas
    def familias(self) -> list[str]:
        return sorted({f.familia for f in self.fuentes}, key=str.lower)

    def _elegir_en_familia(self, lista: list[InfoFuente], peso: int, cursiva: bool) -> InfoFuente:
        return min(lista, key=lambda f: ((f.cursiva_flag or f.cursiva) != cursiva,
                                         abs(f.peso - peso),
                                         len(f.familia)))

    def resolver(self, familia_css: str, peso: int = 400, cursiva: bool = False) -> FuenteResuelta:
        candidatos = [c.strip().strip("'\"") for c in (familia_css or "").split(",") if c.strip()]
        if not candidatos:
            candidatos = ["Arial"]
        for c in candidatos:
            nc = normalizar(c)
            if c.lower() in GENERICAS:
                c = GENERICAS[c.lower()]
                nc = normalizar(c)
            if nc in self._por_familia:
                # Nombre de familia + peso: se elige la variante más cercana.
                info = self._elegir_en_familia(self._por_familia[nc], peso, cursiva)
            else:
                # Nombre PostScript ("Montserrat-SemiBold") o completo ("Montserrat SemiBold")
                info = self._por_ps.get(nc) or self._por_completo.get(nc)
            if info is not None:
                negrita = info.negrita_flag
                curs = info.cursiva_flag or (cursiva and not info.cursiva)
                return FuenteResuelta(info.familia, negrita, curs, info, True, candidatos[0])

        # Fuentes estándar de Windows/Mac ausentes en servidores Linux: se miden con su
        # equivalente métrico libre (mismo ancho de cada carácter) y se conserva el nombre.
        for c in candidatos:
            eq = EQUIVALENTES_METRICOS.get(normalizar(c))
            if eq and normalizar(eq) in self._por_familia:
                info = self._elegir_en_familia(self._por_familia[normalizar(eq)], peso, cursiva)
                return FuenteResuelta(c, info.negrita_flag, info.cursiva_flag or cursiva, info, True, c)

        # No está instalada: deducimos el nombre más probable.
        primero = candidatos[0]
        nombre, negrita, curs = primero, peso >= 700, cursiva
        if "-" in primero:  # nombre PostScript tipo "Montserrat-SemiBoldItalic"
            fam, estilo = primero.split("-", 1)
            est = estilo.lower()
            curs = curs or "italic" in est or "oblique" in est
            est_sin = est.replace("italic", "").replace("oblique", "")
            if est_sin in ("", "regular", "roman", "bold"):
                nombre, negrita = fam, negrita or est_sin == "bold"
            else:
                nombre = f"{fam} {estilo.replace('Italic', '').replace('Oblique', '').strip()}"
                negrita = False
        elif peso not in (400, 700) and " " not in primero:
            # Windows registra los pesos intermedios como familias propias
            sufijo = {100: "Thin", 200: "ExtraLight", 300: "Light", 500: "Medium",
                      600: "SemiBold", 800: "ExtraBold", 900: "Black"}.get(peso)
            if sufijo:
                nombre, negrita = f"{primero} {sufijo}", False
        return FuenteResuelta(nombre, negrita, curs, None, False, primero)


# ------------------------------------------------------------------ métricas
@lru_cache(maxsize=64)
def _tablas_ancho(ruta: str, indice: int):
    try:
        f = TTFont(ruta, fontNumber=indice, lazy=True)
        cmap = f.getBestCmap() or {}
        hmtx = f["hmtx"].metrics
        return cmap, {g: m[0] for g, m in hmtx.items()}
    except Exception:
        return None


def ancho_texto(info: InfoFuente | None, texto: str, tamano: float, espaciado: float = 0.0) -> float:
    """Ancho del texto (en las mismas unidades que `tamano`)."""
    if not texto:
        return 0.0
    tablas = _tablas_ancho(info.ruta, info.indice) if info else None
    if not tablas:
        return len(texto) * 0.55 * tamano + espaciado * len(texto)
    cmap, anchos = tablas
    total = 0
    por_defecto = anchos.get(cmap.get(ord("n")), info.upm // 2)
    for ch in texto:
        g = cmap.get(ord(ch))
        total += anchos.get(g, por_defecto) if g else por_defecto
    return total * tamano / info.upm + espaciado * len(texto)


def metricas_verticales(info: InfoFuente | None) -> tuple[float, float]:
    """(ascendente, descendente) como fracción del tamaño de fuente."""
    if not info or not info.upm:
        return 0.905, 0.212  # valores de Arial
    return info.ascendente / info.upm, info.descendente / info.upm


def linea_base_ppt(info: InfoFuente | None, tamano: float, interlineado: float | None = None) -> float:
    """Distancia desde el borde superior del cuadro (sin márgenes) hasta la PRIMERA
    línea base, tal como la dibuja PowerPoint. Medido empíricamente renderizando con
    PowerPoint 365 varias fuentes e interlineados (error < 1 % del tamaño):
      - interlineado sencillo: alto de línea H = 1,2 em, base en H·r
      - interlineado exacto L > H: base en L·min(r, 0,75)
      - interlineado exacto L <= H: base en L - H·(1-r)   (o L·r si r < 0,75)
    con r = winAscent / (winAscent + winDescent)."""
    asc, desc = metricas_verticales(info)
    r = asc / (asc + desc)
    H = 1.2 * tamano
    if interlineado is None:
        return H * r
    L = interlineado
    if L > H:
        return L * min(r, 0.75)
    return L - H * (1 - r) if r > 0.75 else L * r


def parsear_peso(valor) -> int:
    if valor is None:
        return 400
    v = str(valor).strip().lower()
    if v.isdigit():
        return int(v)
    if v in ("bold", "bolder"):
        return 700
    if v == "lighter":
        return 300
    return PESOS_NOMBRE.get(v.replace(" ", "").replace("-", ""), 400)
