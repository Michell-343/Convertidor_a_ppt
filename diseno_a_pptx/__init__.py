"""
diseno_a_pptx: convierte diseños (SVG o imágenes) en presentaciones de PowerPoint
editables, conservando tipografías, colores y posición de los elementos.

Uso rápido:
    from diseno_a_pptx import convertir_archivos
    pptx_bytes, informe = convertir_archivos([("portada.svg", open("portada.svg","rb").read())])
"""
from __future__ import annotations

import io
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

from . import imagen_a_pptx, pptx_utils as pu
from .fuentes import CatalogoFuentes
from .imagen_a_pptx import OpcionesImagen
from .svg_a_pptx import ConversorSVG, Informe, escala, traslacion

__all__ = ["convertir_archivos", "CatalogoFuentes", "OpcionesImagen", "Informe"]

ANCHO_ESTANDAR_PX = 1280      # 13,333 pulgadas (diapositiva panorámica 16:9)
MAX_PX, MIN_PX = 5376, 96     # límites de PowerPoint: 56 y 1 pulgadas

_CATALOGO: CatalogoFuentes | None = None


def catalogo() -> CatalogoFuentes:
    global _CATALOGO
    if _CATALOGO is None:
        _CATALOGO = CatalogoFuentes()
    return _CATALOGO


FIRMAS_RASTER = (b"\x89PNG", b"\xff\xd8\xff", b"GIF8", b"BM", b"II*\x00", b"MM\x00*")


def es_svg(nombre: str, datos: bytes) -> bool:
    """Decide por la firma binaria; la extensión y el contenido solo si no hay firma."""
    if datos.startswith(FIRMAS_RASTER) or (datos[:4] == b"RIFF" and datos[8:12] == b"WEBP"):
        return False
    n = nombre.lower()
    if n.endswith((".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".tif", ".tiff")):
        return False
    if n.endswith((".svg", ".svgz")):
        return True
    inicio = datos[:2048].lstrip()
    return inicio.startswith((b"<?xml", b"<svg", b"<!DOCTYPE svg")) and b"<svg" in datos[:4096]


def _tamano_diseno(nombre: str, datos: bytes) -> tuple[float, float]:
    if es_svg(nombre, datos):
        w, h, _ = ConversorSVG.dimensiones(ConversorSVG.cargar(datos))
        return w, h
    return imagen_a_pptx.dimensiones(datos)


def convertir_archivos(archivos: list[tuple[str, bytes]], tamano: str = "estandar",
                       opciones_imagen: OpcionesImagen | None = None,
                       carpeta_base: str | Path | None = None,
                       cat: CatalogoFuentes | None = None) -> tuple[bytes, Informe]:
    """
    archivos: lista de (nombre, contenido). Cada archivo = una diapositiva.
    tamano:   "estandar" -> ancho 13,33 in (16:9 de PowerPoint) con la proporción del diseño
              "real"     -> 1 px del diseño = 1 px a 96 ppp
    """
    if not archivos:
        raise ValueError("No hay archivos para convertir.")
    cat = cat or catalogo()
    op = opciones_imagen or OpcionesImagen()
    inf = Informe()

    w0, h0 = _tamano_diseno(*archivos[0])
    if tamano == "real":
        sw, sh = w0, h0
    else:
        sw, sh = ANCHO_ESTANDAR_PX, ANCHO_ESTANDAR_PX * h0 / w0
    f = min(1.0, MAX_PX / max(sw, sh))
    sw, sh = max(MIN_PX, sw * f), max(MIN_PX, sh * f)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(pu.emu(sw)), Emu(pu.emu(sh))
    vacio = prs.slide_layouts[6]

    for nombre, datos in archivos:
        diapo = prs.slides.add_slide(vacio)
        w, h = _tamano_diseno(nombre, datos)
        s = min(sw / w, sh / h)
        dx, dy = (sw - w * s) / 2, (sh - h * s) / 2
        if abs(w / h - w0 / h0) > 0.01:
            inf.aviso(f"«{nombre}» tiene otra proporción que la primera diapositiva; se centró.")
        if es_svg(nombre, datos):
            ConversorSVG(cat, inf).convertir(datos, diapo, traslacion(dx, dy) * escala(s), carpeta_base)
        else:
            imagen_a_pptx.convertir_imagen(datos, diapo, s, dx, dy, cat, inf, op)
            inf.aviso(f"«{nombre}» es una imagen: los textos se reconstruyeron con OCR y "
                      f"fuente «{op.fuente_texto}». Revisa la ortografía y la fuente. "
                      "Para máxima fidelidad exporta el diseño como SVG.")

    faltan = [n for n, ok in inf.fuentes.items() if not ok]
    if faltan:
        inf.aviso("Fuentes no disponibles para medir los textos: " + ", ".join(faltan)
                  + ". Súbelas (.ttf/.otf) para que los textos queden en la posición exacta, "
                  "y asegúrate de que el equipo del cliente las tenga instaladas.")
    salida = io.BytesIO()
    prs.save(salida)
    return salida.getvalue(), inf
