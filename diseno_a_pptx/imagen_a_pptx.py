"""
Conversión de imágenes rasterizadas (PNG / JPG / WEBP) a diapositivas editables.

Una imagen no contiene información de capas, así que el proceso la "reconstruye":
  1. OCR: detecta cada línea de texto y su posición.
  2. Para cada línea mide el color real del texto, el tamaño (ajustado con las
     métricas de la fuente elegida), la línea base y si es negrita.
  3. Borra los textos del fondo (inpainting) para que no queden duplicados.
  4. (Opcional, experimental) convierte las zonas de color plano en formas
     vectoriales editables; fotos y degradados se conservan como imagen.
  5. Coloca el fondo limpio + formas + cuadros de texto editables.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

from . import pptx_utils as pu
from .fuentes import CatalogoFuentes, ancho_texto, linea_base_ppt
from .svg_a_pptx import Informe
from .vectorizar import separar

PT_POR_PX = 0.75


@dataclass
class OpcionesImagen:
    fuente_texto: str = "Arial"
    fuente_titulos: str | None = None      # None = la misma que fuente_texto
    vectorizar: bool = False                # formas planas -> formas editables
    colores_vectorizar: int = 10
    motor_ocr: str = "auto"                 # "auto" | "rapidocr" | "tesseract"
    idioma_tesseract: str = "spa+eng"
    confianza_minima: float = 0.5


@dataclass
class LineaTexto:
    texto: str
    x0: int
    y0: int
    x1: int
    y1: int
    color: tuple = (0, 0, 0)
    fondo: tuple = (255, 255, 255)
    tamano: float = 12.0       # px de la imagen
    linea_base: float = 0.0
    negrita: bool = False
    fuente: str = "Arial"
    mascara: np.ndarray | None = field(default=None, repr=False)
    angulo: float = 0.0                      # grados (sentido horario, como PowerPoint)
    inversa: np.ndarray | None = field(default=None, repr=False)  # enderezado -> original


def _preparar(rgb: np.ndarray, texto: str, quad: np.ndarray, fuente: str, cat) -> LineaTexto | None:
    """Crea la línea; si el texto está inclinado, la analiza sobre la imagen enderezada."""
    (x0, y0), (x1, y1) = quad[0], quad[1]
    ang = float(np.degrees(np.arctan2(y1 - y0, x1 - x0)))
    if abs(ang) <= 1.5:
        ln = LineaTexto(texto, int(quad[:, 0].min()), int(quad[:, 1].min()),
                        int(quad[:, 0].max()), int(quad[:, 1].max()), fuente=fuente)
        return ln if _analizar(rgb, ln, cat) else None
    h, w = rgb.shape[:2]
    c = quad.mean(axis=0)
    M = cv2.getRotationMatrix2D((float(c[0]), float(c[1])), ang, 1.0)
    recto = cv2.warpAffine(rgb, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    q = np.hstack([quad, np.ones((4, 1))]) @ M.T
    ln = LineaTexto(texto, int(q[:, 0].min()), int(q[:, 1].min()),
                    int(q[:, 0].max()), int(q[:, 1].max()), fuente=fuente)
    if not _analizar(recto, ln, cat):
        return None
    inv = cv2.invertAffineTransform(M)
    ln.mascara = cv2.warpAffine(ln.mascara, inv, (w, h), flags=cv2.INTER_NEAREST)
    ln.angulo, ln.inversa = ang, inv
    return ln


def _unir_fragmentos(lineas: list[LineaTexto], cat) -> list[LineaTexto]:
    """El OCR a veces parte una línea en varias cajas: se unen si comparten línea
    base, tamaño, color y el hueco entre ellas es de una palabra."""
    rectas = sorted([l for l in lineas if l.inversa is None], key=lambda l: (round(l.linea_base), l.x0))
    salida: list[LineaTexto] = []
    for ln in sorted(rectas, key=lambda l: l.x0):
        destino = None
        for a in salida:
            if (abs(a.linea_base - ln.linea_base) < 0.15 * a.tamano
                    and abs(a.tamano - ln.tamano) < 0.15 * a.tamano
                    and a.negrita == ln.negrita
                    and np.linalg.norm(np.array(a.color) - np.array(ln.color)) < 45
                    and -0.2 * a.tamano < ln.x0 - a.x1 < 1.0 * a.tamano):
                destino = a
                break
        if destino is None:
            salida.append(ln)
            continue
        destino.texto = f"{destino.texto} {ln.texto}"
        destino.x1 = max(destino.x1, ln.x1)
        destino.mascara = destino.mascara | ln.mascara
        res = cat.resolver(destino.fuente, 700 if destino.negrita else 400)
        a1 = ancho_texto(res.info, destino.texto, 1.0)
        if a1 > 0:
            destino.tamano = float(np.clip((destino.x1 - destino.x0) / a1,
                                           0.9 * destino.tamano, 1.1 * destino.tamano))
    return salida + [l for l in lineas if l.inversa is not None]


# =================================================================== OCR
_MOTOR_RAPID = None


def _ocr_rapid(rgb: np.ndarray, inf: Informe):
    """Devuelve [(texto, cuadrilátero 4x2, confianza)]."""
    global _MOTOR_RAPID
    bgr = rgb[:, :, ::-1].copy()
    try:
        # RapidOCR 3: modelo LATINO (reconoce tildes, ñ, ¿, ¡, ü...)
        from rapidocr import LangRec, ModelType, OCRVersion, RapidOCR
        if _MOTOR_RAPID is None:
            _MOTOR_RAPID = RapidOCR(params={"Rec.lang_type": LangRec.LATIN,
                                            "Rec.ocr_version": OCRVersion.PPOCRV5,
                                            "Rec.model_type": ModelType.MOBILE,
                                            "Global.log_level": "error"})
        r = _MOTOR_RAPID(bgr)
        if r.boxes is None:
            return []
        return [(t, np.array(b, float), float(s)) for b, t, s in zip(r.boxes, r.txts, r.scores)]
    except ImportError:
        from rapidocr_onnxruntime import RapidOCR
        inf.aviso("Instala «rapidocr>=3» para reconocer tildes y ñ (pip install rapidocr onnxruntime).")
        if _MOTOR_RAPID is None:
            _MOTOR_RAPID = RapidOCR()
        res, _ = _MOTOR_RAPID(bgr)
        return [(t, np.array(b, float), float(s)) for b, t, s in res or []]


def _ocr_tesseract(rgb: np.ndarray, idioma: str):
    import pytesseract
    d = pytesseract.image_to_data(Image.fromarray(rgb), lang=idioma,
                                  output_type=pytesseract.Output.DICT)
    lineas: dict = {}
    for i, palabra in enumerate(d["text"]):
        if not palabra.strip() or float(d["conf"][i]) < 0:
            continue
        clave = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        lineas.setdefault(clave, []).append(i)
    out = []
    for idxs in lineas.values():
        idxs.sort(key=lambda i: d["left"][i])
        # separa palabras muy alejadas (columnas distintas)
        grupo = [idxs[0]]
        for i in idxs[1:]:
            prev = grupo[-1]
            hueco = d["left"][i] - (d["left"][prev] + d["width"][prev])
            if hueco > 2.5 * d["height"][prev]:
                out.append(_unir_tess(d, grupo)); grupo = [i]
            else:
                grupo.append(i)
        out.append(_unir_tess(d, grupo))
    return out


def _unir_tess(d, idxs):
    texto = " ".join(d["text"][i] for i in idxs)
    x0 = min(d["left"][i] for i in idxs)
    y0 = min(d["top"][i] for i in idxs)
    x1 = max(d["left"][i] + d["width"][i] for i in idxs)
    y1 = max(d["top"][i] + d["height"][i] for i in idxs)
    conf = sum(float(d["conf"][i]) for i in idxs) / len(idxs) / 100
    quad = np.array([[x0, y0], [x1, y0], [x1, y1], [x0, y1]], float)
    return texto, quad, conf


def reconocer(rgb: np.ndarray, op: OpcionesImagen, inf: Informe):
    if op.motor_ocr == "tesseract":
        try:
            import pytesseract
            pytesseract.get_tesseract_version()
            return _ocr_tesseract(rgb, op.idioma_tesseract)
        except Exception:
            inf.aviso("Tesseract no está instalado; se usó RapidOCR.")
    return _ocr_rapid(rgb, inf)


# =================================================================== análisis
def _grosor(mask: np.ndarray) -> float:
    dist = cv2.distanceTransform(mask.astype(np.uint8), cv2.DIST_L2, 3)
    return float(2 * np.percentile(dist[dist > 0], 90)) if dist.max() > 0 else 0.0


def _grosor_fuente(info, texto: str, tam: float) -> float:
    """Grosor de trazo del texto dibujado con la fuente indicada (variante Regular)."""
    if info is None or tam < 4:
        return 0.0
    try:
        fuente = ImageFont.truetype(info.ruta, max(4, int(round(tam))), index=info.indice)
        x0, y0, x1, y1 = fuente.getbbox(texto)
        im = Image.new("L", (x1 - x0 + 10, y1 - y0 + 10), 0)
        ImageDraw.Draw(im).text((5 - x0, 5 - y0), texto, font=fuente, fill=255)
        return _grosor(np.array(im) > 127)
    except Exception:
        return 0.0


def _analizar(rgb: np.ndarray, ln: LineaTexto, cat: CatalogoFuentes) -> bool:
    h, w = rgb.shape[:2]
    pad = max(2, int((ln.y1 - ln.y0) * 0.08))
    x0, y0 = max(0, ln.x0 - pad), max(0, ln.y0 - pad)
    x1, y1 = min(w, ln.x1 + pad), min(h, ln.y1 + pad)
    rec = rgb[y0:y1, x0:x1].astype(np.float32)
    if rec.shape[0] < 4 or rec.shape[1] < 4:
        return False
    px = rec.reshape(-1, 3)
    crit = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 20, 1.0)
    _, etiquetas, centros = cv2.kmeans(px, 2, None, crit, 3, cv2.KMEANS_PP_CENTERS)
    borde = np.concatenate([rec[0], rec[-1], rec[:, 0], rec[:, -1]])
    med_borde = np.median(borde, axis=0)
    i_fondo = int(np.argmin(np.linalg.norm(centros - med_borde, axis=1)))
    i_texto = 1 - i_fondo
    c_fondo, c_texto = centros[i_fondo], centros[i_texto]
    if np.linalg.norm(c_fondo - c_texto) < 25:
        return False
    d_texto = np.linalg.norm(px - c_texto, axis=1)
    d_fondo = np.linalg.norm(px - c_fondo, axis=1)
    es_texto = (d_texto < d_fondo).reshape(rec.shape[:2])
    # color real: píxeles "núcleo" del texto (sin antialias)
    nucleo = (d_fondo > np.percentile(d_fondo[es_texto.ravel()], 50)) & es_texto.ravel() if es_texto.any() else es_texto.ravel()
    if nucleo.sum() >= 5:
        ln.color = tuple(int(v) for v in np.median(px[nucleo], axis=0))
    else:
        ln.color = tuple(int(v) for v in c_texto)
    ln.fondo = tuple(int(v) for v in c_fondo)

    mask = es_texto.astype(np.uint8)
    filas = np.where(mask.sum(axis=1) > 0)[0]
    cols = np.where(mask.sum(axis=0) > 0)[0]
    if len(filas) < 2 or len(cols) < 2:
        return False
    tinta_x0, tinta_x1 = x0 + cols[0], x0 + cols[-1] + 1
    perfil = mask.sum(axis=1).astype(float)
    activos = perfil[perfil > 0]
    umbral = 0.35 * np.median(activos)
    base_rel = filas[-1]
    for r in range(filas[-1], filas[0], -1):
        if perfil[r] >= umbral:
            base_rel = r
            break
    ln.linea_base = y0 + base_rel + 1
    alto_tinta = filas[-1] - filas[0] + 1

    # tamaño: se ajusta para que el ancho medido con la fuente coincida
    res = cat.resolver(ln.fuente)
    ancho1 = ancho_texto(res.info, ln.texto, 1.0)
    tam_h = alto_tinta / 0.95
    tam = (tinta_x1 - tinta_x0) / ancho1 if ancho1 > 0 and len(ln.texto) >= 3 else tam_h
    ln.tamano = float(np.clip(tam, 0.6 * tam_h, 1.5 * tam_h))

    # negrita: se compara el grosor del trazo con el de la MISMA fuente en Regular
    grosor = _grosor(mask)
    if grosor:
        referencia = _grosor_fuente(res.info, ln.texto, ln.tamano)
        ln.negrita = grosor > 1.12 * referencia if referencia else grosor / ln.tamano > 0.13
    if ln.negrita and len(ln.texto) >= 3:
        # la variante negrita es más ancha: se recalcula el tamaño con sus métricas
        a1b = ancho_texto(cat.resolver(ln.fuente, 700).info, ln.texto, 1.0)
        if a1b > 0:
            ln.tamano = float(np.clip((tinta_x1 - tinta_x0) / a1b, 0.6 * tam_h, 1.5 * tam_h))

    completa = np.zeros((h, w), np.uint8)
    completa[y0:y1, x0:x1] = mask
    ln.mascara = completa
    ln.x0, ln.x1 = int(tinta_x0), int(tinta_x1)
    return True


def _agrupar(lineas: list[LineaTexto]):
    """Agrupa líneas consecutivas en párrafos (misma fuente, color y alineación)."""
    lineas = sorted(lineas, key=lambda l: (l.linea_base, l.x0))
    bloques: list[tuple[list[LineaTexto], str]] = []
    usadas = set()
    for i, ln in enumerate(lineas):
        if i in usadas:
            continue
        bloque, alin = [ln], "l"
        usadas.add(i)
        while True:
            ult = bloque[-1]
            sig = None
            for j in range(i + 1, len(lineas)):
                if j in usadas:
                    continue
                c = lineas[j]
                dy = c.linea_base - ult.linea_base
                if dy <= 0:
                    continue
                if dy > 2.0 * ult.tamano:
                    break
                if (abs(c.tamano - ult.tamano) > 0.12 * ult.tamano or c.negrita != ult.negrita
                        or np.linalg.norm(np.array(c.color) - np.array(ult.color)) > 45):
                    continue
                if len(bloque) >= 2:
                    dy0 = bloque[1].linea_base - bloque[0].linea_base
                    if abs(dy - dy0) > 0.15 * dy0:
                        continue
                tol = 0.6 * ult.tamano
                if abs(c.x0 - bloque[0].x0) < tol and alin in ("l",):
                    sig = (j, "l")
                elif abs((c.x0 + c.x1) / 2 - (bloque[0].x0 + bloque[0].x1) / 2) < tol and alin in ("l", "c") \
                        and (len(bloque) == 1 or alin == "c"):
                    sig = (j, "c")
                elif abs(c.x1 - bloque[0].x1) < tol and (len(bloque) == 1 or alin == "r"):
                    sig = (j, "r")
                if sig:
                    break
            if not sig:
                break
            usadas.add(sig[0])
            bloque.append(lineas[sig[0]])
            alin = sig[1]
        bloques.append((bloque, alin))
    return bloques


# =================================================================== vectorizado
def _dibujar_separado(diapositiva, fondo: np.ndarray, op: OpcionesImagen, inf: Informe, P, esc: float):
    """Fondo editable + fotos sueltas + figuras/curvas + iconos agrupados."""
    h, w = fondo.shape[:2]
    color_fondo, formas, fotos = separar(fondo, op.colores_vectorizar)
    shapes = diapositiva.shapes

    X0, Y0 = P(0, 0)
    f = pu.nueva_forma(shapes, X0, Y0, w * esc, h * esc, "Fondo")
    pu.aplicar_propiedades(f, pu.geometria_preset("rect"), pu.relleno_solido(color_fondo), pu.sin_linea())

    for n, ft in enumerate(fotos, 1):
        buf = io.BytesIO()
        Image.fromarray(ft.imagen, "RGBA" if ft.imagen.shape[2] == 4 else "RGB").save(buf, "PNG")
        x0, y0, x1, y1 = ft.caja
        X, Y = P(x0, y0)
        pic = shapes.add_picture(io.BytesIO(buf.getvalue()), Emu(pu.emu(X)), Emu(pu.emu(Y)),
                                 Emu(pu.emu((x1 - x0) * esc)), Emu(pu.emu((y1 - y0) * esc)))
        pic.name = f"Foto {n}"
    inf.conteo["fotos"] = inf.conteo.get("fotos", 0) + len(fotos)

    contenedores = {}

    def trazo(color):  # contorno fino del mismo color: evita rendijas entre formas vecinas
        return pu.linea(pu.emu(0.6 * esc), pu.relleno_solido(color), "round", "round")

    def a_diapo(c):
        if c[0] == "Z":
            return c
        vals = []
        for k in range(1, len(c), 2):
            vals += P(c[k], c[k + 1])
        return (c[0], *vals)

    # formas grandes primero (abajo); cada icono se crea al llegar a su primera pieza
    for n, fm in enumerate(formas, 1):
        if fm.grupo and fm.grupo not in contenedores:
            grp = shapes.add_group_shape()
            grp.name = f"Icono {len(contenedores) + 1}"
            contenedores[fm.grupo] = grp
        destino = contenedores[fm.grupo].shapes if fm.grupo else shapes
        if fm.tipo in ("rect", "elipse"):
            (cx, cy), (fw, fh) = fm.centro, fm.tam
            X, Y = P(cx - fw / 2, cy - fh / 2)
            nombre = ("Rectángulo" if fm.tipo == "rect" else "Elipse") + f" {n}"
            sh = pu.nueva_forma(destino, X, Y, fw * esc, fh * esc, nombre)
            if fm.angulo:
                sh.rotation = fm.angulo % 360
            pu.aplicar_propiedades(sh, pu.geometria_preset("rect" if fm.tipo == "rect" else "ellipse"),
                                   pu.relleno_solido(fm.color), trazo(fm.color))
        else:
            subs = [[a_diapo(c) for c in sp] for sp in fm.subtrazos]
            xs = [v for sp in subs for c in sp if c[0] != "Z" for v in c[1::2]]
            ys = [v for sp in subs for c in sp if c[0] != "Z" for v in c[2::2]]
            x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
            sh = pu.nueva_forma(destino, x0, y0, x1 - x0, y1 - y0, f"Forma {n}")
            pu.aplicar_propiedades(sh, pu.geometria_libre(subs, x0, y0, x1 - x0, y1 - y0),
                                   pu.relleno_solido(fm.color), trazo(fm.color))
    for grp in contenedores.values():
        pu.desagrupar_si_trivial(grp)
    inf.conteo["iconos"] = inf.conteo.get("iconos", 0) + len(contenedores)
    inf.conteo["formas"] = inf.conteo.get("formas", 0) + len(formas)


# =================================================================== principal
def dimensiones(datos: bytes) -> tuple[int, int]:
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(datos)))
    return img.size


def convertir_imagen(datos: bytes, diapositiva, esc: float, dx: float, dy: float,
                     cat: CatalogoFuentes, inf: Informe, op: OpcionesImagen):
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(datos)))
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        blanco = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(blanco, img)
    rgb = np.array(img.convert("RGB"))
    h, w = rgb.shape[:2]

    # ------------------------------------------------ 1-2. OCR y análisis
    detectados = []
    for texto, quad, conf in reconocer(rgb, op, inf):
        texto = (texto or "").strip()
        # un único carácter suele ser un icono mal leído (teléfono -> "0", logo -> "©")
        if conf < op.confianza_minima or len(texto) < 2 or not any(ch.isalnum() for ch in texto):
            continue
        detectados.append((texto, quad))
    pares = [(t, q, _preparar(rgb, t, q, op.fuente_texto, cat)) for t, q in detectados]
    pares = [p for p in pares if p[2] is not None]
    if op.fuente_titulos and pares:
        # los textos claramente más grandes que el resto se tratan como títulos
        mediana = float(np.median([p[2].tamano for p in pares]))
        pares = [(t, q, (_preparar(rgb, t, q, op.fuente_titulos, cat) or ln)
                  if ln.tamano > 1.5 * mediana else ln) for t, q, ln in pares]
    lineas = _unir_fragmentos([p[2] for p in pares], cat)

    # ------------------------------------------------ 3. borrar textos
    fondo = rgb
    if lineas:
        # cada texto se borra con un margen proporcional a su tamaño (el antialias
        # de los títulos grandes deja halos más anchos)
        mascara = np.zeros((h, w), np.uint8)
        fondo = rgb.copy()
        for ln in lineas:
            k = max(5, int(ln.tamano * 0.06) | 1)
            m = cv2.dilate(ln.mascara * 255, np.ones((k, k), np.uint8))
            # si el texto está sobre un color plano, se rellena con ESE color exacto
            # (el inpainting dejaría manchas suaves que luego parecerían una foto)
            anillo = (cv2.dilate(m, np.ones((9, 9), np.uint8)) > 0) & (m == 0)
            pix = rgb[anillo].astype(np.float32)
            if len(pix) > 20:
                med = np.median(pix, axis=0)
                cerca = np.abs(pix - med).max(axis=1) < 12
                if cerca.mean() > 0.9:
                    fondo[m > 0] = med.astype(np.uint8)
                    continue
            mascara |= m
        if mascara.any():
            fondo = cv2.inpaint(fondo, mascara, 7, cv2.INPAINT_TELEA)

    def P(x, y):
        return dx + x * esc, dy + y * esc

    # ------------------------------------------------ 4. fondo / formas
    if op.vectorizar:
        _dibujar_separado(diapositiva, fondo, op, inf, P, esc)
    else:
        buf = io.BytesIO()
        Image.fromarray(fondo).save(buf, "PNG")
        X, Y = P(0, 0)
        pic = diapositiva.shapes.add_picture(io.BytesIO(buf.getvalue()), Emu(pu.emu(X)), Emu(pu.emu(Y)),
                                             Emu(pu.emu(w * esc)), Emu(pu.emu(h * esc)))
        pic.name = "Fondo (sin textos)"
        inf.contar("imágenes")

    # ------------------------------------------------ 5. textos
    bloques = _agrupar([l for l in lineas if l.inversa is None])
    bloques += [([l], "l") for l in lineas if l.inversa is not None]
    for bloque, alin in bloques:
        primero = bloque[0]
        res = cat.resolver(primero.fuente, 700 if primero.negrita else 400)
        clave = res.nombre + (" (negrita)" if res.negrita else "")
        inf.fuentes[clave] = res.instalada
        tam = float(np.mean([l.tamano for l in bloque]))
        anchos = [ancho_texto(res.info, l.texto, l.tamano) for l in bloque]
        ancho = max(max(anchos), max(l.x1 - l.x0 for l in bloque)) * 1.04 + tam * 0.1
        inter = ((bloque[-1].linea_base - primero.linea_base) / (len(bloque) - 1)
                 if len(bloque) > 1 else None)
        arriba = primero.linea_base - linea_base_ppt(res.info, tam, inter)
        alto = (inter or 1.2 * tam) * len(bloque)
        if alin == "c":
            cx = float(np.mean([(l.x0 + l.x1) / 2 for l in bloque]))
            izq = cx - ancho / 2
        elif alin == "r":
            izq = max(l.x1 for l in bloque) - ancho
        else:
            izq = min(l.x0 for l in bloque)
        if primero.inversa is not None:
            # caja calculada en la imagen enderezada -> se lleva al original y se rota
            cx, cy = primero.inversa @ np.array([izq + ancho / 2, arriba + alto / 2, 1.0])
            izq, arriba = cx - ancho / 2, cy - alto / 2
        X, Y = P(izq, arriba)
        tb = pu.caja_texto(diapositiva.shapes, X, Y, ancho * esc, alto * esc,
                           f"Texto: {primero.texto[:30]}")
        if primero.angulo:
            tb.rotation = primero.angulo % 360
        tf = tb.text_frame
        for i, ln in enumerate(bloque):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}.get(alin, PP_ALIGN.LEFT)
            if inter:
                p.line_spacing = Pt(inter * esc * PT_POR_PX)
            r = p.add_run()
            r.text = ln.texto
            pu.estilo_run(r, res.nombre, tam * esc * PT_POR_PX, res.negrita or (ln.negrita and not res.instalada),
                          res.cursiva, ln.color)
        inf.contar("textos")
    if not lineas:
        inf.aviso("No se detectó texto en la imagen.")
