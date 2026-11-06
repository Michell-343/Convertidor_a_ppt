"""
Separación de una imagen en elementos editables:

  - Fondo            -> rectángulo de color editable
  - Figuras simples  -> rectángulos y elipses NATIVOS de PowerPoint (con rotación)
  - Otras formas     -> curvas Bézier suaves (esquinas conservadas)
  - Iconos           -> formas pequeñas cercanas agrupadas ("Icono 1", "Icono 2"…)
  - Fotos/degradados -> recortadas como imágenes sueltas, cada una en su sitio
                        (rectangulares tal cual; con otra silueta, con transparencia)
"""
from __future__ import annotations

from dataclasses import dataclass, field

import cv2
import numpy as np


@dataclass
class Forma:
    tipo: str                 # "rect" | "elipse" | "libre"
    color: tuple
    area: float               # px² en la imagen original
    caja: tuple               # x0, y0, x1, y1 (px imagen original)
    nitidez: float = 0.0      # fuerza media del borde
    centro: tuple = (0.0, 0.0)
    tam: tuple = (0.0, 0.0)   # ancho, alto
    angulo: float = 0.0       # grados, sentido horario (como PowerPoint)
    subtrazos: list = field(default_factory=list)  # comandos M/L/C/Z (px imagen)
    grupo: int | None = None


@dataclass
class Foto:
    caja: tuple               # x0, y0, x1, y1 (px imagen original)
    imagen: np.ndarray        # RGB o RGBA recortada
    rectangular: bool


# ------------------------------------------------------------------ geometría
def _poligono_rotado(cx, cy, w, h, ang):
    """Rectángulo de w×h centrado y girado `ang`° en sentido horario (convención PowerPoint)."""
    r = np.radians(ang)
    c, s = np.cos(r), np.sin(r)
    esquinas = np.array([[-w / 2, -h / 2], [w / 2, -h / 2], [w / 2, h / 2], [-w / 2, h / 2]])
    return esquinas @ np.array([[c, s], [-s, c]]) + [cx, cy]


def _poligono_elipse(cx, cy, w, h, ang, n=72):
    t = np.linspace(0, 2 * np.pi, n, endpoint=False)
    p = np.stack([np.cos(t) * w / 2, np.sin(t) * h / 2], axis=1)
    r = np.radians(ang)
    c, s = np.cos(r), np.sin(r)
    return p @ np.array([[c, s], [-s, c]]) + [cx, cy]


def _iou(region, poligono, x0, y0):
    """Coincidencia de la figura con la región, en un lienzo que contiene a ambas.
    Devuelve 0 si la figura invade más de un 2,5 % fuera de la región (taparía vecinos)."""
    pol = poligono - [x0, y0]
    ox = int(np.floor(min(0, pol[:, 0].min()))) - 1
    oy = int(np.floor(min(0, pol[:, 1].min()))) - 1
    W = int(np.ceil(max(region.shape[1], pol[:, 0].max()))) - ox + 2
    H = int(np.ceil(max(region.shape[0], pol[:, 1].max()))) - oy + 2
    reg = np.zeros((H, W), np.uint8)
    reg[-oy:-oy + region.shape[0], -ox:-ox + region.shape[1]] = region
    m = np.zeros((H, W), np.uint8)
    cv2.fillPoly(m, [np.round(pol - [ox, oy]).astype(np.int32)], 1)
    inter = np.logical_and(m, reg).sum()
    union = np.logical_or(m, reg).sum()
    fuera = (m.astype(bool) & ~reg.astype(bool)).sum() / max(1, reg.sum())
    return inter / union if union and fuera < 0.025 else 0.0


def _normalizar(w, h, ang):
    ang = ((ang + 90) % 180) - 90          # (-90, 90]
    if ang > 45:
        w, h, ang = h, w, ang - 90
    elif ang < -45:
        w, h, ang = h, w, ang + 90
    return w, h, (0.0 if abs(ang) < 0.5 else ang)


def _clasificar(cont, region_recorte, x0, y0):
    """Prueba si la región es un rectángulo o una elipse. -> (tipo, cx, cy, w, h, ang) o None"""
    mejor = None
    (cx, cy), (w, h), ang = cv2.minAreaRect(cont)
    if w > 2 and h > 2:
        for a in (ang, -ang):
            ww, hh, aa = _normalizar(w, h, a)
            v = _iou(region_recorte, _poligono_rotado(cx, cy, ww, hh, aa), x0, y0)
            if v > 0.95 and (mejor is None or v > mejor[0]):
                mejor = (v, "rect", cx, cy, ww, hh, aa)
    if len(cont) >= 5:
        (cx, cy), (w, h), ang = cv2.fitEllipse(cont)
        if w > 2 and h > 2:
            for a in (ang, -ang):
                ww, hh, aa = _normalizar(w, h, a)
                v = _iou(region_recorte, _poligono_elipse(cx, cy, ww, hh, aa), x0, y0)
                if v > 0.95 and (mejor is None or v > mejor[0] + 0.005):
                    mejor = (v, "elipse", cx, cy, ww, hh, aa)
    return mejor[1:] if mejor else None


def _curva(cont, escala):
    """Contorno de píxeles -> trazado Bézier suave con esquinas conservadas."""
    p = cont.reshape(-1, 2).astype(np.float64)
    if len(p) >= 8:  # suavizado circular para quitar el escalonado de píxeles
        k = np.array([1, 4, 6, 4, 1], float) / 16
        ext = np.vstack([p[-2:], p, p[:2]])
        p = np.stack([np.convolve(ext[:, i], k, mode="valid") for i in range(2)], axis=1)
    p = cv2.approxPolyDP(p.astype(np.float32).reshape(-1, 1, 2), 0.8, True).reshape(-1, 2).astype(np.float64)
    n = len(p)
    if n < 3:
        return None
    largo = 60.0  # segmentos largos = rectos (un arco suave nunca los genera con eps 0,8)
    esquina = []
    for i in range(n):
        a, b = p[i - 1] - p[i], p[(i + 1) % n] - p[i]
        la, lb = np.linalg.norm(a), np.linalg.norm(b)
        cos = np.dot(a, b) / (la * lb) if la and lb else -1
        angulo = np.degrees(np.arccos(np.clip(cos, -1, 1)))
        esquina.append(angulo < 145 or la > largo or lb > largo)
    p = p / escala
    cmds = [("M", *p[0])]
    for i in range(n):
        p0, p1, p2, p3 = p[i - 1], p[i], p[(i + 1) % n], p[(i + 2) % n]
        e1, e2 = esquina[i], esquina[(i + 1) % n]
        if e1 and e2:
            cmds.append(("L", *p2))
            continue
        c1 = p1 if e1 else p1 + (p2 - p0) / 6
        c2 = p2 if e2 else p2 - (p3 - p1) / 6
        cmds.append(("C", *c1, *c2, *p2))
    cmds.append(("Z",))
    return cmds




# ------------------------------------------------------------------ fotografías
BLOQUE = 16  # px (a la escala de trabajo)


def _semillas_fotos(rgb_peq: np.ndarray) -> np.ndarray:
    """Zonas claramente fotográficas: bloques que necesitan muchos colores distintos para
    cubrir el 90 % de sus píxeles (un bloque gráfico se cubre con 2-3 colores)."""
    ph, pw = rgb_peq.shape[:2]
    B = BLOQUE
    nby, nbx = ph // B, pw // B
    if nby == 0 or nbx == 0:
        return np.zeros((ph, pw), np.uint8)
    px = rgb_peq[: nby * B, : nbx * B].astype(np.float32)
    bloques = px.reshape(nby, B, nbx, B, 3).transpose(0, 2, 1, 3, 4).reshape(-1, B * B, 3)
    n90 = np.empty(len(bloques), np.int32)
    meta = 0.1 * B * B
    for i, b in enumerate(bloques):
        # cuántos colores (con tolerancia al ruido JPEG) hacen falta para cubrir el 90 %
        resto = b
        n = 0
        while len(resto) > meta and n < 8:
            q = (resto // 16).astype(np.int32)
            cod = (q[:, 0] << 8) | (q[:, 1] << 4) | q[:, 2]
            moda = np.bincount(cod).argmax()
            centro = resto[cod == moda].mean(axis=0)
            resto = resto[np.abs(resto - centro).max(axis=1) > 14]
            n += 1
        n90[i] = n
    # Pendiente de color: un color plano de diseño tiene pendiente CERO (también en JPEG);
    # hasta la zona más suave de una foto (cielo, fondo de estudio) varía un poco.
    # La mediana por bloque ignora los bordes de las formas.
    # Se cuenta la fracción de píxeles con pendiente SUAVE (ni plana ni borde abrupto):
    # en una foto casi todos lo son; en un gráfico solo unos pocos junto a los bordes.
    L = cv2.GaussianBlur(cv2.cvtColor(rgb_peq, cv2.COLOR_RGB2LAB).astype(np.float32), (0, 0), 1.5)
    g = np.zeros(L.shape[:2], np.float32)
    for c in range(3):
        g += cv2.Sobel(L[..., c], cv2.CV_32F, 1, 0, ksize=3) ** 2 + cv2.Sobel(L[..., c], cv2.CV_32F, 0, 1, ksize=3) ** 2
    g = np.sqrt(g)[: nby * B, : nbx * B] / 8
    suave = ((g > 0.06) & (g < 2.5)).astype(np.float32)
    frac_suave = suave.reshape(nby, B, nbx, B).mean(axis=(1, 3))
    foto_b = ((n90.reshape(nby, nbx) >= 5) | (frac_suave > 0.7)).astype(np.uint8)
    # bloques aislados = ruido (detalles finos de un gráfico)
    vecinos = cv2.filter2D(foto_b, -1, np.ones((3, 3), np.float32), borderType=cv2.BORDER_CONSTANT)
    foto_b = ((foto_b == 1) & (vecinos >= 3)).astype(np.uint8)
    m = np.zeros((ph, pw), np.uint8)
    m[: nby * B, : nbx * B] = cv2.resize(foto_b, (nbx * B, nby * B), interpolation=cv2.INTER_NEAREST)
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, np.ones((B + 1, B + 1), np.uint8))  # no une fotos vecinas
    base = m.copy()
    n, etq, stats, _ = cv2.connectedComponentsWithStats(m, connectivity=8)
    m = np.zeros_like(base)
    for i in range(1, n):
        if stats[i, cv2.CC_STAT_AREA] < 0.01 * ph * pw:
            continue
        comp = (etq == i).astype(np.uint8)
        lleno = _rellenar(comp)
        # el hueco se rellena solo si casi todo ya era foto (un anillo de bordes no cuenta)
        m |= lleno if comp.sum() > 0.6 * lleno.sum() else comp
    # se encoge un bloque: la semilla queda DENTRO de la foto; el borde exacto se busca después
    return cv2.erode(m, np.ones((B + 1, B + 1), np.uint8))


def _rellenar(m: np.ndarray) -> np.ndarray:
    conts, _ = cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    out = np.zeros_like(m)
    cv2.drawContours(out, conts, -1, 1, -1)
    return out


def _iou_simple(region, poligono, x0, y0):
    m = np.zeros_like(region)
    cv2.fillPoly(m, [np.round(poligono - [x0, y0]).astype(np.int32)], 1)
    union = np.logical_or(m, region).sum()
    return np.logical_and(m, region).sum() / union if union else 0.0


def _marco(mask: np.ndarray):
    """¿Qué tan bien encaja la zona en un marco de foto? -> (puntaje, tipo, datos)"""
    ys, xs = np.nonzero(mask)
    if len(xs) < 10:
        return 0.0, "libre", None
    x0, y0, x1, y1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1
    rec = mask[y0:y1, x0:x1]
    rect = rec.sum() / rec.size
    mejor = (rect, "rect", (x0, y0, x1, y1))
    conts, _ = cv2.findContours(rec.astype(np.uint8), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    if conts:
        c = max(conts, key=cv2.contourArea)
        if len(c) >= 5:
            (cx, cy), (w, h), ang = cv2.fitEllipse(c)
            for a in (ang, -ang):
                ww, hh, aa = _normalizar(w, h, a)
                v = _iou_simple(rec, _poligono_elipse(cx, cy, ww, hh, aa), 0, 0)
                if v > mejor[0] + 0.02:
                    mejor = (v, "elipse", (cx + x0, cy + y0, ww, hh, aa))
    return mejor


def _silueta(crudo: np.ndarray, pm: np.ndarray, margen: int) -> np.ndarray:
    """Silueta real de la foto con GrabCut: la zona detectada es primer plano seguro y
    el entorno lejano es fondo seguro; el algoritmo encuentra el borde exacto."""
    ph, pw = pm.shape
    ys, xs = np.nonzero(pm)
    x0, y0 = max(0, xs.min() - margen), max(0, ys.min() - margen)
    x1, y1 = min(pw, xs.max() + 1 + margen), min(ph, ys.max() + 1 + margen)
    r = 8  # anillo exterior de fondo seguro
    cx0, cy0, cx1, cy1 = max(0, x0 - r), max(0, y0 - r), min(pw, x1 + r), min(ph, y1 + r)
    recorte = np.ascontiguousarray(crudo[cy0:cy1, cx0:cx1, ::-1])
    sub = pm[cy0:cy1, cx0:cx1]
    mask = np.full(sub.shape, cv2.GC_BGD, np.uint8)
    mask[y0 - cy0:y1 - cy0, x0 - cx0:x1 - cx0] = cv2.GC_PR_BGD
    mask[sub > 0] = cv2.GC_PR_FGD
    mask[cv2.erode(sub, np.ones((5, 5), np.uint8)) > 0] = cv2.GC_FGD
    try:
        bgd, fgd = np.zeros((1, 65), np.float64), np.zeros((1, 65), np.float64)
        cv2.grabCut(recorte, mask, None, bgd, fgd, 4, cv2.GC_INIT_WITH_MASK)
    except cv2.error:
        return pm
    fg = ((mask == cv2.GC_FGD) | (mask == cv2.GC_PR_FGD)).astype(np.uint8)
    fg = cv2.morphologyEx(fg, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))
    n, etq = cv2.connectedComponents(fg)
    if n <= 1:
        return pm
    # se queda la parte conectada con la zona detectada
    ids = np.unique(etq[(sub > 0) & (fg > 0)])
    fg = np.isin(etq, ids[ids > 0]).astype(np.uint8)
    out = np.zeros_like(pm)
    out[cy0:cy1, cx0:cx1] = _rellenar(fg)
    return out | pm


def _ajustar_borde(gris: np.ndarray, caja, radio: int):
    """Mueve cada lado del rectángulo al borde recto más fuerte cercano."""
    h, w = gris.shape
    x0, y0, x1, y1 = [int(round(v)) for v in caja]
    dx = np.abs(np.diff(gris, axis=1))  # dx[:, x] = borde entre x y x+1
    dy = np.abs(np.diff(gris, axis=0))

    def mejor_col(xc, ya, yb):
        rango = range(max(1, xc - radio), min(w - 1, xc + radio + 1))
        if ya >= yb or not rango:
            return xc
        puntajes = [dx[ya:yb, x - 1].mean() for x in rango]
        return rango[int(np.argmax(puntajes))] if max(puntajes) > 6 else xc

    def mejor_fila(yc, xa, xb):
        rango = range(max(1, yc - radio), min(h - 1, yc + radio + 1))
        if xa >= xb or not rango:
            return yc
        puntajes = [dy[y - 1, xa:xb].mean() for y in rango]
        return rango[int(np.argmax(puntajes))] if max(puntajes) > 6 else yc

    ya, yb = y0 + (y1 - y0) // 5, y1 - (y1 - y0) // 5     # tramo central de cada lado
    xa, xb = x0 + (x1 - x0) // 5, x1 - (x1 - x0) // 5
    nx0 = 0 if x0 <= radio else mejor_col(x0, ya, yb)
    nx1 = w if x1 >= w - radio else mejor_col(x1, ya, yb)
    ny0 = 0 if y0 <= radio else mejor_fila(y0, xa, xb)
    ny1 = h if y1 >= h - radio else mejor_fila(y1, xa, xb)
    return nx0, ny0, nx1, ny1


# ------------------------------------------------------------------ principal
def separar(rgb: np.ndarray, k: int = 10):
    """-> (color_fondo, formas, fotos)"""
    h, w = rgb.shape[:2]
    s = min(1.0, 1400 / max(h, w))
    crudo = cv2.resize(rgb, (int(w * s), int(h * s)), interpolation=cv2.INTER_AREA) if s < 1 else rgb.copy()
    peq = cv2.bilateralFilter(crudo, 5, 30, 5)
    ph, pw = peq.shape[:2]

    # 1. las fotos se detectan ANTES de vectorizar y quedan fuera del proceso
    semillas = _semillas_fotos(crudo)
    lab = cv2.cvtColor(peq, cv2.COLOR_RGB2LAB).astype(np.float32)
    datos = lab[semillas == 0]
    if len(datos) < k * 10:
        datos = lab.reshape(-1, 3)
    if len(datos) > 250_000:
        datos = datos[np.random.default_rng(0).choice(len(datos), 250_000, replace=False)]
    crit = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 30, 0.5)
    _, _, centros = cv2.kmeans(datos, k, None, crit, 2, cv2.KMEANS_PP_CENTERS)
    planos = lab.reshape(-1, 3)
    etq = np.empty(len(planos), np.uint8)
    for i in range(0, len(planos), 200_000):
        d = ((planos[i:i + 200_000, None, :] - centros[None]) ** 2).sum(axis=2)
        etq[i:i + 200_000] = d.argmin(axis=1)
    etq = cv2.medianBlur(etq.reshape(ph, pw), 5)
    etq[semillas > 0] = 255

    gx = cv2.Sobel(lab, cv2.CV_32F, 1, 0, ksize=3)
    gy = cv2.Sobel(lab, cv2.CV_32F, 0, 1, ksize=3)
    borde = cv2.dilate(np.sqrt((gx ** 2 + gy ** 2).sum(axis=2)), np.ones((5, 5), np.uint8))

    # 2. regiones de color plano (candidatas a forma)
    area_min = 0.00004 * ph * pw
    candidatas = []  # (Forma, bx, by, region)
    nucleo = np.ones((3, 3), np.uint8)
    for c in range(k):
        m = cv2.morphologyEx((etq == c).astype(np.uint8), cv2.MORPH_OPEN, nucleo)
        if m.sum() < area_min:
            continue
        conts, jer = cv2.findContours(m, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_NONE)
        if jer is None:
            continue
        jer = jer[0]
        for i, cont in enumerate(conts):
            if jer[i][3] != -1 or cv2.contourArea(cont) < area_min:
                continue
            bx, by, bw, bh = cv2.boundingRect(cont)
            region = np.zeros((bh, bw), np.uint8)
            cv2.drawContours(region, [cont - [bx, by]], -1, 1, -1)
            agujeros = []
            hijo = jer[i][2]
            while hijo != -1:
                if cv2.contourArea(conts[hijo]) > area_min:
                    agujeros.append(conts[hijo])
                    cv2.drawContours(region, [conts[hijo] - [bx, by]], -1, 0, -1)
                hijo = jer[hijo][0]
            pix = peq[by:by + bh, bx:bx + bw][region > 0]
            if len(pix) < 10:
                continue
            pf = pix.astype(np.float32)
            dist = np.linalg.norm(pf - np.median(pf, axis=0), axis=1)
            if np.mean(np.std(pf[dist <= np.percentile(dist, 70)], axis=0)) > 10:   # textura
                continue
            pc = np.vstack([ct.reshape(-1, 2) for ct in [cont] + agujeros])
            interior = (pc[:, 0] > 2) & (pc[:, 0] < pw - 3) & (pc[:, 1] > 2) & (pc[:, 1] < ph - 3)
            nitidez = 999.0
            if interior.sum() > 20:
                fuerza = borde[pc[interior, 1], pc[interior, 0]]
                if np.mean(fuerza < 60) > 0.3:                                       # degradado
                    continue
                nitidez = float(np.mean(fuerza))
            f = Forma("libre", tuple(int(v) for v in np.median(pix, axis=0)), float(region.sum()) / (s * s),
                      (bx / s, by / s, (bx + bw) / s, (by + bh) / s), nitidez)
            prim = _clasificar(cont, region, bx, by) if not agujeros else None
            if prim:
                f.tipo, f.centro, f.tam, f.angulo = prim[0], (prim[1] / s, prim[2] / s), (prim[3] / s, prim[4] / s), prim[5]
            else:
                f.subtrazos = [sp for ct in [cont] + agujeros if (sp := _curva(ct, s))]
                if not f.subtrazos:
                    continue
            candidatas.append((f, bx, by, region))

    # 3. cada foto absorbe las zonas planas que forman parte de ella (fondo negro de un
    #    producto, cielo quemado, pared…): las que la rodean y completan su marco
    n_f, etq_f = cv2.connectedComponents(semillas, connectivity=8)
    fotos_m = [(etq_f == i).astype(np.uint8) for i in range(1, n_f)]
    absorbida = [False] * len(candidatas)
    for fi, pm in enumerate(fotos_m):
        for _ in range(6):
            cambio = False
            dil = cv2.dilate(pm, np.ones((7, 7), np.uint8))
            puntaje, _, _ = _marco(pm)
            ys, xs = np.nonzero(pm)
            px0, py0, px1, py1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1
            area_p = pm.sum()
            for j, (f, bx, by, reg) in enumerate(candidatas):
                if absorbida[j]:
                    continue
                bh, bw = reg.shape
                if bx > px1 + 4 or by > py1 + 4 or bx + bw < px0 - 4 or by + bh < py0 - 4:
                    continue
                contorno = reg - cv2.erode(reg, nucleo)
                toque = (contorno & dil[by:by + bh, bx:bx + bw]).sum() / max(1, contorno.sum())
                dentro = bx >= px0 - 3 and by >= py0 - 3 and bx + bw <= px1 + 3 and by + bh <= py1 + 3
                if toque < 0.25 and not (dentro and toque > 0.02):
                    continue
                if reg.sum() > 4 * area_p:
                    continue
                union = pm.copy()
                union[by:by + bh, bx:bx + bw] |= reg
                if union.sum() > 0.6 * ph * pw:
                    continue
                pu_, _, _ = _marco(union)
                # foto ya completa: solo absorbe lo que está dentro de su marco
                if (dentro and toque > 0.02) or (puntaje < 0.9 and (pu_ >= 0.9 or pu_ >= puntaje + 0.03)):
                    pm, absorbida[j], cambio = union, True, True
                    puntaje, area_p = pu_, union.sum()
                    ys, xs = np.nonzero(pm)
                    px0, py0, px1, py1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1
                    dil = cv2.dilate(pm, np.ones((7, 7), np.uint8))
            if not cambio:
                break
        fotos_m[fi] = pm

    # 4. lo que no es forma ni foto (degradados, zonas suaves de las fotos)
    cubierto = np.zeros((ph, pw), np.uint8)
    for j, (f, bx, by, reg) in enumerate(candidatas):
        if not absorbida[j]:
            cubierto[by:by + reg.shape[0], bx:bx + reg.shape[1]] |= reg
    todas_fotos = np.zeros((ph, pw), np.uint8)
    for pm in fotos_m:
        todas_fotos |= pm
    resto = ((cubierto == 0) & (todas_fotos == 0)).astype(np.uint8)
    resto = cv2.morphologyEx(resto, cv2.MORPH_OPEN, np.ones((5, 5), np.uint8))
    n_r, etq_r, st_r, _ = cv2.connectedComponentsWithStats(resto, connectivity=8)
    sueltas = []
    for i in range(1, n_r):
        if st_r[i, cv2.CC_STAT_AREA] < 0.002 * ph * pw:
            continue
        comp = (etq_r == i).astype(np.uint8)
        unida = False
        for fi, pm in enumerate(fotos_m):  # zona suave pegada a una foto = parte de la foto
            if (comp & cv2.dilate(pm, np.ones((9, 9), np.uint8))).any():
                fotos_m[fi] = pm | comp
                unida = True
                break
        if not unida:
            sueltas.append(comp)

    # 5. recorte final de cada foto con su silueta real
    gris = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY).astype(np.float32)
    fotos: list[Foto] = []
    zonas_foto = np.zeros((h, w), np.uint8)
    for es_foto, pm in [(True, p) for p in fotos_m] + [(False, p) for p in sueltas]:
        pm = _rellenar(cv2.morphologyEx(pm, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8)))
        if es_foto:
            pm = _silueta(crudo, pm, 3 * BLOQUE)
        puntaje, tipo, dat = _marco(pm)
        if tipo == "rect" and puntaje >= 0.85:
            x0, y0, x1, y1 = _ajustar_borde(gris, [v / s for v in dat], int(BLOQUE / s) + 3)
            if x1 - x0 < 4 or y1 - y0 < 4:
                continue
            fotos.append(Foto((x0, y0, x1, y1), rgb[y0:y1, x0:x1].copy(), True))
            zonas_foto[y0:y1, x0:x1] = 1
            continue
        if tipo == "elipse" and puntaje >= 0.9:
            cx, cy, ew, eh, ang = dat
            alfa = np.zeros((h, w), np.uint8)
            cv2.ellipse(alfa, ((cx / s, cy / s), (ew / s, eh / s), ang), 255, -1, cv2.LINE_AA)
        else:
            alfa = cv2.resize(pm.astype(np.float32), (w, h), interpolation=cv2.INTER_LINEAR)
            alfa = (np.clip(cv2.GaussianBlur(alfa, (3, 3), 0), 0, 1) * 255).astype(np.uint8)
        ys, xs = np.nonzero(alfa)
        if len(xs) == 0:
            continue
        x0, y0, x1, y1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1
        fotos.append(Foto((x0, y0, x1, y1), np.dstack([rgb[y0:y1, x0:x1], alfa[y0:y1, x0:x1]]), False))
        zonas_foto |= (alfa > 127).astype(np.uint8)

    # 6. formas: nada que caiga dentro de una foto se vectoriza
    formas = []
    for j, (f, bx, by, reg) in enumerate(candidatas):
        if absorbida[j]:
            continue
        x0, y0, x1, y1 = [int(v) for v in f.caja]
        zona = zonas_foto[max(0, y0):y1, max(0, x0):x1]
        if zona.size and zona.mean() > 0.5:
            dentro = cv2.resize(reg, (max(1, x1 - x0), max(1, y1 - y0)), interpolation=cv2.INTER_NEAREST)
            dentro = dentro[: zona.shape[0], : zona.shape[1]]
            if (dentro & zona).sum() > 0.6 * max(1, dentro.sum()):
                continue
        formas.append(f)

    # 7. fondo
    borde_px = np.concatenate([rgb[0], rgb[-1], rgb[:, 0], rgb[:, -1]])
    color_fondo = tuple(int(v) for v in np.median(borde_px, axis=0))
    if formas:
        grande = max(formas, key=lambda f: f.area)
        if grande.area > 0.25 * w * h:
            color_fondo = grande.color
            if (grande.caja[0] <= 3 and grande.caja[1] <= 3
                    and grande.caja[2] >= w - 3 and grande.caja[3] >= h - 3):
                formas.remove(grande)

    # 8. iconos: piezas pequeñas cercanas -> un grupo
    margen = 0.012 * float(np.hypot(w, h))
    peques = [i for i, f in enumerate(formas) if f.area < 0.025 * w * h]
    padre = {i: i for i in peques}

    def raiz(i):
        while padre[i] != i:
            padre[i] = padre[padre[i]]
            i = padre[i]
        return i

    for a_i, i in enumerate(peques):
        ca = formas[i].caja
        for j in peques[a_i + 1:]:
            cb = formas[j].caja
            if (ca[0] - margen <= cb[2] and cb[0] - margen <= ca[2]
                    and ca[1] - margen <= cb[3] and cb[1] - margen <= ca[3]):
                padre[raiz(i)] = raiz(j)
    grupos: dict[int, list[int]] = {}
    for i in peques:
        grupos.setdefault(raiz(i), []).append(i)
    n_grupo = 0
    for miembros in grupos.values():
        if len(miembros) < 2:
            continue
        x0 = min(formas[i].caja[0] for i in miembros); y0 = min(formas[i].caja[1] for i in miembros)
        x1 = max(formas[i].caja[2] for i in miembros); y1 = max(formas[i].caja[3] for i in miembros)
        if (x1 - x0) * (y1 - y0) < 0.08 * w * h:
            n_grupo += 1
            for i in miembros:
                formas[i].grupo = n_grupo

    formas.sort(key=lambda f: -f.area)
    fotos.sort(key=lambda f: -(f.caja[2] - f.caja[0]) * (f.caja[3] - f.caja[1]))
    return color_fondo, formas, fotos
