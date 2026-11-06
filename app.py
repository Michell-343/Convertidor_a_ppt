"""
App web para convertir diseños en PowerPoint editables.

Ejecutar:   streamlit run app.py
"""
import base64
import hashlib
import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from diseno_a_pptx import CatalogoFuentes, OpcionesImagen, convertir_archivos, es_svg

st.set_page_config(page_title="Diseño → PowerPoint editable", page_icon="🎨", layout="wide")

EXT_FUENTES = (".ttf", ".otf", ".ttc", ".otc")
MAX_FUENTES_MB = 150


@st.cache_resource(show_spinner="Leyendo las tipografías…", max_entries=20)
def obtener_catalogo(clave: str, carpeta: str | None) -> CatalogoFuentes:
    """Fuentes del sistema + las subidas. `clave` identifica el conjunto subido."""
    return CatalogoFuentes([carpeta] if carpeta else None)


def guardar_fuentes(subidas) -> tuple[str, str | None, int]:
    """Guarda las fuentes subidas (sueltas o en .zip) en una carpeta temporal propia."""
    if not subidas:
        return "", None, 0
    h = hashlib.sha256()
    for f in subidas:
        h.update(f.name.encode())
        h.update(f.getvalue())
    clave = h.hexdigest()[:16]
    carpeta = Path(tempfile.gettempdir()) / "diseno_a_pptx_fuentes" / clave
    carpeta.mkdir(parents=True, exist_ok=True)
    n, total = 0, 0
    for f in subidas:
        datos = f.getvalue()
        if f.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(datos)) as z:
                    for info in z.infolist():
                        nombre = Path(info.filename).name  # sin carpetas: evita rutas peligrosas
                        if info.is_dir() or not nombre.lower().endswith(EXT_FUENTES) or nombre.startswith("."):
                            continue
                        total += info.file_size
                        if total > MAX_FUENTES_MB * 1024 * 1024:
                            break
                        (carpeta / nombre).write_bytes(z.read(info))
                        n += 1
            except zipfile.BadZipFile:
                st.sidebar.error(f"«{f.name}» no es un .zip válido.")
        elif f.name.lower().endswith(EXT_FUENTES):
            (carpeta / Path(f.name).name).write_bytes(datos)
            n += 1
    return clave, str(carpeta), n


# ------------------------------------------------------------------ barra lateral
with st.sidebar:
    st.header("Tus tipografías")
    subidas = st.file_uploader(
        "Sube las fuentes del diseño (.ttf, .otf o un .zip)",
        type=["ttf", "otf", "ttc", "otc", "zip"], accept_multiple_files=True,
        help="Sirven para medir los textos y colocarlos exactamente donde estaban. "
             "Sube todas las variantes que uses (Regular, Bold, SemiBold…). "
             "No se guardan ni se comparten: solo se usan durante esta sesión.")
    clave, carpeta, n_subidas = guardar_fuentes(subidas)
    cat = obtener_catalogo(clave, carpeta)
    if n_subidas:
        subidas_fam = sorted({f.familia for f in cat.fuentes if carpeta and f.ruta.startswith(carpeta)})
        st.success(f"{n_subidas} archivo(s) de fuente: " + ", ".join(subidas_fam))
    st.caption(f"{len(cat.familias())} familias disponibles. Arial, Times New Roman y Courier New "
               "se miden con equivalentes exactos aunque no estén instaladas.")
    familias = sorted(set(cat.familias()) | {"Arial", "Times New Roman", "Courier New"}, key=str.lower)

    st.divider()
    st.header("Opciones")
    tamano = st.radio(
        "Tamaño de la diapositiva",
        ["estandar", "real"],
        format_func=lambda v: {"estandar": "Estándar PowerPoint (13,33 in de ancho)",
                               "real": "Tamaño real del diseño (96 ppp)"}[v],
        help="La proporción del diseño siempre se respeta. «Estándar» es lo que "
             "el cliente espera al abrir el archivo; los tamaños de letra se escalan.",
    )
    ordenar = st.checkbox("Ordenar diapositivas por nombre de archivo", value=True)

    st.divider()
    st.subheader("Solo para PNG / JPG")
    st.caption("Una imagen no guarda qué fuente se usó: indícala aquí.")

    def idx(nombre):
        return familias.index(nombre) if nombre in familias else 0

    fuente_texto = st.selectbox("Fuente de los textos", familias, index=idx("Arial"),
                                accept_new_options=True,
                                help="Elige de la lista o escribe el nombre exacto de la fuente.")
    usar_titulos = st.checkbox("Usar otra fuente para los títulos")
    fuente_titulos = st.selectbox("Fuente de los títulos", familias, index=idx(fuente_texto),
                                  disabled=not usar_titulos, accept_new_options=True)
    vectorizar = st.checkbox(
        "Separar en elementos editables (vectorizar)",
        help="Figuras simples → rectángulos y elipses de PowerPoint · otras formas → curvas "
             "editables · iconos → grupos · fotos y degradados → imágenes sueltas recortadas · "
             "fondo → rectángulo de color.")
    colores = st.slider("Colores a detectar", 4, 16, 10, disabled=not vectorizar)

# ------------------------------------------------------------------ principal
st.title("🎨 Diseño → PowerPoint editable")
st.write("Sube tus diseños (**SVG** para máxima fidelidad, o PNG/JPG). "
         "Cada archivo será una diapositiva con textos, formas y colores editables.")

with st.expander("💡 Cómo exportar el SVG para obtener el mejor resultado"):
    st.markdown("""
**Illustrator** → *Archivo › Exportar › Exportar como… › SVG*
- Estilo: **CSS interno** o **Atributos de presentación**
- Fuente: **SVG** (¡no *Convertir en contornos*!, o el texto dejará de ser editable)
- Imágenes: **Incrustar**
- Desmarca *Minimizar*; deja *Adaptable* activado o desactivado, da igual.

**Figma** → selecciona el frame › *Export › SVG* y **desmarca «Outline text»**.

**Canva** → exporta como **SVG** (plan Pro). Canva suele convertir los textos en contornos;
si pasa, sube un PNG y usa el OCR.

**Inkscape** → *Guardar como › SVG simple*, sin convertir los objetos a trayectos.

**Tipografías:** súbelas en la barra lateral para que los textos queden en su posición exacta.
Además deben estar **instaladas** en el equipo que abra el PowerPoint (el tuyo y el del cliente).
""")

archivos = st.file_uploader("Arrastra aquí tus archivos",
                            type=["svg", "png", "jpg", "jpeg", "webp"],
                            accept_multiple_files=True)

if archivos:
    lista = [(a.name, a.getvalue()) for a in archivos]
    if ordenar:
        lista.sort(key=lambda x: x[0].lower())

    cols = st.columns(min(4, len(lista)))
    for i, (nombre, datos) in enumerate(lista):
        with cols[i % len(cols)]:
            if es_svg(nombre, datos):
                b64 = base64.b64encode(datos).decode()
                st.markdown(f'<img src="data:image/svg+xml;base64,{b64}" '
                            f'style="width:100%;border:1px solid #ddd;border-radius:6px">',
                            unsafe_allow_html=True)
            else:
                st.image(datos, width="stretch")
            st.caption(f"{i + 1}. {nombre}")

    if st.button("Convertir a PowerPoint", type="primary", width="stretch"):
        op = OpcionesImagen(fuente_texto=fuente_texto,
                            fuente_titulos=fuente_titulos if usar_titulos else None,
                            vectorizar=vectorizar, colores_vectorizar=colores)
        try:
            with st.spinner("Convirtiendo… (la primera imagen tarda más: se carga el OCR)"):
                pptx, informe = convertir_archivos(lista, tamano, op, cat=cat)
        except Exception as e:  # mostrar el error a la usuaria sin cerrar la app
            st.error(f"No se pudo convertir: {e}")
            st.stop()

        nombre_salida = (Path(lista[0][0]).stem if len(lista) == 1 else "presentacion") + ".pptx"
        st.success(f"¡Listo! {len(lista)} diapositiva(s).")
        st.download_button("⬇️ Descargar " + nombre_salida, pptx, file_name=nombre_salida,
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                           type="primary", width="stretch")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Formas", informe.conteo.get("formas", 0))
        c2.metric("Iconos (grupos)", informe.conteo.get("iconos", 0))
        c3.metric("Textos editables", informe.conteo.get("textos", 0))
        c4.metric("Fotos / imágenes", informe.conteo.get("fotos", 0) + informe.conteo.get("imágenes", 0))

        if informe.fuentes:
            st.subheader("Tipografías")
            for f, ok in sorted(informe.fuentes.items()):
                st.write(("✅ " if ok else "❌ ") + f + ("" if ok else " — **no disponible: súbela en la barra lateral**"))
        if informe.avisos:
            st.subheader("Avisos")
            for a in informe.avisos:
                st.warning(a)
