"""
Uso desde la terminal:

    python convertir.py portada.svg interior.svg -o presentacion.pptx
    python convertir.py flyer.png -o flyer.pptx --fuente "Montserrat" --fuente-titulos "Montserrat Black"
    python convertir.py flyer.png -o flyer.pptx --vectorizar
    python convertir.py *.svg -o deck.pptx --tamano real
"""
import argparse
import glob
import sys
from pathlib import Path

from diseno_a_pptx import OpcionesImagen, convertir_archivos


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap =argparse.ArgumentParser(description="Convierte SVG o imágenes en un PowerPoint editable.")
    ap.add_argument("entradas", nargs="+", help="Archivos .svg/.png/.jpg (uno por diapositiva)")
    ap.add_argument("-o", "--salida", default="presentacion.pptx")
    ap.add_argument("--tamano", choices=["estandar", "real"], default="estandar",
                    help="estandar = 13,33 in de ancho (16:9); real = tamaño del diseño a 96 ppp")
    ap.add_argument("--fuente", default="Arial", help="Fuente para textos detectados en imágenes")
    ap.add_argument("--fuente-titulos", default=None, help="Fuente para títulos detectados en imágenes")
    ap.add_argument("--vectorizar", action="store_true",
                    help="(imágenes) convierte zonas de color plano en formas editables")
    ap.add_argument("--ocr", choices=["auto", "rapidocr", "tesseract"], default="auto")
    a = ap.parse_args()

    rutas = []
    for e in a.entradas:
        rutas += sorted(glob.glob(e)) or [e]
    archivos = [(Path(r).name, Path(r).read_bytes()) for r in rutas]
    op = OpcionesImagen(fuente_texto=a.fuente, fuente_titulos=a.fuente_titulos,
                        vectorizar=a.vectorizar, motor_ocr=a.ocr)
    datos, inf = convertir_archivos(archivos, a.tamano, op, carpeta_base=Path(rutas[0]).parent)
    Path(a.salida).write_bytes(datos)

    print(f"✔ Guardado: {a.salida}  ({len(archivos)} diapositiva(s))")
    print("  Elementos:", ", ".join(f"{v} {k}" for k, v in inf.conteo.items()) or "-")
    for nombre, ok in inf.fuentes.items():
        print(f"  {'✔' if ok else '✘'} Fuente {nombre}{'' if ok else '  (no instalada)'}")
    for av in inf.avisos:
        print("  ⚠", av)


if __name__ == "__main__":
    sys.exit(main())
