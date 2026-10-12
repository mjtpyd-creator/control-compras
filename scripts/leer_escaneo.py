# -*- coding: utf-8 -*-
"""Prepara un escaneo para poder leerlo.

Acepta lo que sueltan los escanerres y los moviles:
  - PDF (de una o varias paginas): renderiza cada pagina a PNG.
  - Imagen (JPG, PNG, TIFF...): la usa tal cual; si es enorme, crea una copia
    reducida para que se pueda leer sin perder los datos.

No toca el archivo original: escribe las imagenes en la carpeta de lectura.

Uso:
  python leer_escaneo.py tickets/entrada/factura.pdf
  python leer_escaneo.py foto.jpg --dpi 200 --max-lado 2400
  python leer_escaneo.py tickets/entrada          (todos los archivos de la carpeta)
"""
import argparse
import os
import subprocess
import sys

import comun as co


class ErrorLectura(Exception):
    pass


def listar(origen):
    """Archivos de escaneo encontrados, en orden."""
    if os.path.isdir(origen):
        dentro = []
        for nombre in sorted(os.listdir(origen)):
            ruta = os.path.join(origen, nombre)
            if os.path.isfile(ruta) and os.path.splitext(nombre)[1].lower() in \
                    co.EXT_IMAGEN + co.EXT_PDF:
                dentro.append(ruta)
        if not dentro:
            raise ErrorLectura(f"no hay imagenes ni PDF en {origen}")
        return dentro
    if not os.path.exists(origen):
        raise ErrorLectura(f"no encuentro {origen}")
    return [origen]


def carpeta_libre(base):
    """Carpeta que todavia no existe (el motor de render no sobrescribe)."""
    if not os.path.exists(base):
        return base
    n = 2
    while os.path.exists(f"{base}-{n}"):
        n += 1
    return f"{base}-{n}"


def renderizar_pdf(ruta, destino, dpi):
    """Renderiza un PDF a PNG (una imagen por pagina). Devuelve las rutas."""
    if not os.path.exists(co.NODE) or not os.path.exists(co.LIBREOFFICE_CLI):
        raise ErrorLectura(
            "falta el motor de render del entorno; no puedo convertir el PDF a imagen.\n"
            f"  node: {co.NODE}\n  cli:  {co.LIBREOFFICE_CLI}"
        )
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    r = subprocess.run(
        [co.NODE, co.LIBREOFFICE_CLI, "render", "--input", ruta,
         "--output-dir", destino, "--dpi", str(dpi)],
        capture_output=True, text=True)
    if r.returncode != 0:
        raise ErrorLectura("fallo el render del PDF: " + (r.stderr or r.stdout).strip())

    generadas = sorted(
        os.path.join(destino, n) for n in os.listdir(destino)
        if os.path.splitext(n)[1].lower() in co.EXT_IMAGEN
    )
    if not generadas:
        raise ErrorLectura("el render no produjo ninguna imagen: " + r.stdout.strip())
    return generadas


def reducir_imagen(ruta, destino, max_lado):
    """Copia reducida de una imagen demasiado grande. Devuelve la ruta final."""
    from PIL import Image

    with Image.open(ruta) as img:
        ancho, alto = img.size
        if max(ancho, alto) <= max_lado:
            return ruta
        factor = max_lado / float(max(ancho, alto))
        nueva = img.convert("RGB").resize((max(1, int(ancho * factor)),
                                           max(1, int(alto * factor))), Image.LANCZOS)
    os.makedirs(destino, exist_ok=True)
    salida = os.path.join(destino, os.path.splitext(os.path.basename(ruta))[0]
                          + f"_reducida{max_lado}.png")
    nueva.save(salida)
    return salida


def main():
    ap = argparse.ArgumentParser(description="Prepara escaneos para su lectura.")
    ap.add_argument("origen", help="archivo o carpeta de escaneos")
    ap.add_argument("--salida", default=co.DIR_LECTURA,
                    help="carpeta donde dejar las imagenes")
    ap.add_argument("--dpi", type=int, default=180, help="resolucion al renderizar PDF")
    ap.add_argument("--max-lado", type=int, default=2200,
                    help="lado mayor maximo de las imagenes (px)")
    args = ap.parse_args()

    preparadas = []
    for origen in listar(args.origen):
        nombre = os.path.splitext(os.path.basename(origen))[0]
        ext = os.path.splitext(origen)[1].lower()
        destino = os.path.join(args.salida, nombre)
        if ext in co.EXT_PDF:
            paginas = renderizar_pdf(origen, carpeta_libre(destino), args.dpi)
            for p in paginas:
                preparadas.append(reducir_imagen(p, os.path.dirname(p), args.max_lado))
        else:
            preparadas.append(reducir_imagen(origen, carpeta_libre(destino), args.max_lado))

    print(f"{len(preparadas)} imagen(es) listas para leer:")
    for p in preparadas:
        ancho = alto = "?"
        try:
            from PIL import Image
            with Image.open(p) as img:
                ancho, alto = img.size
        except Exception:
            pass
        print(f"  {co.ruta_legible(p)}  ({ancho}x{alto})")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ErrorLectura as e:
        print("ERROR: " + str(e), file=sys.stderr)
        sys.exit(2)
