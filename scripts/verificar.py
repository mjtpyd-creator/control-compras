# -*- coding: utf-8 -*-
"""Auditoria del registro de compras y del archivo de escaneos.

No necesita abrir Excel ni recalcular nada: lee el libro y lo comprueba por su
cuenta (implementacion independiente de las formulas), y ademas revisa que los
escaneos archivados sigan intactos.

Comprueba:
  - que cada ticket cuadre con la suma de sus lineas,
  - que no haya duplicados,
  - que no haya lineas sin categoria, sin fecha, sin cantidad ni precio,
  - que cada escaneo archivado siga existiendo y con la misma huella SHA-256,
  - que no haya escaneos archivados que ningun ticket reclame,
y resume el gasto del ejercicio por mes y por categoria.

Uso:
  python verificar.py                        # libro y archivo de produccion
  python verificar.py --libro X --carpeta-tickets Y
  python verificar.py --json                 # salida para programas

Sale con codigo 1 si encuentra problemas de integridad.
"""
import argparse
import datetime as dt
import json
import os
import sys
from collections import defaultdict

from openpyxl import load_workbook

import comun as co

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def fecha_de(valor):
    if isinstance(valor, dt.datetime):
        return valor.date()
    return valor if isinstance(valor, dt.date) else None


def auditar(libro, carpeta_tickets, ruta_manifiesto, anio=None):
    """Devuelve el informe completo como diccionario."""
    problemas, avisos = [], []

    wb = load_workbook(libro)
    compras = co.leer_tabla(wb["Compras"], co.IC)
    tickets = co.leer_tabla(wb["Tickets"], co.IT)
    wb.close()

    # ---------------------------------------------------- lineas
    por_id = defaultdict(list)
    for l in compras:
        idt = co.normalizar(l["ID ticket"])
        por_id[idt].append(l)
        if not co.normalizar(l["Categoria"]):
            problemas.append(f"Compras fila {l['_fila']}: linea sin categoria "
                             f"({co.normalizar(l['Descripcion'])!r})")
        if fecha_de(l["Fecha"]) is None:
            problemas.append(f"Compras fila {l['_fila']}: linea sin fecha valida")
        if l["Cant."] in (None, "") or l["Precio unit."] in (None, ""):
            problemas.append(f"Compras fila {l['_fila']}: falta la cantidad o el precio")

    # ---------------------------------------------------- tickets
    registrados = {}
    for t in tickets:
        idt = co.normalizar(t["ID ticket"])
        if idt in registrados:
            problemas.append(f"Tickets {idt}: el ID esta repetido en la hoja")
        registrados[idt] = t
        lineas = por_id.get(idt, [])
        if not lineas:
            problemas.append(f"Tickets {idt}: no tiene ninguna linea en Compras")
        total = t["Total ticket"] or 0
        suma = round(sum(co.importe_linea(l["Cant."], l["Precio unit."], l["Descuento"])
                         for l in lineas), 2)
        dif = round(total - suma, 2)
        t["_suma"], t["_dif"] = suma, dif
        if abs(dif) > 0.02:
            avisos.append(f"REVISAR Tickets {idt} ({co.normalizar(t['Comercio'])}): "
                          f"el papel dice {total:.2f} y las lineas suman {suma:.2f} "
                          f"(diferencia {dif:+.2f})")
        if fecha_de(t["Fecha"]) is None:
            problemas.append(f"Tickets {idt}: sin fecha valida")

    for idt in por_id:
        if idt not in registrados:
            problemas.append(f"Compras: hay lineas del ticket {idt} que no esta en Tickets")

    # ---------------------------------------------------- duplicados
    vistos = defaultdict(list)
    for idt, t in registrados.items():
        try:
            vistos[co.clave_duplicado(t["Fecha"], t["Comercio"], t["Total ticket"],
                                      t["N. ticket"])].append(idt)
        except (ValueError, TypeError):
            problemas.append(f"Tickets {idt}: no se puede calcular la clave de duplicado")
    duplicados = [sorted(ids) for ids in vistos.values() if len(ids) > 1]
    for ids in duplicados:
        avisos.append("REVISAR Posible duplicado: " + ", ".join(ids)
                      + " comparten fecha, comercio, total y numero")

    # ---------------------------------------------------- escaneos
    manifiesto = co.leer_manifiesto(ruta_manifiesto)
    huellas = {co.normalizar(m.get("archivo")): m for m in manifiesto}
    referenciados = set()
    sin_escaneo = 0
    for idt, t in registrados.items():
        arch = co.normalizar(t["Archivo escaneado"])
        if not arch:
            sin_escaneo += 1
            avisos.append(f"Tickets {idt}: sin escaneo archivado")
            continue
        referenciados.add(arch)
        ruta = arch if os.path.isabs(arch) else os.path.join(co.BASE, arch)
        if not os.path.exists(ruta):
            problemas.append(f"Tickets {idt}: falta el escaneo {arch}")
            continue
        entrada = huellas.get(arch)
        if entrada is None:
            avisos.append(f"Tickets {idt}: el escaneo {arch} no figura en el manifiesto")
        elif co.sha256(ruta) != entrada.get("sha256"):
            problemas.append(f"Tickets {idt}: el escaneo {arch} ha cambiado "
                             "(la huella no coincide con el manifiesto)")

    for m in manifiesto:
        arch = co.normalizar(m.get("archivo"))
        ruta = arch if os.path.isabs(arch) else os.path.join(co.BASE, arch)
        if not os.path.exists(ruta):
            problemas.append(f"Manifiesto: falta el escaneo {arch} (ticket {m.get('id')})")

    dir_arch = os.path.join(carpeta_tickets, "archivados")
    huerfanos = []
    if os.path.isdir(dir_arch):
        for nombre in sorted(os.listdir(dir_arch)):
            ruta = os.path.join(dir_arch, nombre)
            if not os.path.isfile(ruta):
                continue
            clave = co.ruta_legible(ruta)
            # un escaneo del manifiesto es historico (ticket anulado o corregido)
            if clave not in referenciados and clave not in huellas:
                huerfanos.append(nombre)
    for nombre in huerfanos:
        avisos.append(f"Archivados: {nombre} no lo reclama ningun ticket "
                      "ni figura en el manifiesto")

    # ---------------------------------------------------- anulaciones
    anulados = co.leer_anulados(os.path.join(carpeta_tickets, "anulados.csv"))
    for a in anulados:
        idt = co.normalizar(a.get("id"))
        if idt and idt in registrados:
            problemas.append(f"Anulados: {idt} figura como anulado pero sigue en el libro")

    # ---------------------------------------------------- resumen propio
    if anio is None:
        anios = [f.year for t in registrados.values()
                 if (f := fecha_de(t["Fecha"])) is not None]
        anio = max(anios) if anios else dt.date.today().year

    total = 0.0
    iva_total = 0.0
    por_mes, por_cat, por_pago = defaultdict(float), defaultdict(float), defaultdict(float)
    lineas_anio = 0
    for idt, t in registrados.items():
        f = fecha_de(t["Fecha"])
        if f is None or f.year != anio:
            continue
        for l in por_id.get(idt, []):
            imp = co.importe_linea(l["Cant."], l["Precio unit."], l["Descuento"])
            total += imp
            lineas_anio += 1
            por_mes[f.month] += imp
            por_cat[co.normalizar(l["Categoria"]) or "(sin categoria)"] += imp
            por_pago[co.normalizar(t["Forma de pago"]) or "(sin forma de pago)"] += imp
            v = co.iva_linea(imp, l["IVA %"])
            if v:
                iva_total += v

    return {
        "libro": libro,
        "ejercicio": anio,
        "tickets": len(registrados),
        "lineas": len(compras),
        "gasto_total": round(total, 2),
        "iva_incluido": round(iva_total, 2),
        "lineas_del_ejercicio": lineas_anio,
        "por_mes": {MESES[m - 1]: round(v, 2) for m, v in sorted(por_mes.items())},
        "por_categoria": {k: round(v, 2) for k, v in
                          sorted(por_cat.items(), key=lambda x: -x[1])},
        "por_forma_pago": {k: round(v, 2) for k, v in
                           sorted(por_pago.items(), key=lambda x: -x[1])},
        "escaneos_sin_registrar": sin_escaneo,
        "anulados": len(anulados),
        "duplicados": duplicados,
        "problemas": problemas,
        "avisos": avisos,
    }


def imprimir(informe):
    print(f"AUDITORIA DEL CONTROL DE COMPRAS")
    print(f"  Libro:      {co.ruta_legible(informe['libro'])}")
    print(f"  Ejercicio:  {informe['ejercicio']}")
    print(f"  Registrado: {informe['tickets']} tickets, {informe['lineas']} lineas, "
          f"gasto {informe['gasto_total']:.2f}, IVA incluido {informe['iva_incluido']:.2f}")
    if informe.get("anulados"):
        print(f"  Anulados:   {informe['anulados']} (los escaneos se conservan)")
    print()
    if informe["por_mes"]:
        print("  Gasto por mes:")
        for mes, v in informe["por_mes"].items():
            print(f"    {mes:<12} {v:>12,.2f}")
        print("  Gasto por categoria:")
        for cat, v in informe["por_categoria"].items():
            print(f"    {cat:<28} {v:>12,.2f}")
        print("  Gasto por forma de pago:")
        for pago, v in informe["por_forma_pago"].items():
            print(f"    {pago:<28} {v:>12,.2f}")
        print()
    if informe["avisos"]:
        print(f"  AVISOS ({len(informe['avisos'])}) - conviene mirarlos:")
        for a in informe["avisos"]:
            print(f"    - {a}")
        print()
    if informe["problemas"]:
        print(f"  PROBLEMAS ({len(informe['problemas'])}) - hay que arreglarlos:")
        for p in informe["problemas"]:
            print(f"    - {p}")
        print()
    if not informe["problemas"] and not informe["avisos"]:
        print("  Sin problemas ni avisos: el registro y el archivo estan coherentes.")
    else:
        print(f"  Resumen: {len(informe['problemas'])} problemas, "
              f"{len(informe['avisos'])} avisos.")


def main():
    ap = argparse.ArgumentParser(description="Audita el registro de compras.")
    ap.add_argument("--libro", default=co.XLSX)
    ap.add_argument("--carpeta-tickets", default=co.DIR_TICKETS)
    ap.add_argument("--manifiesto", default=None)
    ap.add_argument("--ejercicio", type=int, default=None)
    ap.add_argument("--json", action="store_true", help="salida para programas")
    args = ap.parse_args()

    if not os.path.exists(args.libro):
        print(f"ERROR: no existe el libro {args.libro}", file=sys.stderr)
        return 2
    manifiesto = args.manifiesto or os.path.join(args.carpeta_tickets, "manifiesto.csv")

    informe = auditar(args.libro, args.carpeta_tickets, manifiesto, args.ejercicio)
    if args.json:
        print(json.dumps(informe, ensure_ascii=False, indent=2, default=str))
    else:
        imprimir(informe)
    return 1 if informe["problemas"] else 0


if __name__ == "__main__":
    sys.exit(main())
