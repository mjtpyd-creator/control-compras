# -*- coding: utf-8 -*-
"""Registra tickets escaneados en el libro de control.

Acepta un ticket suelto o un lote (lista JSON). En lote el registro es
TODO O NADA: si algun ticket no valida, no se escribe ninguno.

De una sola pasada:
  1. valida los datos de cada ticket,
  2. rechaza duplicados (contra el libro y dentro del lote),
  3. asigna los IDs correlativos,
  4. escribe la cabecera en 'Tickets' y una fila por producto en 'Compras',
  5. archiva cada escaneo como tickets/archivados/AAAA-MM-DD_comercio_total.ext,
  6. informa de si las lineas cuadran con el total del papel.

Uso:
  python registrar_ticket.py ticket.json [opciones]
  python registrar_ticket.py lote.json   [opciones]     # lista de tickets

Opciones:
  --libro RUTA            libro a modificar (por defecto control-compras.xlsx)
  --carpeta-tickets RUTA  raiz de los escaneos (por defecto ./tickets)
  --permitir-duplicado    registra aunque detecte posibles duplicados
  --sin-archivar          no mueve los archivos escaneados
  --sin-copia             no guarda copia de seguridad del libro antes de escribir
  --seco                  simula sin escribir ni mover nada

Formato del JSON (importes numericos; fechas AAAA-MM-DD):
{
  "id": "TK-0007",                       // opcional: si falta se asigna solo
  "fecha": "2026-03-15",
  "comercio": "Supermercado Central",
  "numero": "A-1024",                    // n. de ticket/factura (si no hay: "-")
  "forma_pago": "Tarjeta debito",
  "moneda": "ARS",
  "total": 8.00,                         // total impreso en el ticket
  "iva_global": 21,                      // opcional: % de IVA por defecto
  "iva_importe": 1.39,                   // opcional: IVA impreso (se coteja)
  "archivo": "tickets/entrada/foto.jpg", // opcional: se archiva
  "notas": "compra semanal",
  "lineas": [
    {"descripcion": "Leche 1 L", "cantidad": 2, "unidad": "ud",
     "precio": 1.20, "descuento": 0, "iva": 10, "categoria": "Supermercado"},
    {"descripcion": "Descuento general", "cantidad": 1, "unidad": "ud",
     "precio": -1.00, "categoria": "Supermercado"}
  ]
}

Una linea con 'precio' negativo es una linea de descuento (como las que
imprimen algunos tickets en un renglon aparte).
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import sys
from types import SimpleNamespace

from openpyxl import load_workbook

import comun as co

AVISO_DIF = 0.02  # tolerancia de redondeo entre el papel y las lineas


class ErrorTicket(Exception):
    """Error de validacion: no se registra nada."""


# ------------------------------------------------------------------ lectura

def leer_json(ruta):
    if ruta == "-":
        return json.load(sys.stdin)
    if not os.path.exists(ruta):
        raise ErrorTicket(f"no encuentro el archivo de datos: {ruta}")
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def cargar_libro(ruta):
    if not os.path.exists(ruta):
        raise ErrorTicket(f"no existe el libro: {ruta}")
    wb = load_workbook(ruta)
    for hoja in ("Compras", "Tickets", "Listas"):
        if hoja not in wb.sheetnames:
            raise ErrorTicket(f"el libro no tiene la hoja {hoja}")
    return wb, wb["Compras"], wb["Tickets"], wb["Listas"]


def leer_registrados(ws_t):
    """Tickets ya presentes en la hoja Tickets."""
    campos = ("ID ticket", "Fecha", "Comercio", "Total ticket", "N. ticket",
              "Archivo escaneado")
    idx = {c: co.IT[c] for c in campos}
    registrados = []
    filas = ws_t.iter_rows(min_row=2, max_row=co.FILAS, min_col=1,
                           max_col=co.IT["Archivo escaneado"] + 1, values_only=True)
    for r, fila in enumerate(filas, start=2):
        idt = fila[idx["ID ticket"]]
        if idt in (None, ""):
            continue
        registrados.append({
            "fila": r, "id": str(idt), "fecha": fila[idx["Fecha"]],
            "comercio": fila[idx["Comercio"]], "total": fila[idx["Total ticket"]],
            "numero": fila[idx["N. ticket"]], "archivo": fila[idx["Archivo escaneado"]],
        })
    return registrados


def siguiente_numero(registrados):
    mayor = 0
    for t in registrados:
        m = re.match(r"^TK-(\d+)$", t["id"].strip(), re.IGNORECASE)
        if m:
            mayor = max(mayor, int(m.group(1)))
    return mayor + 1


def leer_lista(ws_l, columna, capacidad):
    """Valores utiles de una columna de Listas."""
    valores = []
    for r in range(2, capacidad + 2):
        v = ws_l.cell(row=r, column=columna).value
        if v not in (None, ""):
            valores.append((r, co.normalizar(v)))
    return valores


def canonico(nombre, valores):
    """Valor canonico de Listas que coincide (exacto o sin mayusculas)."""
    buscado = co.normalizar(nombre)
    for _, v in valores:
        if v == buscado:
            return v
    for _, v in valores:
        if v.lower() == buscado.lower():
            return v
    return None


def asegurar_categoria(ws_l, nombre, valores, capacidad):
    """Devuelve la categoria canonica; si es nueva la anade a Listas."""
    canon = canonico(nombre, valores)
    if canon:
        return canon, None
    libre = None
    for r in range(2, capacidad + 2):
        if ws_l.cell(row=r, column=1).value in (None, ""):
            libre = r
            break
    if libre is None:
        raise ErrorTicket(
            f"la categoria '{nombre}' es nueva y la lista ya esta llena "
            f"({capacidad} categorias). Ampliala en la hoja Listas."
        )
    ws_l.cell(row=libre, column=1, value=co.normalizar(nombre))
    valores.append((libre, co.normalizar(nombre)))
    return co.normalizar(nombre), libre


# ------------------------------------------------------------------ validacion

def numero_o_error(valor, campo, permite_vacio=False):
    if valor in (None, "") and permite_vacio:
        return None
    try:
        return round(float(valor), 2)
    except (TypeError, ValueError):
        raise ErrorTicket(f"'{campo}' debe ser un numero, llego {valor!r}")


def validar(datos, ws_l, categorias, pagos, unidades, monedas):
    """Normaliza y comprueba un ticket. Devuelve el ticket limpio."""
    if not isinstance(datos, dict):
        raise ErrorTicket("cada ticket debe ser un objeto JSON")

    faltan = [c for c in ("fecha", "comercio", "total", "lineas")
              if datos.get(c) in (None, "")]
    if faltan:
        raise ErrorTicket("faltan campos obligatorios: " + ", ".join(faltan))

    try:
        fecha = co.formatear_fecha(datos["fecha"])
    except ValueError as e:
        raise ErrorTicket(str(e))

    comercio = co.normalizar(datos["comercio"])
    if not comercio:
        raise ErrorTicket("'comercio' no puede estar vacio")

    numero = co.normalizar(datos.get("numero")) or "-"
    total = numero_o_error(datos["total"], "total")
    if total is None or total <= 0:
        raise ErrorTicket(f"'total' debe ser mayor que cero, llego {datos['total']!r}")

    pago = datos.get("forma_pago")
    if not pago:
        raise ErrorTicket("falta 'forma_pago'")
    pago_canon = canonico(pago, pagos)
    if not pago_canon:
        raise ErrorTicket("forma de pago no valida: " + repr(pago)
                          + ". Validas: " + ", ".join(v for _, v in pagos))

    moneda = co.normalizar(datos.get("moneda")) or (monedas[0][1] if monedas else "ARS")
    moneda_canon = canonico(moneda, monedas)
    if not moneda_canon:
        raise ErrorTicket("moneda no valida: " + repr(moneda)
                          + ". Validas: " + ", ".join(v for _, v in monedas))

    iva_global = numero_o_error(datos.get("iva_global"), "iva_global", permite_vacio=True)
    if iva_global is not None and not 0 <= iva_global <= 100:
        raise ErrorTicket(f"'iva_global' fuera de rango: {iva_global}")
    iva_importe = numero_o_error(datos.get("iva_importe"), "iva_importe", permite_vacio=True)
    if iva_importe is not None and iva_importe < 0:
        raise ErrorTicket(f"'iva_importe' no puede ser negativo: {iva_importe}")

    lineas_crudas = datos["lineas"]
    if not isinstance(lineas_crudas, list) or not lineas_crudas:
        raise ErrorTicket("'lineas' debe ser una lista con al menos un producto")

    lineas = []
    nuevas_categorias = []
    for i, cruda in enumerate(lineas_crudas, start=1):
        if not isinstance(cruda, dict):
            raise ErrorTicket(f"la linea {i} debe ser un objeto")
        desc = co.normalizar(cruda.get("descripcion"))
        if not desc:
            raise ErrorTicket(f"la linea {i} no tiene 'descripcion'")
        cant = numero_o_error(cruda.get("cantidad", 1), f"linea {i}: cantidad")
        precio = numero_o_error(cruda.get("precio"), f"linea {i}: precio")
        if cant is None or cant <= 0:
            raise ErrorTicket(f"la linea {i}: 'cantidad' debe ser mayor que cero")
        if precio is None:
            raise ErrorTicket(f"la linea {i}: falta 'precio'")
        descu = numero_o_error(cruda.get("descuento", 0), f"linea {i}: descuento") or 0
        importe = round(cant * precio - descu, 2)
        if precio < 0:
            # linea de descuento impresa aparte
            if cant != 1 or descu:
                raise ErrorTicket(
                    f"la linea {i}: una linea de descuento (precio negativo) debe llevar "
                    "cantidad 1 y descuento 0")
        elif importe < 0:
            raise ErrorTicket(
                f"la linea {i}: el descuento ({descu}) supera el importe "
                f"({round(cant * precio, 2)})")

        iva = cruda.get("iva", iva_global)
        iva = numero_o_error(iva, f"linea {i}: iva", permite_vacio=True)
        if iva is not None and not 0 <= iva <= 100:
            raise ErrorTicket(f"la linea {i}: 'iva' fuera de rango ({iva})")

        unid = co.normalizar(cruda.get("unidad", "ud")) or "ud"
        unid_canon = canonico(unid, unidades)
        if not unid_canon:
            raise ErrorTicket(f"la linea {i}: unidad no valida {unid!r}. "
                              "Validas: " + ", ".join(v for _, v in unidades))

        cat = co.normalizar(cruda.get("categoria"))
        if not cat:
            raise ErrorTicket(f"la linea {i} no tiene 'categoria'")
        cat_canon, nueva = asegurar_categoria(ws_l, cat, categorias, co.CAP_CAT)
        if nueva and cat_canon not in nuevas_categorias:
            nuevas_categorias.append(cat_canon)

        lineas.append({
            "descripcion": desc, "cantidad": cant, "unidad": unid_canon,
            "precio": precio, "descuento": descu, "iva": iva, "categoria": cat_canon,
            "importe": importe,
        })

    registrado = round(sum(l["importe"] for l in lineas), 2)

    # --- cotejo del IVA impreso, solo si todas las lineas comparten tipo
    tipos = {l["iva"] for l in lineas}
    iva_calculado = None
    iva_aviso = None
    if len(tipos) == 1 and None not in tipos:
        iva_calculado = round(sum(l["importe"] - l["importe"] / (1 + (l["iva"] or 0) / 100)
                                  for l in lineas), 2)
        if iva_importe is not None and abs(iva_calculado - iva_importe) > 0.05:
            iva_aviso = (f"el IVA impreso ({iva_importe:.2f}) no coincide con el "
                         f"calculado ({iva_calculado:.2f}): revisa el tipo de IVA")

    return {
        "id": co.normalizar(datos.get("id")) or None,
        "fecha": fecha, "comercio": comercio, "numero": numero,
        "forma_pago": pago_canon, "moneda": moneda_canon, "total": total,
        "archivo": datos.get("archivo"), "notas": co.normalizar(datos.get("notas")),
        "lineas": lineas, "registrado": registrado,
        "diferencia": round(total - registrado, 2),
        "iva_importe": iva_importe, "iva_calculado": iva_calculado,
        "iva_aviso": iva_aviso, "categorias_nuevas": nuevas_categorias,
        "duplicado_de": None, "filas_compras": [], "fila_ticket": None,
    }


# ------------------------------------------------------------------ lote

def resolver(ruta):
    if not ruta:
        return None
    return ruta if os.path.isabs(ruta) else os.path.join(co.BASE, ruta)


def preparar_lote(lista, ws, wt, wl, args):
    """Valida todo el lote y comprueba duplicados SIN escribir nada."""
    categorias = leer_lista(wl, 1, co.CAP_CAT)
    pagos = leer_lista(wl, 2, co.CAP_PAGO)
    unidades = leer_lista(wl, 3, co.CAP_UNID)
    monedas = leer_lista(wl, 4, 11)

    # 1) validacion de cada ticket
    preparados, errores = [], []
    for i, datos in enumerate(lista, start=1):
        try:
            preparados.append(validar(datos, wl, categorias, pagos, unidades, monedas))
        except ErrorTicket as e:
            errores.append(f"ticket {i}: {e}")
    if errores:
        raise ErrorTicket("no se ha registrado nada. Problemas encontrados:\n  - "
                          + "\n  - ".join(errores))

    registrados = leer_registrados(wt)

    # 2) IDs correlativos (los explicitos mandan)
    usados = {t["id"].lower() for t in registrados}
    siguiente = siguiente_numero(registrados)
    for tk in preparados:
        if tk["id"]:
            if tk["id"].lower() in usados:
                raise ErrorTicket(f"el ID {tk['id']} ya esta en uso")
            usados.add(tk["id"].lower())
        else:
            while f"tk-{siguiente:04d}" in usados:
                siguiente += 1
            tk["id"] = f"TK-{siguiente:04d}"
            usados.add(tk["id"].lower())
            siguiente += 1

    # 3) duplicados: contra lo ya registrado y dentro del propio lote
    conocidos = {}
    for t in registrados:
        try:
            conocidos.setdefault(
                co.clave_duplicado(t["fecha"], t["comercio"], t["total"], t["numero"]),
                f"{t['id']} (ya en el libro)")
        except (ValueError, TypeError):
            continue
    archivos = {resolver(t["archivo"]) for t in registrados if t["archivo"]}
    repetidos = []
    for tk in preparados:
        clave = co.clave_duplicado(tk["fecha"], tk["comercio"], tk["total"], tk["numero"])
        motivo = conocidos.get(clave)
        ruta_nueva = resolver(tk["archivo"])
        if motivo is None and ruta_nueva and ruta_nueva in archivos:
            motivo = "el mismo archivo escaneado ya estaba registrado"
        if motivo:
            tk["duplicado_de"] = motivo
            repetidos.append(f"{tk['id']} repite {motivo}")
        conocidos[clave] = f"{tk['id']} (en este mismo lote)"
        if ruta_nueva:
            archivos.add(ruta_nueva)
    if repetidos and not args.permitir_duplicado:
        raise ErrorTicket(
            "POSIBLE DUPLICADO, no se ha registrado nada:\n  - "
            + "\n  - ".join(repetidos)
            + "\nSi de verdad son compras distintas, repite con --permitir-duplicado.")

    # 4) los escaneos deben existir antes de mover nada
    if not args.sin_archivar:
        for tk in preparados:
            ruta = resolver(tk["archivo"])
            if ruta and not os.path.exists(ruta):
                raise ErrorTicket(f"{tk['id']}: no encuentro el archivo escaneado {ruta}")

    # 5) asignar filas libres una a una (el libro puede tener huecos)
    cursor_c, cursor_t = 2, 2
    try:
        for tk in preparados:
            tk["filas_compras"] = []
            for _ in tk["lineas"]:
                r = co.fila_libre(ws, cursor_c)
                tk["filas_compras"].append(r)
                cursor_c = r + 1
            tk["fila_ticket"] = co.fila_libre(wt, cursor_t)
            cursor_t = tk["fila_ticket"] + 1
    except ValueError as e:
        raise ErrorTicket(str(e))
    return preparados


# ------------------------------------------------------------------ archivo

def archivar(origen, fecha, comercio, total, carpeta_tickets):
    """Mueve el escaneo a <carpeta_tickets>/archivados con nombre normalizado."""
    if not origen:
        return None
    ruta = resolver(origen)
    if not os.path.exists(ruta):
        raise ErrorTicket(f"no encuentro el archivo escaneado: {ruta}")
    destino_dir = os.path.join(carpeta_tickets, "archivados")
    os.makedirs(destino_dir, exist_ok=True)
    if os.path.abspath(os.path.dirname(ruta)) == os.path.abspath(destino_dir):
        return ruta
    nombre = co.nombre_archivo(fecha, comercio, total, os.path.splitext(ruta)[1] or ".jpg")
    destino = os.path.join(destino_dir, nombre)
    n = 2
    while os.path.exists(destino):
        raiz, ext = os.path.splitext(nombre)
        destino = os.path.join(destino_dir, f"{raiz}-{n}{ext}")
        n += 1
    shutil.move(ruta, destino)
    return destino


def destino_previsto(origen, fecha, comercio, total, carpeta_tickets):
    """Ruta donde quedaria archivado (para --seco)."""
    if not origen:
        return None
    ext = os.path.splitext(resolver(origen))[1] or ".jpg"
    return os.path.join(carpeta_tickets, "archivados",
                        co.nombre_archivo(fecha, comercio, total, ext))


# ------------------------------------------------------------------ escritura

def escribir_compras(ws, tk):
    """Escribe las lineas del ticket. Devuelve la ultima fila usada."""
    for i, l in enumerate(tk["lineas"], start=1):
        r = tk["filas_compras"][i - 1]
        valores = {
            "ID ticket": tk["id"], "Linea": i, "Fecha": tk["fecha"],
            "Comercio": tk["comercio"], "Categoria": l["categoria"],
            "Descripcion": l["descripcion"], "Cant.": l["cantidad"],
            "Unid.": l["unidad"], "Precio unit.": l["precio"],
            "Descuento": l["descuento"], "IVA %": l["iva"],
            "Forma de pago": tk["forma_pago"], "Moneda": tk["moneda"],
            "N. ticket": tk["numero"], "Archivo": tk["archivo"], "Notas": tk["notas"],
        }
        for campo, valor in valores.items():
            c = ws.cell(row=r, column=co.IC[campo] + 1, value=valor)
            if campo == "Fecha":
                c.number_format = co.FECHA
            elif campo in ("Precio unit.", "Descuento"):
                c.number_format = co.MONEDA
            elif campo == "IVA %":
                c.number_format = '0.0"%"'
        if r > co.FILAS_FORM:  # filas sin formulas de plantilla
            ws.cell(row=r, column=co.IC["Mes"] + 1, value=(
                f'=IF(${co.cl("Fecha")}{r}="","",YEAR(${co.cl("Fecha")}{r})&"-"'
                f'&TEXT(MONTH(${co.cl("Fecha")}{r}),"00"))'))
            ws.cell(row=r, column=co.IC["Importe"] + 1, value=(
                f'=IF(${co.cl("Cant.")}{r}="","",ROUND(${co.cl("Cant.")}{r}'
                f'*${co.cl("Precio unit.")}{r}-N(${co.cl("Descuento")}{r}),2))'
            )).number_format = co.MONEDA
            ws.cell(row=r, column=co.IC["IVA $"] + 1, value=(
                f'=IF(OR(${co.cl("Importe")}{r}="",${co.cl("IVA %")}{r}=""),"",'
                f'ROUND(${co.cl("Importe")}{r}-${co.cl("Importe")}{r}'
                f'/(1+${co.cl("IVA %")}{r}/100),2))')).number_format = co.MONEDA
    return tk["filas_compras"][-1]


def escribir_ticket(wt, tk):
    fila = tk["fila_ticket"]
    valores = {
        "ID ticket": tk["id"], "Fecha": tk["fecha"], "Comercio": tk["comercio"],
        "Total ticket": tk["total"], "N. ticket": tk["numero"],
        "Forma de pago": tk["forma_pago"], "Moneda": tk["moneda"],
        "Archivo escaneado": tk["archivo"], "Notas": tk["notas"],
    }
    for campo, valor in valores.items():
        c = wt.cell(row=fila, column=co.IT[campo] + 1, value=valor)
        if campo == "Fecha":
            c.number_format = co.FECHA
        elif campo == "Total ticket":
            c.number_format = co.MONEDA
    if fila > co.FILAS_TK:
        wt.cell(row=fila, column=co.IT["Total registrado"] + 1, value=(
            f'=IF(${co.tl("ID ticket")}{fila}="","",ROUND(SUMIFS('
            f'Compras!${co.cl("Importe")}$2:${co.cl("Importe")}${co.FILAS},'
            f'Compras!${co.cl("ID ticket")}$2:${co.cl("ID ticket")}${co.FILAS},'
            f'${co.tl("ID ticket")}{fila}),2))')).number_format = co.MONEDA
        wt.cell(row=fila, column=co.IT["Diferencia"] + 1, value=(
            f'=IF(${co.tl("ID ticket")}{fila}="","",ROUND('
            f'${co.tl("Total ticket")}{fila}-${co.tl("Total registrado")}{fila},2))'
        )).number_format = co.MONEDA
        wt.cell(row=fila, column=co.IT["Estado"] + 1, value=(
            f'=IF(${co.tl("ID ticket")}{fila}="","",IF(ABS('
            f'${co.tl("Diferencia")}{fila})<=0.02,"OK","REVISAR"))'))
        wt.cell(row=fila, column=co.IT["Duplicado"] + 1, value=(
            f'=IF(${co.tl("ID ticket")}{fila}="","",IF(COUNTIFS('
            f'${co.tl("Fecha")}$2:${co.tl("Fecha")}${co.DUP_HASTA},${co.tl("Fecha")}{fila},'
            f'${co.tl("Comercio")}$2:${co.tl("Comercio")}${co.DUP_HASTA},${co.tl("Comercio")}{fila},'
            f'${co.tl("Total ticket")}$2:${co.tl("Total ticket")}${co.DUP_HASTA},'
            f'${co.tl("Total ticket")}{fila},'
            f'${co.tl("N. ticket")}$2:${co.tl("N. ticket")}${co.DUP_HASTA},'
            f'${co.tl("N. ticket")}{fila})>1,"DUPLICADO",""))'))


# ------------------------------------------------------------------ informe

def informe(preparados, args):
    for i, tk in enumerate(preparados, start=1):
        if len(preparados) > 1:
            print(f"--- ticket {i} de {len(preparados)} ---")
        print(f"{'SIMULADO' if args.seco else 'REGISTRADO'}: {tk['id']}")
        print(f"  Fecha:      {tk['fecha'].isoformat()}")
        print(f"  Comercio:   {tk['comercio']}"
              + (f"  (n. {tk['numero']})" if tk["numero"] != "-" else ""))
        print(f"  Pago:       {tk['forma_pago']} en {tk['moneda']}")
        filas = tk["filas_compras"]
        rango = (f"{filas[0]}-{filas[-1]}" if filas[-1] - filas[0] == len(filas) - 1
                 else ", ".join(str(x) for x in filas))
        print(f"  Lineas:     {len(tk['lineas'])}  -> filas {rango} de Compras, "
              f"fila {tk['fila_ticket']} de Tickets")
        print(f"  Total papel:      {tk['total']:.2f}")
        print(f"  Total lineas:     {tk['registrado']:.2f}")
        print(f"  Diferencia:       {tk['diferencia']:+.2f}  ->  "
              f"{'OK' if abs(tk['diferencia']) <= AVISO_DIF else 'REVISAR'}")
        descuentos = [l for l in tk["lineas"] if l["precio"] < 0]
        if descuentos:
            print(f"  Descuentos en linea propia: {len(descuentos)}")
        if tk["iva_calculado"] is not None:
            print(f"  IVA calculado (incluido): {tk['iva_calculado']:.2f}")
        if tk["categorias_nuevas"]:
            print(f"  Categorias nuevas anadidas a Listas: {', '.join(tk['categorias_nuevas'])}")
        if tk["iva_aviso"]:
            print("  AVISO: " + tk["iva_aviso"])
        if tk["duplicado_de"]:
            print(f"  AVISO: posible duplicado de {tk['duplicado_de']} "
                  "(registrado de todos modos)")
        if tk["archivo"]:
            print(f"  Escaneo:    {co.ruta_legible(tk['archivo'])}")

    if len(preparados) > 1:
        lineas = sum(len(tk["lineas"]) for tk in preparados)
        importe = round(sum(tk["registrado"] for tk in preparados), 2)
        revisar = len([tk for tk in preparados if abs(tk["diferencia"]) > AVISO_DIF])
        print()
        print(f"== {len(preparados)} tickets, {lineas} lineas, {importe:.2f} en total"
              + (f", {revisar} con diferencias" if revisar else "") + " ==")


# ------------------------------------------------------------------ principal

def registrar(lista, libro=co.XLSX, carpeta_tickets=co.DIR_TICKETS,
              permitir_duplicado=False, sin_archivar=False, sin_copia=False,
              seco=False):
    """Registra una lista de tickets. Devuelve (preparados, copia).

    Es el nucleo compartido por la terminal y por la web: valida, rechaza
    duplicados, asigna IDs, escribe las lineas y la cabecera, archiva los
    escaneos, guarda una copia de seguridad y actualiza el manifiesto.
    Lanza ErrorTicket si el lote no se puede registrar (nada queda a medias).
    """
    if not lista:
        raise ErrorTicket("no hay ningun ticket para registrar")

    args = SimpleNamespace(permitir_duplicado=permitir_duplicado,
                           sin_archivar=sin_archivar, seco=seco)
    wb, ws, wt, wl = cargar_libro(libro)
    preparados = preparar_lote(lista, ws, wt, wl, args)

    anotaciones = []
    for tk in preparados:
        if seco:
            tk["archivo"] = destino_previsto(tk["archivo"], tk["fecha"], tk["comercio"],
                                             tk["total"], carpeta_tickets) or tk["archivo"]
            continue
        if tk["archivo"] and not sin_archivar:
            tk["archivo"] = archivar(tk["archivo"], tk["fecha"], tk["comercio"],
                                     tk["total"], carpeta_tickets)
            anotaciones.append({
                "id": tk["id"], "fecha": tk["fecha"].isoformat(),
                "archivo": co.ruta_legible(tk["archivo"]),
                "bytes": os.path.getsize(tk["archivo"]),
                "sha256": co.sha256(tk["archivo"]),
                "registrado_en": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
        tk["archivo"] = co.ruta_legible(tk["archivo"]) if tk["archivo"] else None

    for tk in preparados:
        escribir_compras(ws, tk)
        escribir_ticket(wt, tk)

    copia = None
    if not seco:
        if not sin_copia:
            copia = co.copia_seguridad(libro, "registro")
        wb.save(libro)
        co.anotar_manifiesto(anotaciones, os.path.join(carpeta_tickets, "manifiesto.csv"))

    return preparados, copia


def main():
    ap = argparse.ArgumentParser(
        description="Registra uno o varios tickets escaneados.")
    ap.add_argument("json_ticket", help="archivo JSON con un ticket o una lista")
    ap.add_argument("--libro", default=co.XLSX)
    ap.add_argument("--carpeta-tickets", default=co.DIR_TICKETS)
    ap.add_argument("--permitir-duplicado", action="store_true")
    ap.add_argument("--sin-archivar", action="store_true")
    ap.add_argument("--sin-copia", action="store_true")
    ap.add_argument("--seco", action="store_true")
    args = ap.parse_args()

    datos = leer_json(args.json_ticket)
    lista = datos if isinstance(datos, list) else [datos]
    preparados, copia = registrar(lista, args.libro, args.carpeta_tickets,
                                  args.permitir_duplicado, args.sin_archivar,
                                  args.sin_copia, args.seco)
    informe(preparados, args)
    if copia:
        print(f"Copia de seguridad del libro anterior: {co.ruta_legible(copia)}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ErrorTicket as e:
        print("ERROR: " + str(e), file=sys.stderr)
        sys.exit(2)
    except json.JSONDecodeError as e:
        print(f"ERROR: el JSON no es valido ({e})", file=sys.stderr)
        sys.exit(2)
