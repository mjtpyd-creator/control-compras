# -*- coding: utf-8 -*-
"""Prediccion de compras a reponer, segun el historial de tickets.

Agrupa las lineas de Compras por producto (descripcion normalizada) y, para los
que se compraron al menos dos veces en fechas distintas, estima:
  - la frecuencia (mediana de los dias entre compras),
  - la proxima fecha de compra,
  - cuantos dias faltan (o cuantos van de atraso).

Uso:
  python prediccion.py                 # lista "que comprar" para hoy
  python prediccion.py --dias 7        # horizonte (por defecto 7 dias)
  python prediccion.py --libro X.xlsx

Se excluyen las lineas de descuento (precio negativo) y las que no tienen fecha.
"""
import argparse
import datetime as dt
import sys
from collections import Counter, defaultdict
from statistics import median

from openpyxl import load_workbook

import comun as co


def normalizar_producto(descripcion):
    """Clave de agrupado: minusculas y espacios colapsados."""
    return " ".join(str(descripcion or "").split()).lower()


def eventos_productos(libro=co.XLSX):
    """Agrupa las lineas por producto y fecha de compra.

    Devuelve dict: clave -> {"original", "unidades", "eventos": {fecha: cantidad}}.
    """
    wb = load_workbook(libro)
    lineas = co.leer_tabla(wb["Compras"], co.IC)
    wb.close()

    productos = defaultdict(lambda: {"original": "", "unidades": Counter(),
                                     "eventos": defaultdict(float)})
    for l in lineas:
        precio = l["Precio unit."] or 0
        if precio < 0:
            continue  # renglones de descuento, no productos
        f = l["Fecha"]
        if isinstance(f, dt.datetime):
            f = f.date()
        if not isinstance(f, dt.date):
            continue
        clave = normalizar_producto(l["Descripcion"])
        if not clave:
            continue
        g = productos[clave]
        g["original"] = co.normalizar(l["Descripcion"]) or clave
        g["unidades"][co.normalizar(l["Unid."]) or "ud"] += 1
        g["eventos"][f] += float(l["Cant."] or 0)
    return productos


def sugerencias(libro=co.XLSX, hoy=None, horizonte=7):
    """Lista de productos con su prediccion de reposicion."""
    hoy = hoy or dt.date.today()
    if isinstance(hoy, str):
        hoy = dt.date.fromisoformat(hoy)
    productos = eventos_productos(libro)

    previstos = []
    sin_historial = []
    for clave, g in productos.items():
        fechas = sorted(g["eventos"].keys())
        unidades = g["unidades"].most_common(1)[0][0] if g["unidades"] else "ud"
        cantidad_total = round(sum(g["eventos"].values()), 2)
        base = {"producto": g["original"] or clave, "clave": clave, "unidad": unidades,
                "veces": len(fechas), "cantidad_total": cantidad_total,
                "ultima": fechas[-1].isoformat()}
        if len(fechas) < 2:
            sin_historial.append(base)
            continue
        intervalos = [(fechas[i + 1] - fechas[i]).days for i in range(len(fechas) - 1)]
        intervalo = max(1, int(round(median(intervalos))))
        proxima = fechas[-1] + dt.timedelta(days=intervalo)
        span = max(1, (fechas[-1] - fechas[0]).days)
        ritmo_semana = round(cantidad_total / span * 7, 1)
        previstos.append({
            **base,
            "frecuencia_dias": intervalo,
            "proxima": proxima.isoformat(),
            "dias": (proxima - hoy).days,
            "ritmo_semana": ritmo_semana,
        })

    previstos.sort(key=lambda p: p["dias"])
    atrasados = [p for p in previstos if p["dias"] <= 0]
    esta_semana = [p for p in previstos if 0 < p["dias"] <= horizonte]
    proximamente = [p for p in previstos if horizonte < p["dias"] <= horizonte + 7]
    return {
        "hoy": hoy.isoformat(),
        "horizonte_dias": horizonte,
        "n_productos": len(productos),
        "atrasados": atrasados,
        "esta_semana": esta_semana,
        "proximamente": proximamente,
        "sin_historial": sin_historial,
    }


def _linea(p):
    detalle = f"cada {p['frecuencia_dias']} dias"
    return (f"- {p['producto']}  ({p['unidad']}, {detalle})  "
            f"ultima {p['ultima']}  ->  {p['dias']} dias")


def imprimir(s):
    print("QUE COMPRAR - segun tu historial")
    print(f"  Hoy: {s['hoy']}  ·  horizonte: {s['horizonte_dias']} dias  ·  "
          f"productos con historial: {s['n_productos']}")
    print()

    def bloque(titulo, lista, color=None):
        if not lista:
            return
        print(titulo + f" ({len(lista)})")
        for p in lista:
            if "frecuencia_dias" in p:
                print(_linea(p))
            else:
                print(f"- {p['producto']}  ({p['unidad']}, comprado {p['veces']} vez)")
        print()

    bloque("ATRASADOS (ya deberia estar comprado)", s["atrasados"])
    bloque(f"ESTA SEMANA (proximos {s['horizonte_dias']} dias)", s["esta_semana"])
    bloque("PROXIMAMENTE (8-14 dias)", s["proximamente"])
    bloque("SIN FRECUENCIA AUN (comprados 1 sola vez)", s["sin_historial"])

    if not any([s["atrasados"], s["esta_semana"], s["proximamente"], s["sin_historial"]]):
        print("  Todavia no hay compras cargadas como para predecir.")


def proximos_dias_compra(libro=co.XLSX, hoy=None, dias_semana=None):
    """
    Devuelve qué comprar antes de cada día de compra (ej: martes=1, sábado=5).
    
    dias_semana: lista de enteros 0=lunes...6=domingo. Default [1, 5] = martes y sábado.
    """
    if dias_semana is None:
        dias_semana = [1, 5]  # martes y sábado
    
    s = sugerencias(libro, hoy)
    hoy_dt = dt.date.fromisoformat(s["hoy"]) if isinstance(s["hoy"], str) else s["hoy"]
    
    # Calcular próximas fechas de cada día de compra
    proximas_fechas = {}
    for dia in dias_semana:
        # Python: lunes=0, domingo=6
        dias_hasta = (dia - hoy_dt.weekday()) % 7
        if dias_hasta == 0:
            dias_hasta = 7  # si hoy es el día, mostrar la próxima semana
        proximas_fechas[dia] = hoy_dt + dt.timedelta(days=dias_hasta)
    
    # Agrupar productos por día de compra
    por_dia = {dia: [] for dia in dias_semana}
    atrasados = []
    
    for p in s["atrasados"] + s["esta_semana"] + s["proximamente"]:
        if "frecuencia_dias" not in p:
            continue
        prox = dt.date.fromisoformat(p["proxima"])
        # Asignar al día de compra más cercano ANTES o EN la fecha de reposición
        mejor_dia = None
        mejor_diff = 999
        for dia, fecha_compra in proximas_fechas.items():
            diff = (prox - fecha_compra).days
            if diff <= 0:  # fecha de compra es después de la reposición = tarde
                if diff > -7:  # pero no más de una semana tarde
                    # Asignar al día anterior de compra
                    dias_previos = [d for d in dias_semana if d != dia]
                    if dias_previos:
                        mejor_dia = max(dias_previos, key=lambda d: proximas_fechas[d])
            elif diff < mejor_diff:  # compra antes de necesitar
                mejor_diff = diff
                mejor_dia = dia
        
        if mejor_dia is not None:
            p["dia_compra"] = mejor_dia
            p["dias_antes"] = (proximas_fechas[mejor_dia] - prox).days
            por_dia[mejor_dia].append(p)
        else:
            atrasados.append(p)
    
    nombres_dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    
    resultado = {
        "hoy": s["hoy"],
        "dias_semana": dias_semana,
        "por_dia": {},
        "atrasados": atrasados,
        "sin_historial": s["sin_historial"]
    }
    
    for dia in dias_semana:
        lista = por_dia[dia]
        lista.sort(key=lambda p: p.get("dias_antes", 999))
        resultado["por_dia"][dia] = {
            "nombre": nombres_dias[dia],
            "fecha": proximas_fechas[dia].isoformat(),
            "productos": lista
        }
    
    return resultado


def imprimir_por_dia(s):
    """Imprime la predicción organizada por día de compra."""
    print("QUE COMPRAR ANTES DE CADA TURNO")
    print(f"  Hoy: {s['hoy']}")
    print()
    
    for dia in s["dias_semana"]:
        info = s["por_dia"][dia]
        print(f"📅 {info['nombre']} {info['fecha']} ({len(info['productos'])} productos)")
        for p in info["productos"]:
            freq = f"cada {p['frecuencia_dias']} días" if "frecuencia_dias" in p else "?"
            print(f"  - {p['producto']} ({p['unidad']}, {freq}) - última: {p['ultima']}")
        print()
    
    if s["atrasados"]:
        print("⚠️ ATRASADOS (ya deberían estar comprados)")
        for p in s["atrasados"]:
            freq = f"cada {p['frecuencia_dias']} días" if "frecuencia_dias" in p else "?"
            print(f"  - {p['producto']} ({p['unidad']}, {freq}) - última: {p['ultima']}")
        print()
    
    if s["sin_historial"]:
        print("📝 SIN FRECUENCIA (comprados 1 sola vez)")
        for p in s["sin_historial"]:
            print(f"  - {p['producto']} ({p['unidad']})")


def main():
    ap = argparse.ArgumentParser(description="Predice que compras reponer.")
    ap.add_argument("--libro", default=co.XLSX)
    ap.add_argument("--dias", type=int, default=7, help="horizonte en dias")
    ap.add_argument("--hoy", default=None, help="fecha de referencia AAAA-MM-DD")
    ap.add_argument("--dias-semana", nargs="+", type=int, default=[1, 5],
                    help="días de compra 0=lunes..6=domingo (default: 1 5 = martes sábado)")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    if args.dias_semana:
        s = proximos_dias_compra(args.libro, args.hoy, args.dias_semana)
        if args.json:
            import json
            print(json.dumps(s, ensure_ascii=False, indent=2))
        else:
            imprimir_por_dia(s)
    else:
        s = sugerencias(args.libro, args.hoy, args.dias)
        if args.json:
            import json
            print(json.dumps(s, ensure_ascii=False, indent=2))
        else:
            imprimir(s)
    return 0


if __name__ == "__main__":
    sys.exit(main())
