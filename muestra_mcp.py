#!/usr/bin/env python3
"""Apoyo para relevar AI Overviews desde Claude Code usando el MCP de DataForSEO.

El MCP solo expone el metodo Live Advanced (USD 0.002 por keyword) y responde de
a una keyword por llamada, asi que la muestra se arma incrementalmente: Claude
consulta cada keyword, anota el resultado con `anotar` y al final `excel` arma
la planilla con las columnas originales.

Comandos:
    python3 muestra_mcp.py siguientes 8        -> proximas keywords a consultar
    python3 muestra_mcp.py anotar kw=Si kw=No  -> guarda resultados
    python3 muestra_mcp.py estado              -> avance, gasto y margen de error
    python3 muestra_mcp.py excel               -> genera el Excel de la muestra
"""

import csv
import os
import sys
from collections import Counter, OrderedDict
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_io import read_xlsx, write_xlsx

BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_KEYWORDS = os.path.join(BASE, "data", "Keywords_Ecommerce.xlsx")
ARCHIVO_ORDEN = os.path.join(BASE, "data", "orden_keywords.csv")
DIR_ESTADO = os.path.join(BASE, "estado")
ARCHIVO_MUESTRA = os.path.join(DIR_ESTADO, "muestra_mcp.csv")
DIR_SALIDAS = os.path.join(BASE, "salidas")

COSTO_LIVE = 0.002        # USD por keyword, SERP Live Advanced (lo que cobra el MCP)
MERCADO = "Argentina / es / mobile"

CAMPOS = ["keyword", "ai_overview", "fecha_utc", "mercado"]


def orden_filas():
    with open(ARCHIVO_ORDEN, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def orden_keywords():
    return [r["keyword"] for r in orden_filas()]


def pesos_verticales():
    """Participacion real de cada vertical en las 13.065 keywords."""
    filas = orden_filas()
    tot = len(filas)
    cuenta = Counter(r["vertical"] for r in filas)
    return {v: n / tot for v, n in cuenta.items()}, cuenta


def vertical_por_keyword():
    return {r["keyword"].strip().lower(): r["vertical"] for r in orden_filas()}


def cargar():
    """keyword normalizada -> Si/No, en orden de relevamiento."""
    hechas = OrderedDict()
    if os.path.exists(ARCHIVO_MUESTRA):
        with open(ARCHIVO_MUESTRA, encoding="utf-8-sig", newline="") as fh:
            for fila in csv.DictReader(fh):
                kw = (fila.get("keyword") or "").strip().lower()
                if kw and fila.get("ai_overview") in ("Si", "No"):
                    hechas[kw] = fila["ai_overview"]
    return hechas


def cmd_siguientes(n, balanceado=True):
    """Proximas keywords a consultar.

    Con cuota pareja (`balanceado`) se rota entre verticales, tomando de cada
    uno en el orden estratificado ya versionado. Asi cada vertical acumula una
    cantidad parecida de keywords y se puede leer su porcentaje por separado,
    incluso en los verticales chicos como Autos y Motos (1,4% del archivo).
    El total global se recompone despues re-ponderando por el peso real de
    cada vertical: ver cmd_estado.
    """
    hechas = cargar()
    colas = OrderedDict()
    for r in orden_filas():
        if r["keyword"].strip().lower() in hechas:
            continue
        colas.setdefault(r["vertical"], []).append(r["keyword"])

    if not balanceado:
        faltan = [kw for kw in orden_keywords() if kw.strip().lower() not in hechas]
        for kw in faltan[:n]:
            print(kw)
        return

    # Se arranca por el vertical con menos keywords ya relevadas, para que la
    # cuota se empareje aunque una tanda se haya cortado por la mitad.
    v_de = vertical_por_keyword()
    ya = Counter(v_de.get(kw, "") for kw in hechas)
    salida = []
    while len(salida) < n and any(colas.values()):
        for v in sorted(colas, key=lambda v: (ya[v], v)):
            if not colas[v] or len(salida) >= n:
                continue
            salida.append(colas[v].pop(0))
            ya[v] += 1
    for kw in salida:
        print(kw)


def cmd_anotar(pares):
    hechas = cargar()
    os.makedirs(DIR_ESTADO, exist_ok=True)
    existe = os.path.exists(ARCHIVO_MUESTRA)
    ahora = datetime.now(timezone.utc).isoformat(timespec="seconds")
    nuevas = repetidas = 0

    with open(ARCHIVO_MUESTRA, "a", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=CAMPOS)
        if not existe:
            w.writeheader()
        for par in pares:
            kw, _, val = par.rpartition("=")
            kw, val = kw.strip(), val.strip()
            if val not in ("Si", "No"):
                raise SystemExit("Valor invalido en %r: se espera Si o No." % par)
            if not kw:
                raise SystemExit("Falta la keyword en %r." % par)
            if kw.lower() in hechas:      # nunca pagar ni anotar dos veces
                repetidas += 1
                continue
            w.writerow({"keyword": kw, "ai_overview": val,
                        "fecha_utc": ahora, "mercado": MERCADO})
            hechas[kw.lower()] = val
            nuevas += 1

    print("Anotadas %d nuevas%s. Total en la muestra: %d"
          % (nuevas, (", %d repetidas ignoradas" % repetidas) if repetidas else "", len(hechas)))


Z = 1.96   # 95% de confianza


def intervalo_wilson(n, x):
    """Intervalo de confianza al 95% para una proporcion, en porcentaje.

    Se usa Wilson y no el clasico p +/- z*sqrt(p(1-p)/n) porque ese ultimo
    colapsa a +/- 0 cuando la muestra da 0% o 100%, que es justo lo que pasa
    con los verticales chicos al principio del relevamiento. Con 0 de 5 la
    respuesta honesta no es "0% +/- 0", es "entre 0% y 43%".
    """
    if n == 0:
        return 0.0, 100.0
    p = x / n
    denom = 1 + Z * Z / n
    centro = (p + Z * Z / (2 * n)) / denom
    semi = (Z / denom) * ((p * (1 - p) / n + Z * Z / (4 * n * n)) ** 0.5)
    return 100 * max(0.0, centro - semi), 100 * min(1.0, centro + semi)


def p_ajustada(n, x):
    """Proporcion de Agresti-Coull: (x + z^2/2) / (n + z^2).

    Evita que un estrato con 0 o 100% aporte varianza cero al total ponderado.
    """
    if n == 0:
        return 0.5
    return (x + Z * Z / 2) / (n + Z * Z)


def estimacion_ponderada(por_vertical):
    """(porcentaje, margen) del total, re-pesando cada vertical por su peso real.

    La muestra tiene cuota pareja, no proporcional, asi que el promedio simple
    sobreestimaria los verticales chicos. El estimador correcto es
    p = suma(W_v * p_v), con W_v = participacion del vertical en las 13.065.
    El error se propaga como suma(W_v^2 * p_v*(1-p_v)/n_v).
    """
    pesos, _ = pesos_verticales()
    p = var = 0.0
    peso_cubierto = 0.0
    for v, (n, con) in por_vertical.items():
        if not n:
            continue
        w = pesos.get(v, 0.0)
        p += w * (con / n)
        # La varianza usa la proporcion ajustada para que un estrato con 0%
        # o 100% no aporte incertidumbre cero.
        pa = p_ajustada(n, con)
        var += (w ** 2) * pa * (1 - pa) / n
        peso_cubierto += w
    if peso_cubierto <= 0:
        return 0.0, 100.0, 0.0
    # Se renormaliza por si algun vertical todavia no tiene ninguna keyword.
    p /= peso_cubierto
    margen = 100 * Z * (var ** 0.5) / peso_cubierto
    return 100 * p, margen, 100 * peso_cubierto


def desglose(hechas):
    v_de = vertical_por_keyword()
    por_vertical = OrderedDict()
    for kw, val in hechas.items():
        v = v_de.get(kw, "(sin dato)")
        n, con = por_vertical.get(v, (0, 0))
        por_vertical[v] = (n + 1, con + (1 if val == "Si" else 0))
    return por_vertical


def cmd_estado():
    hechas = cargar()
    n = len(hechas)
    con = sum(1 for v in hechas.values() if v == "Si")
    print("Keywords relevadas : %d de 13.065" % n)
    print("Gasto estimado     : USD %.4f  (USD %.4f por keyword, Live Advanced)"
          % (n * COSTO_LIVE, COSTO_LIVE))
    if not n:
        return

    por_vertical = desglose(hechas)
    print("\nPor vertical (cuota pareja):")
    _, cuenta = pesos_verticales()
    for v, (nv, cv) in sorted(por_vertical.items(), key=lambda x: -x[1][0]):
        lo, hi = intervalo_wilson(nv, cv)
        print("  %-26s %4d kw   %5.1f%% con AIO   IC95%%: %4.1f%% - %4.1f%%   (vertical real: %d kw)"
              % (v, nv, 100.0 * cv / nv, lo, hi, cuenta.get(v, 0)))

    p, margen, cubierto = estimacion_ponderada(por_vertical)
    lo, hi = intervalo_wilson(n, con)
    print("\nCrudo de la muestra       : %.1f%% con AI Overview (%d de %d), IC95%%: %.1f%% - %.1f%%"
          % (100.0 * con / n, con, n, lo, hi))
    print("Estimacion para las 13.065: %.1f%% +/- %.1f pp (95%%), re-ponderado por vertical"
          % (p, margen))
    if cubierto < 99.9:
        print("  (cubre el %.1f%% del archivo: faltan verticales sin relevar)" % cubierto)


def cmd_excel():
    hechas = cargar()
    if not hechas:
        raise SystemExit("Todavia no hay ninguna keyword relevada.")

    headers, filas = read_xlsx(ARCHIVO_KEYWORDS)
    por_kw = {f[0].strip().lower(): f for f in filas}

    salida = []
    for kw in hechas:                       # respeta el orden de relevamiento
        f = por_kw.get(kw)
        if not f:
            continue
        vol = str(f[3]).strip()
        try:
            vol = str(int(float(vol))) if vol else ""
        except ValueError:
            pass
        salida.append([f[0], f[1], f[2], vol, hechas[kw]])

    os.makedirs(DIR_SALIDAS, exist_ok=True)
    ruta = os.path.join(DIR_SALIDAS, "AIO_muestra_mcp.xlsx")
    write_xlsx(ruta, headers, salida, sheet_name="Muestra", numeric_cols=[3])

    with open(ruta[:-5] + ".csv", "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(headers)
        w.writerows(salida)

    print("Excel  : %s" % ruta)
    print("CSV    : %s\n" % (ruta[:-5] + ".csv"))
    cmd_estado()

    print("\nPor tipo de keyword:")
    tot = Counter(f[2] for f in salida)
    si = Counter(f[2] for f in salida if f[4] == "Si")
    for clave, cant in tot.most_common():
        lo, hi = intervalo_wilson(cant, si[clave])
        print("  %-26s %4d kw   %5.1f%% con AIO   IC95%%: %4.1f%% - %4.1f%%"
              % (clave, cant, 100.0 * si[clave] / cant, lo, hi))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    cmd = sys.argv[1]
    if cmd == "siguientes":
        cmd_siguientes(int(sys.argv[2]) if len(sys.argv) > 2 else 10)
    elif cmd == "anotar":
        cmd_anotar(sys.argv[2:])
    elif cmd == "estado":
        cmd_estado()
    elif cmd == "excel":
        cmd_excel()
    else:
        raise SystemExit(__doc__)
