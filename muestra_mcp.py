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


def orden_keywords():
    with open(ARCHIVO_ORDEN, encoding="utf-8-sig", newline="") as fh:
        return [r["keyword"] for r in csv.DictReader(fh)]


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


def cmd_siguientes(n):
    hechas = cargar()
    faltan = [kw for kw in orden_keywords() if kw.strip().lower() not in hechas]
    for kw in faltan[:n]:
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


def margen_error(n, p):
    """Semiancho del intervalo de confianza al 95%, en puntos porcentuales."""
    if n == 0:
        return 100.0
    return 100 * 1.96 * ((p * (1 - p) / n) ** 0.5)


def cmd_estado():
    hechas = cargar()
    n = len(hechas)
    con = sum(1 for v in hechas.values() if v == "Si")
    print("Keywords relevadas : %d de 13.065" % n)
    print("Gasto estimado     : USD %.4f  (USD %.4f por keyword, Live Advanced)"
          % (n * COSTO_LIVE, COSTO_LIVE))
    if n:
        p = con / n
        print("Con AI Overview    : %d  (%.1f%%)" % (con, 100 * p))
        print("Sin AI Overview    : %d  (%.1f%%)" % (n - con, 100 * (1 - p)))
        print("Margen de error    : +/- %.1f puntos porcentuales (95%% de confianza)"
              % margen_error(n, p))


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

    con = sum(1 for f in salida if f[4] == "Si")
    n = len(salida)
    print("Excel  : %s" % ruta)
    print("CSV    : %s" % (ruta[:-5] + ".csv"))
    print("Filas  : %d  |  con AI Overview: %d (%.1f%%)  +/- %.1f pp"
          % (n, con, 100.0 * con / n, margen_error(n, con / n)))

    for etiqueta, idx in (("Por vertical", 1), ("Por tipo de keyword", 2)):
        print("\n%s:" % etiqueta)
        tot = Counter(f[idx] for f in salida)
        si = Counter(f[idx] for f in salida if f[4] == "Si")
        for clave, cant in tot.most_common():
            print("  %-26s %4d kw   %5.1f%% con AIO" % (clave, cant, 100.0 * si[clave] / cant))


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
