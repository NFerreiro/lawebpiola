#!/usr/bin/env python3
"""Detecta si cada keyword dispara un AI Overview en Google, via DataForSEO.

Pensado para correr por tandas con saldo minimo: cada corrida consume solo las
keywords que le indiques, guarda el progreso y genera un Excel con esa tanda.
La siguiente corrida arranca automaticamente donde termino la anterior, aunque
uses otra cuenta de DataForSEO.

Metodo: SERP API de Google Organic, cola Standard (`priority: 1`) con `depth: 10`,
que es la combinacion mas barata: USD 0.0006 por keyword. La presencia del AI
Overview se lee del array `item_types` del resultado, que incluye "ai_overview"
tambien cuando el overview es asincronico (Google lo genera on-demand), sin
necesidad de pagar el extra de `load_async_ai_overview`.

Solo libreria estandar de Python 3.8+.
"""

import argparse
import base64
import csv
import json
import os
import random
import ssl
import sys
import time
import urllib.error
import urllib.request
from collections import Counter, OrderedDict, defaultdict
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_io import read_xlsx, write_xlsx  # noqa: E402

BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_KEYWORDS = os.path.join(BASE, "data", "Keywords_Ecommerce.xlsx")
ARCHIVO_ORDEN = os.path.join(BASE, "data", "orden_keywords.csv")
DIR_ESTADO = os.path.join(BASE, "estado")
DIR_SALIDAS = os.path.join(BASE, "salidas")
ARCHIVO_RESULTADOS = os.path.join(DIR_ESTADO, "resultados.csv")

API = "https://api.dataforseo.com"
COSTO_POR_KEYWORD = 0.0006          # cola Standard, depth 10, resultado Advanced
COSTO_POR_KEYWORD_ASYNC = 0.0012    # con load_async_ai_overview activado
TAREAS_POR_POST = 100               # maximo que acepta DataForSEO por request
SEMILLA_ORDEN = 42                  # fija el orden estratificado para siempre

COL_KEYWORD = 0
COL_VERTICAL = 1
COL_TIPO = 2
COL_VOLUMEN = 3
COL_AIO = 4

CAMPOS_RESULTADO = [
    "keyword", "ai_overview", "corrida", "fecha_utc",
    "location", "language", "device", "item_types", "task_id",
]


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------

def log(msg):
    print("[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg), flush=True)


def normalizar_volumen(v):
    """'190000.0' -> '190000' para que el Excel no muestre decimales de mas."""
    s = str(v).strip()
    if not s:
        return ""
    try:
        f = float(s)
    except ValueError:
        return s
    return str(int(f)) if f == int(f) else s


# --------------------------------------------------------------------------
# Cliente HTTP de DataForSEO
# --------------------------------------------------------------------------

class DataForSEO:
    def __init__(self, login, password, timeout=180):
        token = base64.b64encode(("%s:%s" % (login, password)).encode("utf-8")).decode("ascii")
        self.headers = {
            "Authorization": "Basic " + token,
            "Content-Type": "application/json",
        }
        self.timeout = timeout
        self.ctx = ssl.create_default_context()

    def _request(self, metodo, ruta, payload=None, reintentos=4):
        url = API + ruta
        body = json.dumps(payload).encode("utf-8") if payload is not None else None
        ultimo_error = None
        for intento in range(reintentos):
            req = urllib.request.Request(url, data=body, headers=self.headers, method=metodo)
            try:
                with urllib.request.urlopen(req, timeout=self.timeout, context=self.ctx) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as e:
                detalle = e.read().decode("utf-8", "replace")[:400]
                if e.code == 401:
                    raise SystemExit(
                        "\nERROR 401: usuario o contrasena de DataForSEO invalidos.\n"
                        "Revisa el login (email) y la API password del panel de DataForSEO.\n"
                    )
                if e.code == 402:
                    raise SaldoInsuficiente(
                        "La cuenta de DataForSEO no tiene saldo suficiente (HTTP 402)."
                    )
                # 429 y 5xx son transitorios: conviene reintentar.
                if e.code in (429, 500, 502, 503, 504) and intento < reintentos - 1:
                    espera = 2 ** (intento + 1)
                    log("HTTP %d en %s, reintento en %ds" % (e.code, ruta, espera))
                    time.sleep(espera)
                    ultimo_error = "HTTP %d: %s" % (e.code, detalle)
                    continue
                raise RuntimeError("HTTP %d en %s: %s" % (e.code, ruta, detalle))
            except (urllib.error.URLError, TimeoutError, ssl.SSLError) as e:
                if intento < reintentos - 1:
                    espera = 2 ** (intento + 1)
                    log("Error de red en %s (%s), reintento en %ds" % (ruta, e, espera))
                    time.sleep(espera)
                    ultimo_error = str(e)
                    continue
                raise RuntimeError("Error de red en %s: %s" % (ruta, e))
        raise RuntimeError("Fallaron los reintentos en %s: %s" % (ruta, ultimo_error))

    def saldo(self):
        data = self._request("POST", "/v3/appendix/user_data", [])
        try:
            return float(data["tasks"][0]["result"][0]["money"]["balance"])
        except (KeyError, IndexError, TypeError):
            raise RuntimeError("No se pudo leer el saldo: %s" % json.dumps(data)[:400])

    def crear_tareas(self, tareas):
        return self._request("POST", "/v3/serp/google/organic/task_post", tareas)

    def tareas_listas(self):
        return self._request("GET", "/v3/serp/google/organic/tasks_ready")

    def obtener_tarea(self, task_id):
        return self._request("GET", "/v3/serp/google/organic/task_get/advanced/" + task_id)


class SaldoInsuficiente(Exception):
    pass


# --------------------------------------------------------------------------
# Orden estratificado
# --------------------------------------------------------------------------

def construir_orden(filas):
    """Ordena las keywords de forma que cualquier prefijo sea una muestra
    representativa del total.

    Se agrupa por (Vertical, tipo_keyword), se mezcla cada grupo con una semilla
    fija y se intercalan los grupos por posicion relativa. Resultado: las
    primeras 1.666 keywords respetan la proporcion real de cada vertical y cada
    tipo, asi el % de AI Overviews de la primera corrida ya es extrapolable al
    total. El orden es deterministico: siempre da el mismo resultado.
    """
    grupos = defaultdict(list)
    for i, fila in enumerate(filas):
        grupos[(fila[COL_VERTICAL], fila[COL_TIPO])].append(i)

    con_clave = []
    for clave in sorted(grupos):
        indices = grupos[clave]
        random.Random(SEMILLA_ORDEN).shuffle(indices)
        n = len(indices)
        for pos, idx in enumerate(indices):
            # (pos + 0.5) / n reparte cada grupo de forma pareja sobre [0, 1).
            con_clave.append(((pos + 0.5) / n, clave, idx))

    con_clave.sort(key=lambda x: (x[0], x[1]))
    return [idx for _, _, idx in con_clave]


def cargar_o_crear_orden(filas):
    keywords = [f[COL_KEYWORD] for f in filas]
    if os.path.exists(ARCHIVO_ORDEN):
        with open(ARCHIVO_ORDEN, encoding="utf-8-sig", newline="") as fh:
            orden = [int(r["indice"]) for r in csv.DictReader(fh)]
        if len(orden) == len(filas) and sorted(orden) == list(range(len(filas))):
            return orden
        log("El orden guardado no coincide con el Excel actual: se regenera.")

    orden = construir_orden(filas)
    os.makedirs(os.path.dirname(ARCHIVO_ORDEN), exist_ok=True)
    with open(ARCHIVO_ORDEN, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["posicion", "indice", "keyword", "vertical", "tipo_keyword"])
        for pos, idx in enumerate(orden, start=1):
            w.writerow([pos, idx, keywords[idx], filas[idx][COL_VERTICAL], filas[idx][COL_TIPO]])
    log("Orden estratificado generado en %s" % ARCHIVO_ORDEN)
    return orden


# --------------------------------------------------------------------------
# Estado acumulado entre corridas
# --------------------------------------------------------------------------

def cargar_resultados():
    """keyword normalizada -> dict con el resultado ya obtenido."""
    if not os.path.exists(ARCHIVO_RESULTADOS):
        return OrderedDict()
    hechos = OrderedDict()
    with open(ARCHIVO_RESULTADOS, encoding="utf-8-sig", newline="") as fh:
        for fila in csv.DictReader(fh):
            kw = (fila.get("keyword") or "").strip().lower()
            if kw and fila.get("ai_overview") in ("Si", "No"):
                hechos[kw] = fila
    return hechos


def guardar_resultados(nuevos):
    os.makedirs(DIR_ESTADO, exist_ok=True)
    existe = os.path.exists(ARCHIVO_RESULTADOS)
    with open(ARCHIVO_RESULTADOS, "a", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=CAMPOS_RESULTADO)
        if not existe:
            w.writeheader()
        for fila in nuevos:
            w.writerow(fila)


def siguiente_numero_corrida(hechos):
    n = 0
    for fila in hechos.values():
        try:
            n = max(n, int(fila.get("corrida") or 0))
        except ValueError:
            pass
    return n + 1


# --------------------------------------------------------------------------
# Deteccion del AI Overview
# --------------------------------------------------------------------------

def tiene_ai_overview(resultado):
    """True si el SERP incluye un AI Overview.

    Se mira `item_types`, que DataForSEO completa con "ai_overview" incluso
    cuando el overview es asincronico y su contenido viene vacio. Como respaldo
    tambien se recorre `items` por si la estructura cambia.
    """
    tipos = resultado.get("item_types") or []
    if "ai_overview" in tipos:
        return True
    for item in resultado.get("items") or []:
        if isinstance(item, dict) and item.get("type") == "ai_overview":
            return True
    return False


# --------------------------------------------------------------------------
# Corrida
# --------------------------------------------------------------------------

def postear_tareas(cliente, keywords, cfg, corrida):
    """Envia las keywords a la cola y devuelve {task_id: keyword}."""
    pendientes = {}
    total_lotes = (len(keywords) + TAREAS_POR_POST - 1) // TAREAS_POR_POST

    for n_lote, inicio in enumerate(range(0, len(keywords), TAREAS_POR_POST), start=1):
        lote = keywords[inicio:inicio + TAREAS_POR_POST]
        tareas = []
        for kw in lote:
            tarea = {
                "keyword": kw,
                "location_name": cfg["location"],
                "language_code": cfg["language"],
                "device": cfg["device"],
                "depth": 10,
                "priority": 1,          # 1 = cola Standard, la tarifa mas barata
                "tag": "corrida_%02d" % corrida,
            }
            if cfg["device"] == "mobile":
                tarea["os"] = "android"
            if cfg["load_async"]:
                tarea["load_async_ai_overview"] = True
            tareas.append(tarea)

        respuesta = cliente.crear_tareas(tareas)
        if respuesta.get("status_code") == 40200:
            raise SaldoInsuficiente(respuesta.get("status_message", "Saldo insuficiente"))

        creadas = 0
        for t in respuesta.get("tasks") or []:
            kw_tarea = ((t.get("data") or {}).get("keyword")) or ""
            if t.get("status_code") == 20100 and t.get("id"):
                pendientes[t["id"]] = kw_tarea
                creadas += 1
            else:
                log("  no se pudo crear la tarea de '%s': %s %s"
                    % (kw_tarea, t.get("status_code"), t.get("status_message")))
        log("Lote %d/%d: %d/%d tareas creadas" % (n_lote, total_lotes, creadas, len(lote)))
        time.sleep(0.3)   # margen holgado frente al limite de 2.000 requests/minuto

    return pendientes


def recolectar(cliente, pendientes, espera_max_min):
    """Espera a que la cola resuelva y devuelve (resultados, no_resueltas)."""
    resueltos = {}
    limite = time.time() + espera_max_min * 60
    espera = 20

    while pendientes and time.time() < limite:
        time.sleep(espera)
        espera = min(espera * 1.5, 120)

        try:
            listas = cliente.tareas_listas()
        except RuntimeError as e:
            log("tasks_ready fallo (%s); se reintenta." % e)
            continue

        ids_listos = []
        for t in listas.get("tasks") or []:
            for r in t.get("result") or []:
                if r.get("id") in pendientes:
                    ids_listos.append(r["id"])

        if not ids_listos:
            log("Aun en cola: %d tareas pendientes..." % len(pendientes))
            continue

        for task_id in ids_listos:
            kw = pendientes.pop(task_id, None)
            if kw is None:
                continue
            try:
                data = cliente.obtener_tarea(task_id)
            except RuntimeError as e:
                log("  no se pudo leer la tarea %s (%s)" % (task_id, e))
                pendientes[task_id] = kw   # se reintenta en la proxima vuelta
                continue

            tarea = (data.get("tasks") or [{}])[0]
            resultados = tarea.get("result") or []
            if tarea.get("status_code") != 20000 or not resultados:
                log("  tarea sin resultado para '%s': %s" % (kw, tarea.get("status_message")))
                continue

            resultado = resultados[0]
            resueltos[kw] = {
                "ai_overview": "Si" if tiene_ai_overview(resultado) else "No",
                "item_types": ",".join(resultado.get("item_types") or []),
                "task_id": task_id,
            }

        log("Resueltas %d, pendientes %d" % (len(resueltos), len(pendientes)))

    return resueltos, pendientes


def escribir_csv(ruta, headers, filas):
    """Espejo en CSV de cada Excel, por si preferis importarlo a Sheets o Looker."""
    with open(ruta, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(headers)
        w.writerows(filas)


def escribir_salidas(headers, filas, hechos, keywords_corrida, corrida, cfg):
    os.makedirs(DIR_SALIDAS, exist_ok=True)
    por_kw = {f[COL_KEYWORD].strip().lower(): f for f in filas}

    # Excel de esta corrida: las mismas columnas del original, con la ultima ya completa.
    de_la_corrida = []
    for kw in keywords_corrida:
        clave = kw.strip().lower()
        if clave not in hechos or clave not in por_kw:
            continue
        f = por_kw[clave]
        de_la_corrida.append([
            f[COL_KEYWORD], f[COL_VERTICAL], f[COL_TIPO],
            normalizar_volumen(f[COL_VOLUMEN]), hechos[clave]["ai_overview"],
        ])

    ruta_corrida = os.path.join(DIR_SALIDAS, "AIO_corrida_%02d.xlsx" % corrida)
    write_xlsx(ruta_corrida, headers, de_la_corrida,
               sheet_name="Corrida %02d" % corrida, numeric_cols=[COL_VOLUMEN])
    escribir_csv(ruta_corrida[:-5] + ".csv", headers, de_la_corrida)

    # Excel consolidado con todo lo procesado hasta ahora.
    headers_cons = list(headers) + ["Corrida", "Fecha consulta (UTC)", "Mercado", "Dispositivo"]
    consolidado = []
    for fila in filas:
        clave = fila[COL_KEYWORD].strip().lower()
        if clave not in hechos:
            continue
        h = hechos[clave]
        consolidado.append([
            fila[COL_KEYWORD], fila[COL_VERTICAL], fila[COL_TIPO],
            normalizar_volumen(fila[COL_VOLUMEN]), h["ai_overview"],
            h.get("corrida", ""), (h.get("fecha_utc", "") or "")[:19],
            "%s / %s" % (h.get("location", ""), h.get("language", "")),
            h.get("device", ""),
        ])

    ruta_cons = os.path.join(DIR_SALIDAS, "AIO_consolidado.xlsx")
    write_xlsx(ruta_cons, headers_cons, consolidado,
               sheet_name="Consolidado", numeric_cols=[COL_VOLUMEN])
    escribir_csv(ruta_cons[:-5] + ".csv", headers_cons, consolidado)

    return ruta_corrida, ruta_cons, de_la_corrida, consolidado


def resumen(titulo, filas_res, idx_aio, idx_vertical, idx_tipo):
    total = len(filas_res)
    if not total:
        return
    con = sum(1 for f in filas_res if f[idx_aio] == "Si")
    print("\n" + "=" * 62)
    print(titulo)
    print("=" * 62)
    print("Keywords analizadas : %d" % total)
    print("Con AI Overview     : %d  (%.1f%%)" % (con, 100.0 * con / total))
    print("Sin AI Overview     : %d  (%.1f%%)" % (total - con, 100.0 * (total - con) / total))

    for etiqueta, idx in (("Por vertical", idx_vertical), ("Por tipo de keyword", idx_tipo)):
        print("\n%s:" % etiqueta)
        tot = Counter(f[idx] for f in filas_res)
        si = Counter(f[idx] for f in filas_res if f[idx_aio] == "Si")
        for clave, n in tot.most_common():
            print("  %-26s %5d kw   %5.1f%% con AIO" % (clave or "(sin dato)", n, 100.0 * si[clave] / n))
    print()


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def resolver_credenciales(args):
    login = args.login or os.environ.get("DATAFORSEO_LOGIN")
    password = args.password or os.environ.get("DATAFORSEO_PASSWORD")

    ruta = os.path.join(BASE, "credenciales.txt")
    if (not login or not password) and os.path.exists(ruta):
        with open(ruta, encoding="utf-8") as fh:
            for linea in fh:
                linea = linea.strip()
                if not linea or linea.startswith("#") or ":" not in linea:
                    continue
                l, _, p = linea.partition(":")
                login = login or l.strip()
                password = password or p.strip()
                break

    if not login or not password:
        raise SystemExit(
            "\nFaltan las credenciales de DataForSEO. Podes pasarlas de tres formas:\n"
            "  1) python3 aio_check.py --login TU_EMAIL --password TU_API_PASSWORD\n"
            "  2) export DATAFORSEO_LOGIN=... ; export DATAFORSEO_PASSWORD=...\n"
            "  3) creando el archivo credenciales.txt con una linea 'email:api_password'\n"
        )
    return login, password


def main():
    p = argparse.ArgumentParser(
        description="Marca si cada keyword dispara AI Overview en Google (via DataForSEO).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--login", help="Email de la cuenta de DataForSEO")
    p.add_argument("--password", help="API password de DataForSEO")
    p.add_argument("-n", "--cantidad", type=int, default=1666,
                   help="Keywords a procesar en esta corrida (default: 1666, el maximo con USD 1)")
    p.add_argument("--location", default="Argentina", help="Pais del SERP (default: Argentina)")
    p.add_argument("--language", default="es", help="Idioma del SERP (default: es)")
    p.add_argument("--device", default="mobile", choices=["mobile", "desktop"],
                   help="Dispositivo simulado (default: mobile)")
    p.add_argument("--load-async", action="store_true",
                   help="Carga tambien el contenido de los AI Overviews asincronicos. "
                        "Duplica el costo a USD 0.0012 por keyword; no hace falta para saber Si/No.")
    p.add_argument("--espera-max", type=int, default=90,
                   help="Minutos maximos de espera por la cola Standard (default: 90)")
    p.add_argument("--dry-run", action="store_true",
                   help="Muestra que keywords tocarian en esta corrida y el costo, sin gastar nada")
    p.add_argument("--sin-chequeo-saldo", action="store_true",
                   help="No aborta si el saldo parece insuficiente")
    args = p.parse_args()

    if args.cantidad < 1:
        raise SystemExit("--cantidad debe ser al menos 1.")

    if not os.path.exists(ARCHIVO_KEYWORDS):
        raise SystemExit("No encuentro el Excel de keywords en %s" % ARCHIVO_KEYWORDS)

    headers, filas = read_xlsx(ARCHIVO_KEYWORDS)
    log("Excel leido: %d keywords, columnas %s" % (len(filas), headers))

    orden = cargar_o_crear_orden(filas)
    hechos = cargar_resultados()
    corrida = siguiente_numero_corrida(hechos)

    pendientes_kw = []
    vistas = set()
    for idx in orden:
        kw = filas[idx][COL_KEYWORD].strip()
        clave = kw.lower()
        if not kw or clave in hechos or clave in vistas:
            continue
        vistas.add(clave)
        pendientes_kw.append(kw)

    log("Ya procesadas: %d | pendientes: %d" % (len(hechos), len(pendientes_kw)))
    if not pendientes_kw:
        log("No queda ninguna keyword pendiente. Ya esta todo el archivo procesado.")
        return

    lote = pendientes_kw[:args.cantidad]
    costo_unitario = COSTO_POR_KEYWORD_ASYNC if args.load_async else COSTO_POR_KEYWORD
    costo = len(lote) * costo_unitario

    cfg = {
        "location": args.location, "language": args.language,
        "device": args.device, "load_async": args.load_async,
    }

    print("\n" + "-" * 62)
    print("CORRIDA %02d" % corrida)
    print("-" * 62)
    print("Keywords en esta corrida : %d" % len(lote))
    print("Mercado                  : %s / %s / %s" % (args.location, args.language, args.device))
    print("Costo estimado           : USD %.4f  (%.4f por keyword)" % (costo, costo_unitario))
    print("Primeras 5               : %s" % ", ".join(lote[:5]))
    print("-" * 62 + "\n")

    if args.dry_run:
        log("Modo --dry-run: no se envio nada a DataForSEO.")
        return

    login, password = resolver_credenciales(args)
    cliente = DataForSEO(login, password)

    saldo = cliente.saldo()
    log("Cuenta %s | saldo USD %.4f" % (login, saldo))
    if saldo < costo and not args.sin_chequeo_saldo:
        maximo = int(saldo / costo_unitario)
        raise SystemExit(
            "\nEl saldo (USD %.4f) no alcanza para %d keywords (USD %.4f).\n"
            "Con este saldo te alcanza para %d keywords: volve a correr con  -n %d\n"
            % (saldo, len(lote), costo, maximo, maximo)
        )

    inicio = time.time()
    try:
        log("Enviando %d keywords a la cola Standard..." % len(lote))
        pendientes = postear_tareas(cliente, lote, cfg, corrida)
        if not pendientes:
            raise SystemExit("No se creo ninguna tarea. Revisa las credenciales y el saldo.")

        # Se guardan los ids por si hay que recuperar la corrida despues de un corte.
        os.makedirs(DIR_ESTADO, exist_ok=True)
        with open(os.path.join(DIR_ESTADO, "corrida_%02d_tareas.json" % corrida), "w",
                  encoding="utf-8") as fh:
            json.dump(pendientes, fh, ensure_ascii=False, indent=1)

        log("%d tareas en cola. Esperando resultados (hasta %d min)..."
            % (len(pendientes), args.espera_max))
        resueltos, no_resueltas = recolectar(cliente, pendientes, args.espera_max)
    except SaldoInsuficiente as e:
        raise SystemExit("\nSaldo insuficiente en DataForSEO: %s\n"
                         "Recarga la cuenta o corre con un -n mas chico.\n" % e)

    if not resueltos:
        raise SystemExit("No se recupero ningun resultado. El progreso guardado no cambio.")

    ahora = datetime.now(timezone.utc).isoformat(timespec="seconds")
    nuevos = []
    for kw, dato in resueltos.items():
        fila = {
            "keyword": kw, "ai_overview": dato["ai_overview"], "corrida": corrida,
            "fecha_utc": ahora, "location": args.location, "language": args.language,
            "device": args.device, "item_types": dato["item_types"], "task_id": dato["task_id"],
        }
        nuevos.append(fila)
        hechos[kw.strip().lower()] = fila
    guardar_resultados(nuevos)
    log("Progreso guardado en %s" % ARCHIVO_RESULTADOS)

    ruta_corrida, ruta_cons, filas_corrida, filas_cons = escribir_salidas(
        headers, filas, hechos, lote, corrida, cfg)

    resumen("RESULTADO DE LA CORRIDA %02d" % corrida, filas_corrida,
            COL_AIO, COL_VERTICAL, COL_TIPO)

    # Control de sanidad: en el mercado argentino se espera un porcentaje
    # distinto de cero. Un 0% redondo suele indicar un problema de configuracion
    # antes que una realidad del SERP.
    con_aio = sum(1 for f in filas_corrida if f[COL_AIO] == "Si")
    if len(filas_corrida) >= 20 and con_aio == 0:
        log("AVISO: ninguna keyword dio AI Overview. Antes de seguir gastando, "
            "revisa la columna item_types en estado/resultados.csv: si nunca "
            "aparece 'ai_overview', puede que el mercado/idioma esten mal o que "
            "haga falta correr con --load-async.")
    if len(filas_cons) > len(filas_corrida):
        resumen("ACUMULADO DE TODAS LAS CORRIDAS", filas_cons,
                COL_AIO, COL_VERTICAL, COL_TIPO)

    if no_resueltas:
        log("Quedaron %d tareas sin resolver dentro del limite de espera. "
            "Esas keywords siguen pendientes y entran en la proxima corrida."
            % len(no_resueltas))

    restantes = len(pendientes_kw) - len(resueltos)
    print("Excel de esta corrida : %s" % ruta_corrida)
    print("Excel consolidado     : %s" % ruta_cons)
    print("(cada uno tiene su espejo en .csv en la misma carpeta)")
    print("Keywords pendientes   : %d  (~%d corridas mas de %d)"
          % (restantes, -(-restantes // max(args.cantidad, 1)), args.cantidad))
    print("Tiempo total          : %.1f min\n" % ((time.time() - inicio) / 60))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrumpido. El progreso ya guardado se conserva.", file=sys.stderr)
        sys.exit(130)
