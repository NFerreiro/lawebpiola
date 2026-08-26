# AI Overviews en keywords de eCommerce (DataForSEO)

Marca, para cada una de las **13.065 keywords** de `data/Keywords_Ecommerce.xlsx`,
si Google devuelve un **AI Overview** (`Si`) o no (`No`).

Está pensado para correr **por tandas con saldo mínimo**: cada corrida gasta
solo lo que le indiques, guarda el progreso y arranca donde terminó la anterior,
aunque uses **otra cuenta de DataForSEO** en cada tanda.

## Cómo se abarata el costo

| Decisión | Por qué |
|---|---|
| SERP API Google Organic, **cola Standard** (`priority: 1`) | USD 0,0006 por keyword. La cola Priority cuesta el doble y el modo Live, más del triple. |
| **`depth: 10`** | DataForSEO factura por SERP de hasta 100 resultados, asi que `depth: 10` cuesta lo mismo que `depth: 100`: no lo usamos para ahorrar, sino porque el AI Overview esta siempre arriba de todo y pedir menos resultados devuelve respuestas mas chicas y rapidas. |
| **Sin `load_async_ai_overview`** | Ese parámetro duplica el costo a USD 0,0012 y solo sirve para traer el *texto* del AI Overview. Para saber Si/No no hace falta: `item_types` ya incluye `ai_overview` aunque el overview sea asincrónico. |
| Envío en lotes de 100 tareas | Menos requests, misma tarifa. |
| Deduplicación y progreso persistente | Ninguna keyword se paga dos veces, ni siquiera si cortás una corrida a la mitad. |
| Presupuesto medido con el costo **real** | El script no confía en el precio de lista: manda un primer lote de 20, lee el campo `cost` que devuelve la API y recién ahí decide cuántas keywords más entran en el dólar. Si la tarifa fuera más cara, la corrida se achica sola en vez de cortarse por saldo. |

**Costo total del archivo completo: USD 7,84** (13.065 × 0,0006), o sea
**8 corridas de 1.666 keywords**, no 10-12.

## Configuración de las consultas

- Mercado: **Argentina**, idioma **español** (`location_name: "Argentina"`, `language_code: "es"`)
- Dispositivo: **mobile** (Android)
- Motor: Google, resultados **Advanced** (los únicos que exponen features del SERP)

Todo se puede cambiar por línea de comandos, pero **no mezcles configuraciones
entre corridas**: el porcentaje final dejaría de ser comparable.

## Orden de las keywords

`data/orden_keywords.csv` fija un orden **estratificado** por `Vertical` y
`tipo_keyword`: cada tanda de 1.666 respeta la proporción real del archivo
completo (desvío máximo verificado: 0,08 puntos porcentuales, para cualquier tamaño de tanda). Es decir, **el porcentaje de
AI Overviews de la primera corrida ya es extrapolable a las 13.065 keywords**,
sin esperar a terminarlas todas.

El orden es determinístico (semilla fija) y queda versionado, así que las
corridas son reproducibles y nunca se pisan entre sí.

## Requisitos

Python 3.8 o superior. **Nada más** — no usa openpyxl, pandas ni requests.

## Uso

### 1. Probar sin gastar nada

```bash
python3 aio_check.py --dry-run
```

Muestra qué keywords entran en la próxima corrida y cuánto costaría.

### 2. Corrida completa (gastá el dólar entero, sin pasarte)

```bash
python3 aio_check.py --login TU_EMAIL --password TU_API_PASSWORD
```

Con los valores por defecto (`--presupuesto 1.0`) el script:

1. lee el saldo de la cuenta,
2. manda un lote de calibración de 20 keywords y **mide el costo real** que
   cobró DataForSEO,
3. con ese número calcula cuántas keywords más entran en el dólar y las manda,
4. corta los envíos apenas el gasto llega al tope.

Si la tarifa es la esperada (USD 0,0006), entran **1.666 keywords**. Si fuera
más cara, la corrida se achica sola: nunca se cuelga por saldo insuficiente a
mitad de camino.

Para gastar menos en la primera prueba:

```bash
python3 aio_check.py --login TU_EMAIL --password TU_API_PASSWORD --presupuesto 0.05
```

Revisá la columna `item_types` en `estado/resultados.csv`: si en ninguna fila
aparece `ai_overview`, algo está mal configurado — no sigas gastando.

### 3. Siguientes corridas, con la cuenta nueva

Mismo comando con las credenciales nuevas. El script detecta solo cuáles
keywords ya están hechas y sigue por las que faltan:

```bash
python3 aio_check.py --login OTRO_EMAIL --password OTRA_API_PASSWORD
```

Lo que gastes en pruebas sale del mismo dólar, pero no se pierde: esas
keywords quedan guardadas y no se vuelven a pagar en la corrida siguiente.

### Otras formas de pasar las credenciales

```bash
export DATAFORSEO_LOGIN="tu_email"
export DATAFORSEO_PASSWORD="tu_api_password"
python3 aio_check.py
```

O un archivo `credenciales.txt` (ya está en `.gitignore`) con una línea:

```
tu_email:tu_api_password
```

## Qué genera cada corrida

```
salidas/
  AIO_corrida_01.xlsx     <- las 1.666 keywords de esta tanda, con las 5 columnas originales
  AIO_corrida_01.csv         y la última ("¿Tiene AI overviews?") completada con Si / No
  AIO_consolidado.xlsx    <- todo lo procesado hasta ahora, + corrida, fecha, mercado, dispositivo
  AIO_consolidado.csv

estado/
  resultados.csv          <- progreso acumulado (incluye item_types crudo y task_id)
  corrida_01_tareas.json  <- ids de las tareas, por si hay que recuperar una corrida cortada
```

**No borres `estado/resultados.csv`**: es lo que evita volver a pagar por
keywords ya consultadas.

Al terminar, la consola imprime el porcentaje de AI Overviews de la corrida y el
acumulado, abierto por vertical y por tipo de keyword.

## Opciones

| Flag | Default | Para qué |
|---|---|---|
| `-n`, `--cantidad` | `1666` | Keywords de esta corrida |
| `--location` | `Argentina` | País del SERP |
| `--language` | `es` | Idioma del SERP |
| `--device` | `mobile` | `mobile` o `desktop` |
| `--load-async` | apagado | Fuerza la carga de AI Overviews asincrónicos. **Duplica el costo**; no hace falta para saber Si/No |
| `--presupuesto` | `1.0` | Tope de gasto en USD de esta corrida, medido con el costo real que devuelve la API |
| `--espera-max` | `90` | Minutos máximos de espera por la cola Standard |
| `--dry-run` | apagado | Simula la corrida sin gastar |
| `--sin-chequeo-saldo` | apagado | No aborta si el saldo parece insuficiente |

## Comportamiento ante problemas

- **Saldo insuficiente**: aborta *antes* de enviar nada y te dice con qué `-n` sí alcanza.
- **Credenciales inválidas**: corta con un mensaje claro (HTTP 401).
- **Errores de red, 429 o 5xx**: reintenta con espera creciente (2s, 4s, 8s, 16s).
- **Tareas que no resuelven a tiempo**: quedan pendientes y entran en la corrida siguiente.
- **Corte a mitad de camino**: lo ya resuelto queda guardado; nada se paga dos veces.
- **0% de AI Overviews en una tanda**: imprime un aviso, porque casi siempre es un error de configuración y no un dato real.

## Cómo se detecta el AI Overview

```python
tipos = resultado.get("item_types") or []
if "ai_overview" in tipos:                   # igualdad exacta, no substring
    return True
for item in resultado.get("items") or []:    # solo primer nivel
    if item.get("type") == "ai_overview":
        return True
```

DataForSEO agrega `ai_overview` a `item_types` **siempre que lo detecta en el
SERP**, sin importar si es sincrónico (cacheado por Google) o asincrónico
(generado on-demand). En el caso asincrónico el *contenido* viene vacío salvo
que pagues `load_async_ai_overview`, pero la **presencia** ya está registrada,
que es justo lo que necesitamos.

### El falso positivo que hay que evitar

Dentro de `people_also_ask`, cada pregunta puede traer un
`people_also_ask_ai_overview_expanded_element`. Eso es el mini-resumen que
Google arma **al desplegar una pregunta del PAA**, no un AI Overview del SERP,
y contarlo inflaría el porcentaje final.

Por eso la búsqueda es por **igualdad exacta** contra `"ai_overview"` y **solo
sobre items de primer nivel**: nunca por substring ni recursiva.

### Validación contra SERPs reales

Verificado el 26/08/2026 contra Google Argentina (es, mobile):

| Keyword | Qué devolvió el SERP | Resultado |
|---|---|---|
| `como elegir una freidora de aire` | item `ai_overview` de primer nivel, con 6 bloques de contenido y 25 referencias | **Si** |
| `freidoras` | ningún `ai_overview` de primer nivel, pero **4** `people_also_ask_ai_overview_expanded_element` anidados en el PAA | **No** |

Dos conclusiones prácticas:

1. El AI Overview llegó **con contenido completo sin pagar
   `load_async_ai_overview`**, así que ese extra no hace falta.
2. Sin el filtro de primer nivel, `freidoras` habría dado un falso "Si".

### Una alternativa más barata que se descartó

DataForSEO Labs *Keyword Overview* acepta 700 keywords por request y saldría
unas 5 veces más barato. Se probó y **no sirve**: no devuelve
`serp_info.serp_item_types` (o sea, ningún dato de AI Overview) y además su
base tiene huecos — de 8 keywords de prueba, omitió 1 por falta de cobertura.
Encima serían SERPs cacheados, no consultas en vivo. La SERP API es el camino
correcto.
