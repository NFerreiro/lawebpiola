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
| **`depth: 10`** | DataForSEO cobra por unidades de 10 resultados. Como el AI Overview aparece arriba de todo, con 10 alcanza y se paga 1 unidad en vez de 10. |
| **Sin `load_async_ai_overview`** | Ese parámetro duplica el costo a USD 0,0012 y solo sirve para traer el *texto* del AI Overview. Para saber Si/No no hace falta: `item_types` ya incluye `ai_overview` aunque el overview sea asincrónico. |
| Envío en lotes de 100 tareas | Menos requests, misma tarifa. |
| Deduplicación y progreso persistente | Ninguna keyword se paga dos veces, ni siquiera si cortás una corrida a la mitad. |

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
completo (desvío menor a 0,05 puntos porcentuales). Es decir, **el porcentaje de
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

### 2. Piloto de validación (recomendado antes de la primera tanda completa)

```bash
python3 aio_check.py --login TU_EMAIL --password TU_API_PASSWORD -n 20
```

Gasta USD 0,012 y confirma que las credenciales, el mercado y la detección
funcionan. Revisá la columna `item_types` en `estado/resultados.csv`: si en
ninguna fila aparece `ai_overview`, algo está mal configurado — no sigas.

### 3. Corrida completa (1.666 keywords ≈ USD 1)

```bash
python3 aio_check.py --login TU_EMAIL --password TU_API_PASSWORD -n 1666
```

### 4. Siguientes corridas, con la cuenta nueva

Mismo comando con las credenciales nuevas. El script detecta solo cuáles
keywords ya están hechas y toma las 1.666 siguientes:

```bash
python3 aio_check.py --login OTRO_EMAIL --password OTRA_API_PASSWORD -n 1666
```

> El piloto de 20 keywords consume saldo de la tanda. Si querés exactamente
> 1.666 por cuenta, hacé el piloto con la primera cuenta y después corré
> `-n 1646` con esa misma.

### Otras formas de pasar las credenciales

```bash
export DATAFORSEO_LOGIN="tu_email"
export DATAFORSEO_PASSWORD="tu_api_password"
python3 aio_check.py -n 1666
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
if "ai_overview" in tipos:
    return True
```

DataForSEO agrega `ai_overview` a `item_types` **siempre que lo detecta en el
SERP**, sin importar si es sincrónico (cacheado por Google) o asincrónico
(generado on-demand). En el caso asincrónico el *contenido* viene vacío salvo
que pagues `load_async_ai_overview`, pero la **presencia** ya está registrada,
que es justo lo que necesitamos. Como respaldo, el script también recorre
`items` buscando `type == "ai_overview"` por si la estructura cambia.
