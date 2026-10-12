# Control de compras por tickets

Sistema para registrar las compras a partir de los tickets que vas escaneando y
tener un resumen claro de en qué se va el dinero.

## Archivo principal

**control-compras.xlsx** — el libro de control, con 5 hojas:

| Hoja | Para qué sirve |
|---|---|
| **Resumen** | Estado del ejercicio, gasto por mes, por categoría y por forma de pago, más dos gráficos (gasto por mes y reparto por categoría). Se calcula solo. |
| **Compras** | Una fila por **producto** del ticket. Aquí está el detalle. |
| **Tickets** | Una fila por **ticket**. Compara el total del papel con lo registrado y avisa si no cuadra o si está repetido. |
| **Presupuesto** | Presupuesto mensual por categoría y % ejecutado del mes en curso. |
| **Listas** | Categorías, formas de pago, unidades y monedas. Editable: los desplegables salen de aquí. |

## Carpetas y programas

```
Control-Compras/
├─ control-compras.xlsx          <- el registro
├─ servidor.py                   <- web para el móvil (red local)
├─ LEEME.md                      <- este archivo
├─ web/                          <- la app del móvil (HTML + JS + CSS)
├─ copias/                       <- copias de seguridad del libro (las 20 últimas)
├─ ejemplo/                      <- libro de EJEMPLO con 10 tickets ficticios (no es el tuyo)
├─ tickets/
│  ├─ entrada/                   <- deja aquí los escaneos nuevos
│  ├─ archivados/                <- los ya registrados, renombrados
│  └─ manifiesto.csv             <- índice con la huella SHA-256 de cada escaneo
├─ lecturas/                     <- imágenes legibles sacadas de los PDF
└─ scripts/
   ├─ comun.py                   <- constantes, rutas y utilidades compartidas
   ├─ crear_libro.py             <- recrea el libro VACÍO (¡borra datos!)
   ├─ editar.py                  <- corrige o anula un ticket ya registrado
   ├─ leer_escaneo.py            <- pasa un PDF escaneado a imágenes legibles
   ├─ registrar_ticket.py        <- registra un ticket (o un lote) desde un JSON
   ├─ verificar.py               <- audita el registro y el archivo de escaneos
   └─ prueba_logica.py           <- prueba automática de todo el circuito
```

## Web (móvil y ordenador)

Podes cargar compras y ver el resumen desde el **teléfono o el ordenador**, sin
instalar nada. La página es responsive: en pantalla ancha se ve como un panel de
escritorio (columnas lado a lado) y en el móvil como una app.

```powershell
python servidor.py
```

Te imprime dos direcciones: una para este PC (`http://localhost:8080`) y otra
para el teléfono (`http://192.168.x.x:8080`). En el móvil, con la **misma WiFi**
que este ordenador, abrí esa dirección en el navegador.

Para que parezca una app: en Chrome/Android tocá el menú ⋮ → *Añadir a la pantalla
de inicio*; en Safari/iPhone tocá el botón *Compartir* → *Añadir a pantalla de inicio*.

La app permite: ver el resumen (mes, ejercicio, por categoría), ver la lista de
compras, y cargar una compra nueva con **foto del ticket** (abre la cámara
directamente), total, categoría, forma de pago y detalle opcional. Todo pasa por
la misma validación, control de duplicados, archivado y copias de seguridad que
el registro desde la terminal.

- Solo funciona en la **red local** (no se expone a internet).
- Si el teléfono no puede entrar, suele ser el Firewall de Windows: permití el
  acceso de Python en redes privadas, o ejecutá como administrador:
  `New-NetFirewallRule -DisplayName "Control Compras" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8080`.
- Para cambiar el puerto: `python servidor.py --puerto 8090`.

## Cómo se registra un ticket

1. Escaneas el ticket (foto, JPG o PDF) y me lo pasas, o lo dejas en `tickets/entrada/`.
2. Si es un **PDF**, lo convierto antes a imágenes con `leer_escaneo.py`: una
   imagen por página, con el motor que ya trae el entorno (no hay que instalar
   nada). Si es una foto enorme, la reduce para poder leerla sin perder datos.
3. Lo leo y, si algo está dudoso (manuscrito, importe ilegible, descuentos
   raros), **te lo pregunto antes** de darlo por bueno.
4. Registro cabecera + líneas con `registrar_ticket.py`, que de una sola pasada:
   - valida los datos (fechas, importes, unidades, forma de pago),
   - **rechaza duplicados**,
   - asigna el ID correlativo (`TK-0001`, `TK-0002`…),
   - escribe el ticket y sus líneas,
   - archiva el escaneo como `AAAA-MM-DD_comercio_total.ext`,
   - informa de si las líneas cuadran con el total del papel,
   - avisa si el **IVA impreso** en el ticket no cuadra con el calculado.
5. El libro queda listo: verde si cuadra, rojo si hay que revisar.
6. Si me pasas **varios tickets de una vez**, los registro **en lote**: es
   **todo o nada**, si uno falla no se escribe ninguno y te digo cuál falla.

### Formato del archivo de entrada (JSON)

```json
{
  "fecha": "2026-03-15",
  "comercio": "Supermercado Central",
  "numero": "A-1024",
  "forma_pago": "Tarjeta debito",
  "moneda": "ARS",
  "total": 8.00,
  "iva_global": 21,
  "iva_importe": 1.39,
  "archivo": "tickets/entrada/foto.jpg",
  "notas": "compra semanal",
  "lineas": [
    {"descripcion": "Leche entera 1 L", "cantidad": 2, "unidad": "ud",
     "precio": 1.20, "descuento": 0, "iva": 10, "categoria": "Supermercado"},
    {"descripcion": "Descuento general", "cantidad": 1, "unidad": "ud",
     "precio": -1.00, "categoria": "Supermercado"}
  ]
}
```

- `id` es opcional: si falta se asigna el siguiente.
- `numero`: número de ticket o factura. Si el papel no lo trae, se pone `-`.
- `iva_global` es el % de IVA que se aplica a las líneas que no traigan el suyo.
- `iva_importe` es el IVA que imprime el ticket. Si lo pongo y todas las líneas
  comparten tipo, el sistema comprueba que cuadre y avisa si no.
- `descuento` es un importe, no un porcentaje.
- Un `precio` **negativo** es una línea de descuento (los tickets que imprimen
  el descuento en un renglón aparte). Lleva cantidad 1 y descuento 0.

Varios tickets a la vez: el mismo archivo con una **lista** `[ {...}, {...} ]`.

```powershell
python scripts/registrar_ticket.py ticket.json
python scripts/registrar_ticket.py lote.json                 # lista de tickets
python scripts/registrar_ticket.py lote.json --seco          # simula, no guarda
python scripts/registrar_ticket.py ticket.json --permitir-duplicado
```

## Controles automáticos

| Control | Dónde se ve | Qué significa |
|---|---|---|
| **Conciliación** | Tickets → *Estado* | `OK` si las líneas suman el total del papel (±0,02); `REVISAR` en rojo si no. |
| **Duplicados** | Tickets → *Duplicado* y Resumen → *Posibles duplicados* | `DUPLICADO` si hay otro ticket con la misma fecha, comercio, total y número. Se marca en las dos filas. |
| **Sin clasificar** | Resumen → *Sin clasificar* | Importe de líneas sin categoría. Debería ser 0. |
| **Panel de estado** | Resumen → filas 6-12 | Gasto del ejercicio, tickets, IVA incluido, tickets a revisar y duplicados. Los dos últimos en rojo si son > 0. |

El registro **no permite duplicados**, pero avisa en vez de esconder: dos compras
iguales el mismo día en el mismo comercio y por el mismo importe son sospechosas,
y si de verdad son distintas se registran con `--permitir-duplicado` y quedan
marcadas para que lo revises.

## Seguridad y auditoría

- **Copia de seguridad automática.** Antes de cada escritura, el libro anterior se
  guarda en `copias/` con fecha y hora. Se conservan las 20 últimas y se puede
  desactivar con `--sin-copia`. Si algo sale mal, siempre hay marcha atrás.
- **Manifiesto del archivo.** Al archivar cada escaneo se anota en
  `tickets/manifiesto.csv` su ruta, tamaño y **huella SHA-256**, para poder
  demostrar que el original sigue intacto.
- **Auditoría.** `verificar.py` revisa que los tickets cuadren con sus líneas, que
  no haya duplicados, líneas sin categoría, escaneos perdidos o alterados, y
  escaneos que ningún ticket reclame. Calcula el gasto **por su cuenta**
  (implementación independiente de las fórmulas del libro), así que sirve también
  para comprobar que el Excel no se ha desviado. Sale con código 1 si encuentra
  problemas de integridad; los avisos (diferencias, duplicados) no lo hacen fallar.

## Datos que extraigo de cada ticket

- Fecha (si no se lee, te la pregunto)
- Comercio o proveedor
- Nº de ticket o factura (si aparece)
- Forma de pago
- Total del ticket
- Cada línea: descripción, cantidad, unidad, precio unitario y descuento
- IVA: el % por línea y el importe de IVA incluido, calculado sobre el precio

## Parámetros que puedes cambiar

- **Resumen → B2**: ejercicio que quieres analizar.
- **Resumen → B3**: moneda. Configurado en **ARS** (los importes se muestran con `$`).
- **Listas**: hasta 24 categorías, 14 formas de pago y 10 unidades. Añade las
  que quieras: el Resumen y el Presupuesto las toman solas.
- **Presupuesto**: escribe tu presupuesto mensual por categoría (celdas amarillas).
- **Compras → Notas**: cualquier comentario libre.

## Reglas del sistema

- **Uso previsto:** gastos personales / del hogar, con el detalle **línea por línea**.
- **Moneda: pesos argentinos (ARS)**, formateados como `$ 1.234,56`.
- **Una sola moneda por ejercicio.** Si compras en otra divisa, indícalo en Notas
  o registra esos tickets aparte; las sumas no distinguen divisa.
- **El importe de línea** se calcula como `cantidad × precio − descuento`, así los
  totales son siempre coherentes con el detalle.
- **El total del papel manda.** Si no cuadra con el detalle, lo verás como
  `REVISAR` en vez de esconder la diferencia.
- **El IVA se considera incluido** en el precio del ticket (como se factura al
  consumidor final). Si un ticket no informa del %, la columna queda vacía y no
  se estima nada.

## Scripts

```powershell
# Prueba de extremo a extremo (no toca el registro real: trabaja sobre una copia)
python scripts/prueba_logica.py

# Auditar el registro y el archivo de escaneos
python scripts/verificar.py
python scripts/verificar.py --json

# Pasar un PDF escaneado a imágenes legibles (una por página)
python scripts/leer_escaneo.py tickets/entrada/factura.pdf

# Recrear el libro VACÍO. Se niega si ya hay tickets, salvo con --forzar
python scripts/crear_libro.py
```
