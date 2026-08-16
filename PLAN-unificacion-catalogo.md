# Unificación del catálogo: 2 hojas maestras (TODO + Carmen)

> **ESTADO: EN PAUSA (2026-07-29).** No implementar nada todavía. Se retoma cuando el usuario
> termine la **Fase 0** (preparar manualmente las hojas TODO + Carmen) y dé el OK explícito.
> El código arranca recién ahí, fase por fase.
>
> _(Copia de trabajo dentro del repo. El original también vive en
> `C:\Users\VAlmao\.claude\plans\smooth-beaming-pearl.md` y se auto-restaura al retomar.)_

## Context
Hoy la "lista de ítems" está fragmentada en ~4 hojas y hay que mantener un repuesto nuevo en
varias a la vez: **DB_REPUESTOS** (define qué es pedible en el Portal), **STOCK_REPUESTOS**
(metadata: mínimo/categoría/REQUIERE_SN), **CATALOGO_DJI** (staging del importador) y la
**lista de precios** (`1DWjX`, pestañas TODO/ACCESORIOS/Lista_Repuestos). El stock real ya sale
de **Carmen** (`1-BH5m…`, pestaña STOCK) y varios lectores (SM `cargarStock`, WOS, Portal stock)
ya la usan como fuente. El usuario quiere mantener **una sola** lista a mano.

Objetivo: dejar **2 hojas maestras** — un **catálogo maestro** (pestaña `TODO`) con todos los
códigos + atributos, y **Carmen » STOCK** solo para stock (solo lectura, col C es fórmula).
Todo el resto se retira reciclando antes lo que sirva.

## Decisiones tomadas (usuario)
- Catálogo maestro = pestaña **`TODO`** de `1DWjX` "Listado de repuestos DJI Agras".
- Stock = **Carmen » STOCK** (solo lectura).
- Conservar como columna en TODO: **Descripción ES**, **Reemplazado por**, **Requiere SN**.
- **Descartar** (se quitan sus features): **Categoría ABC**, **Stock mínimo**, **Precio FOB**.
  → El dashboard de SM pierde: alertas por categoría, KPI "bajo mínimo" y "valor de depósito"
    (o "valor" pasa a estimarse sobre PVP). Se quitan/neutralizan en Fase 2.
- Se retiran como catálogo: **DB_REPUESTOS**, **STOCK_REPUESTOS** (deja de ser catálogo),
  **CATALOGO_DJI**, y la pestaña **ACCESORIOS** (sus ítems se integran a TODO).

## Estado real de las 2 hojas hoy
**TODO** (getSheets()[0] de 1DWjX; lo lee SM `cargarCatalogoBorrador`):
`A CÓDIGO Largo · B CÓDIGO Corto · C DESCRIPCIÓN(EN) · D MODELO · E Cant/equipo · F PRECIO PVP · G Nivel rep · H IMAGEN · I TIPO`

**Carmen » STOCK**:
`A PN(código) · B Descripción · C Stock Actual(fórmula) · D Modelo(grupo) · E Serie(=modelos compatibles) · F Inicial`
> Carmen PN mezcla formato corto (`BC.AG.PP000354`) y con sufijo (`BC.AG.SS000668.01`).

## Fase 0 — Preparación MANUAL de las hojas  ← lo hace el usuario, ANTES de tocar código
Layout final de **TODO** (agregar las columnas nuevas **a la derecha**, para NO mover los índices existentes 0–8):

| A | B | C | D | E | F | G | H | I | **J** | **K** | **L** |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Largo | Corto | Desc EN | Modelo | Cant/eq | PVP | Nivel | Imagen | Tipo | **Desc ES** | **Reemplazado por** | **Requiere SN** |

Pasos:
1. **Clave canónica = Código Largo (col A).** Distingue variantes `.01/.02`. El cruce con Carmen
   se hará por esta clave (el código tolerará largo→corto como fallback, pero apuntá a que
   coincidan). No usar Corto como clave: colapsaría variantes.
2. **Traer TODOS los códigos a TODO** (que sea la lista completa):
   - Códigos que hoy solo están en **DB_REPUESTOS** (pedibles en Portal) y no en TODO → agregar.
   - Ítems de la pestaña **ACCESORIOS** → agregar a TODO.
   - Códigos que están en **Carmen** con stock y no en TODO → agregar.
3. **Poblar las 3 columnas nuevas** reciclando:
   - **J Desc ES** ← DB_REPUESTOS col H (descripción_ES).
   - **K Reemplazado por** ← DB_REPUESTOS col I (reemplazado_por).
   - **L Requiere SN** ← STOCK_REPUESTOS col J (REQUIERE_SN): TRUE en los serializados.
4. **Reconciliar Carmen ↔ TODO**: cada PN de Carmen debe existir en TODO (por largo o corto).
   Listar los que no matcheen y corregir (agregar al catálogo o normalizar el código).
5. **Carmen**: no tocar col C (fórmula). Solo agregar filas de códigos faltantes si aplica.

**Cuando TODO + Carmen estén listas y confirmadas, arranca el código.**

## Fases de código (después de Fase 0 — una por una, con deploy/test entre cada una)
- **Fase 1 — Helper único + SCHEMA.** Un accesor central `_getCatalogo()` que lea `TODO` →
  `{ codigoLargo: { corto, descEN, descES, modelos, cant, precioPVP, reemplazadoPor, requiereSN, tipo } }`,
  con cache y normalización de clave (largo→corto fallback contra Carmen). Definir el SCHEMA de TODO
  en cada proyecto (SM/Portal/WOS). Reutiliza el patrón de `_getCarmenStockMap` (STOCK_MANAGER/Sm_Core.js).
- **Fase 2 — SM.** Repointar a TODO: `_descDeSku` y `getModelosDisponibles` (Sm_Core/Sm_Etiquetas),
  overlay de `cargarStock`, y la fuente de `REQUIERE_SN` (etiquetas) → col L de TODO.
  Dashboard: quitar/neutralizar KPIs de categoría, bajo-mínimo y valor-FOB (atributos descartados).
  Retirar `sincronizarCatalogoDJI` (importador CATALOGO_DJI→DB/STOCK) como vía de mantenimiento.
- **Fase 3 — Portal.** `obtenerIndiceRepuestosPortal` y `buscarRepuestoConStockPortal` (RS_Pedidos)
  + búsqueda de RS_Repuestos: leer **TODO** en vez de DB_REPUESTOS; ACCESORIOS ya integrada en TODO.
  Precio sigue saliendo de la lista de precios (col F PVP × factor de descuento). Stock sigue de Carmen.
  reemplazadoPor/descES ahora salen de TODO (col K/J).
- **Fase 4 — WOS.** Quitar el overlay de metadata de STOCK_REPUESTOS (min/cat/ubic) en
  `WOS_cargarStock`/Despacho_Código.js; modelos desde TODO (o Carmen col E). Lista+stock ya de Carmen.
- **Fase 5 — Limpieza.** Retirar del uso DB_REPUESTOS, STOCK_REPUESTOS (como catálogo) y CATALOGO_DJI.
  Documentar en Context.md que el catálogo se mantiene SOLO en TODO y el stock SOLO en Carmen.

## Constraints
- Carmen STOCK = solo lectura (col C fórmula). Nunca escribir.
- Cada archivo: bump `@version` + fila en VERSIONS.md. Sin cambios de `appsscript.json` → sin re-auth
  (verificar en cada fase; avisar si alguna requiere scope nuevo).
- No push/deploy (lo hace el usuario). Solo rama `main`.

## Verificación (por fase, tras deploy la corre el usuario)
- `node --check` de cada .js tocado.
- Fase 2: alta de un repuesto SOLO en Carmen+TODO → aparece en SM (grid, borrador, etiqueta con/sin SN).
- Fase 3: ese mismo repuesto es **pedible** por un reseller en el Portal (buscador lo encuentra,
  con precio y estado de stock correctos) sin haberlo cargado en DB_REPUESTOS.
- Fase 4: WOS lista el ítem con sus modelos, sin depender de STOCK_REPUESTOS.
- Regresión: pedidos/etiquetas/despacho existentes siguen funcionando.
