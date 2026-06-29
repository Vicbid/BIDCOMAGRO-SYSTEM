# BIDCOMAGRO-SYSTEM — Control de Versiones

Cada archivo tiene un comentario `// @version X.Y` (JS) o `<!-- @version X.Y -->` (HTML).
Regla: **incrementar la versión cada vez que se edita un archivo**.
- Corrección de bug → X.Y+1 (ej: 1.0 → 1.1)
- Feature nueva → X+1.0 (ej: 1.3 → 2.0)

---

## HUB_PRO
| Archivo | Versión | Notas |
|---------|---------|-------|
| HUB_Código.js | 2.1 | Flujo batería unificado: trigger reposición mueve de "Aprobado por DJI" a "Scrap Enviado (Evidencias)"; nuevos estados Caso Enviado y Bateria enviada a reseller en ESTADOS_NOTIFICAR_RESELLER |
| Index.html | 2.0 | EST_BAT reemplaza EST_RPB con 7 estados; detección batería sin restricción de circuito; esTerminal incluye Rechazado DJI y Sin respuesta · Cerrado |
| GUIA_FLUJOS.html | 1.3 | Flujo batería: tab renombrado, sección reescrita con nuevos 7 estados; aplica a Reseller Común y Propio |
| Env.js | 1.2 | Agrega SCHEMA.EQUIPOS (NOMBRE, TIPO, PREFIJO, MESES) |

## PORTAL_RESELLER
| Archivo | Versión | Notas |
|---------|---------|-------|
| Index.html | 1.1 | Fix enviarPedido() → generarPedidoRepuestosPortal |
| RS_Pedidos.js | 1.5 | Fórmula CANT_PEND col G incluye -Z (CANT_CANCEL) al crear filas | Fix PDF: total ocupa cols 6-7 merge (160px) en vez de col 7 sola (80px) |
| RS_OTs.js | 1.7 | Fix threading: replyHtml→replyAll en agregarComentario y aprobarPresupuestoPortal (replyHtml no existe en GmailThread, siempre fallback a sendEmail nuevo) |
| RS_OTs.js | 1.6 | Flujo batería unificado: detección esBateriaOT antes de ramas de flujo; 7 estados nuevos con paso/quePasa propios; esBateria en return usa variable en lugar de IIFE duplicado |
| Index.html | 1.3 | Stepper batería: 4 pasos (Registro→Caso DJI→Aprobado DJI→Finalizado); detección esBat aplica a cualquier circuito (no solo Reseller Propio); btn-repuestos oculto para baterías |
| RS_Auth.js | 1.1 | Perfil: agrega provincia (col 5) en obtener y actualizar |
| Env.js | 1.1 | SCHEMA.RESELLERS: agrega PROVINCIA: 5 |
| RS_Email.js | 1.0 | |
| RS_Main.js | 1.0 | |
| RS_OTs.js | 1.0 | |
| RS_Auth.js | 1.0 | |
| RS_Repuestos.js | 1.0 | |
| RS_Código.js | 1.0 | |
| RS_CuentaCorriente.js | 1.0 | |
| RS_Registro.js | 1.0 | |
| RS_RTV.js | 1.0 | |
| RS_Campana.js | 1.0 | |
| RS_Avisos.js | 1.0 | |
| Env.js | 1.0 | |

## WOS
| Archivo | Versión | Notas |
|---------|---------|-------|
| Despacho_Index.html | 3.7 | Modal despacho: selector de ubicacion WMS por item (recomendacion = mayor stock); pick ticket muestra ubicacion sugerida en negrita |
| WOS_GmailFlow.js | 1.8 | WOS_despacharCompleto: escribe ubicacion en col H de Entregados y resta cantidad de UBICACIONES |
| Despacho_Código.js | 2.9 | WOS_cargarUbicacionesPedido(skus): carga ubicaciones WMS de multiples SKUs en una lectura, ordenadas desc por cantidad |
| Despacho_Código.js | 3.0 | CANT_CANCEL col Z: demanda perdida usa CANT_CANCEL con fallback a CANCELADO viejo | Optimizaciones: elimina funciones duplicadas (_lookupResellerEmail, _emailHead/_emailFoot), unifica cache key CARMEN, enCaminoMap en WOS_cargarStock llama a WOS_getEnCaminoMap(), alinea lógica de compras activas en reporteBackorder, batch writes en WOS_revertirAPreparado, elimina logs debug |
| Despacho_Index.html | 3.0 | Pick ticket: columna Estado + fix multi-página al imprimir (position:fixed→static) |
| Despacho_Index.html | 3.6 | Pick ticket: badge de cada ítem usa it.estado real (mismo que la tarjeta de pedido); print-color-adjust para colores en impresión |
| Despacho_Index.html | 3.3 | Pick tracker: badge ciclable reemplazado por select directo; _buildPickCajaOpts/_onPickCajaSel/_actualizarPickCajaSelects en lugar de _cyclePickCaja |
| WOS_GmailFlow.js | 1.7 | Nota de entrega agrupa items por caja con fila separadora Caja N + tracking cuando hay multi-bulto; Fix 1.6: fila TOTAL en col 6 |
| Despacho_Código.js | 2.8 | WOS_cargarStock: lee ubicaciones desde Carmen UBICACIONES tab (multi-ubicacion WMS); fallback a STOCK_REPUESTOS para items sin mapear |
| Despacho_Env.js | 1.1 | Agrega CARMEN_UBICACIONES_TAB = 'UBICACIONES' |
| Despacho_Env.js | 1.0 | |

## STOCK_MANAGER
| Archivo | Versión | Notas |
|---------|---------|-------|
| SM_Index.html | 2.0 | Borrador: buscador usa catálogo completo DJI (1372 repuestos del sheet externo) en lugar de solo STOCK_REPUESTOS; muestra modelo y precio en resultados |
| SM_Index.html | 2.9 | WMS panel Sin mapear: tercer modo con lista de items con stock pero sin ubicacion, badge con conteo, click abre modo Por item pre-cargado |
| Sm_Código.js | 3.0 | cargarSinMapear(): cruza Carmen STOCK vs UBICACIONES, devuelve hasta 200 items sin ubicacion ordenados desc por stock |
| SM_Index.html | 2.8 | Ingreso DJI: campo Destino WMS en modal detalle-cas y modal estado-cas; ubicacion se pasa a cada item al recibir |
| Sm_Código.js | 2.9 | _escribirEnRecibidos: escribe ubicacion en col I de Recibidos y suma cantidad a UBICACIONES tally |
| SM_Index.html | 2.7 | WMS modo sector: toggle Por item / Por sector; conteo rapido por rack con autocomplete y lista editable |
| Sm_Código.js | 2.8 | cargarUbicacionesSector(): devuelve todos los items de una ubicacion especifica |
| SM_Index.html | 2.6 | WMS: advertencia B cuando total asignado supera stock Carmen; barra de estado Asignado/Sin asignar/Todo asignado |
| Sm_Código.js | 2.7 | UBICACIONES sin formula: guardarUbicacionInicial y _registrarMovimientoCarmen usan numero directo en col C; STOCK col C de Carmen nunca se toca |
| SM_Index.html | 2.5 | WMS tab: tabla Ubicacion/Cantidad sin columna Inicial; edicion directa de cantidad |
| Sm_Código.js | 2.6 | WMS: cargarUbicacionesItem, guardarUbicacionInicial, eliminarUbicacion |
| SM_Index.html | 2.4 | Modal ubicaciones SKU: reemplaza cargarBinsSKU (TABLA_POSICIONES) por cargarUbicacionesItem (Carmen WMS) |
| SM_Index.html | 2.3 | Tab WMS: busqueda de repuesto, asignacion de ubicaciones con cantidad inicial, edicion inline, eliminacion |
| Sm_Código.js | 2.5 | WMS Opcion A: _registrarMovimientoCarmen() appendea a Entregados/Recibidos segun diff; asegura fila con formula SUMIFS en UBICACIONES; ajustarInventario ya no llama _actualizarCarmenStock (Stock Actual es formula en Carmen) |
| Env.js | 1.3 | Agrega CARMEN_ENTREGADOS_TAB y CARMEN_RECIBIDOS_TAB |
| Sm_Código.js | 2.4 | Multi-ubicación WMS en Carmen: _getCarmenUbicMap() + _actualizarCarmenUbicacion() leen/escriben tab UBICACIONES de Carmen; cargarStock y ajustarInventario redirigidos a Carmen |
| Env.js | 1.2 | Agrega CARMEN_UBICACIONES_TAB = 'UBICACIONES' |
| Sm_Código.js | 2.3 | cargarCatalogoBorrador(): agrega stockActual cruzando con Carmen (_getCarmenStockMap), sin precio |
| SM_Index.html | 2.2 | Pantallas de carga: spinner global con label dinámico, filas skeleton en stock y movimientos, texto de carga en despachos/compras/ventas, placeholder del borrador indica estado del catálogo |
| Sm_Código.js | 2.2 | cargarStock(): fuente primaria cambiada a Carmen STOCK (5428 ítems) en lugar de STOCK_REPUESTOS; STOCK_REPUESTOS sigue proveyendo metadata (mínimo, categoría, ubicación) |
| Sm_Código.js | 2.1 | cargarCatalogoBorrador(): lee sheet externo CATALOGO_REPUESTOS_ID, devuelve codigo/codigoCorto/descripcion/modelo/precio |
| Env.js | 1.1 | Agrega CATALOGO_REPUESTOS_ID con ID del sheet externo de catálogo DJI |
| SM_Index.html | 1.8 | Borrador de pedido: buscador interactivo con dropdown (código+descripción+stock), navegación por teclado ↑↓/Enter, igual al portal reseller |
| Sm_Código.js | 1.5 | cruzarComprasExternas: incluye items (cols D/E/F) en cada entrada nuevas |
| Env.js | 1.0 | |

## LAUNCHER
| Archivo | Versión | Notas |
|---------|---------|-------|
| Launcher_Código.js | 1.1 | Default emailFact incluye lucia.c@bidcom.com.ar |
| Launcher_Index.html | 1.2 | HUB PRO card: boton Guia de flujos; Tools WOS: trigger backorder; campo emailFact acepta multiples |

## Raíz
| Archivo | Versión | Notas |
|---------|---------|-------|
| SharedLogger.js | 1.0 | |
