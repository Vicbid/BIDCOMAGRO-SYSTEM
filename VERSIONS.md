# BIDCOMAGRO-SYSTEM — Control de Versiones

Cada archivo tiene un comentario `// @version X.Y` (JS) o `<!-- @version X.Y -->` (HTML).
Regla: **incrementar la versión cada vez que se edita un archivo**.
- Corrección de bug → X.Y+1 (ej: 1.0 → 1.1)
- Feature nueva → X+1.0 (ej: 1.3 → 2.0)

---

## HUB_PRO
| Archivo | Versión | Notas |
|---------|---------|-------|
| Index.html | 2.2 | Botón "Pedir Repuestos": visible para IW y OOW (antes solo OOW); condición cambiada a !esTerminal |
| HUB_Código.js | 2.6 | HUB_generarPedidoRepuestos: 6 fixes — (1) OBS ya no duplica garantia; (2) permite múltiples pedidos por OT saltando SKUs ya pedidos; (3) MASTER abierto una sola vez; (4) bodyHtml diferenciado primer pedido vs adicional; (5) backfill threadId en filas preexistentes; (6) registrarEmailLog solo en primer pedido |
| Index.html | 2.5 | Botón "Piezas Recibidas": condición ampliada a cualquier circuito que contenga "Reseller" (cubre Reseller y Reseller Propio) y estado no terminal |
| Index.html | 2.4 | Botón "Piezas Recibidas" (naranja) visible en Reseller Estándar con estado "Informe de reparacion"; abre modal con textarea libre; al confirmar escribe entrada timestampeada en Notas Internas y avanza estado a "Piezas dañadas recibidas" via guardar() |
| Index.html | 2.3 | Modal Pedir Repuestos: número preview corregido OT-X (era REP-OT-X); checkbox DJI basado en estado "Aprobado"; botón btn-rep-wos muestra "WOS: OT-X" en verde tras éxito |
| HUB_Código.js | 2.5 | HUB_generarPedidoRepuestos: usa _enviarConHilo para continuar el hilo Gmail existente de la OT en lugar de crear thread nuevo; WOS responde en el mismo hilo al despachar |
| HUB_Código.js | 2.4 | HUB_generarPedidoRepuestos: lookup precio desde Lista_Repuestos (×0.60 precio reseller) al crear filas WOS; fallback a 0 si SKU no está en catálogo |
| HUB_Código.js | 2.3 | HUB_generarPedidoRepuestos: lee data.envio y lo escribe en col L; OBS incluye "Taller · entrega técnico + nombre" cuando circuito=Taller |
| Index.html | 2.1 | Modal "Pedir Repuestos": selector de modalidad de entrega — Taller muestra badge fijo "Entrega directa al técnico", Reseller muestra radio Envío/Retiro; envio y circuito se pasan al backend |
| HUB_Código.js | 2.2 | HUB_generarPedidoRepuestos: 26 cols (COL schema igual a Pedidos_resellers); formula =E-F-Z en col G; hilo Gmail ancla via draft.send() guardado en col R; lookup email reseller en MASTER; estado inicial Confirmado |
| HUB_Código.js | 2.1 | Flujo batería unificado: trigger reposición mueve de "Aprobado por DJI" a "Scrap Enviado (Evidencias)"; nuevos estados Caso Enviado y Bateria enviada a reseller en ESTADOS_NOTIFICAR_RESELLER |
| Index.html | 2.0 | EST_BAT reemplaza EST_RPB con 7 estados; detección batería sin restricción de circuito; esTerminal incluye Rechazado DJI y Sin respuesta · Cerrado |
| GUIA_FLUJOS.html | 1.3 | Flujo batería: tab renombrado, sección reescrita con nuevos 7 estados; aplica a Reseller Común y Propio |
| Env.js | 1.2 | Agrega SCHEMA.EQUIPOS (NOMBRE, TIPO, PREFIJO, MESES) |

## PORTAL_RESELLER
| Archivo | Versión | Notas |
|---------|---------|-------|
| Index.html | 1.1 | Fix enviarPedido() → generarPedidoRepuestosPortal |
| Index.html | 1.4 | Registro: 4 campos de dirección (Calle, CP, Localidad, Provincia) — se muestran solo si están vacíos en el sheet; Calle es requerida si visible; sección "Dirección fiscal" con separador; CP+Localidad en row flex |
| RS_Registro.js | 1.1 | REG_obtenerCamposVacios: expone campos direccion/cp/localidad/provincia; REG_solicitarAcceso: acepta y persiste los 4 campos en SOLICITUDES_RESELLER (cols J-M); _REG_procesarDecision: los escribe en hoja Resellers al aprobar; email admin muestra dirección consolidada |
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
| Despacho_Código.js | 3.1 | WOS despacha todo: WOS_cargarPedidos fusiona Pedidos_resellers + Pedidos_OTs; _procesarFilasPedidos extrae loop; todas las funciones per-numero (cancelar, cambiarEstado, reactivar, revertir, obs, etiqueta) usan _getHojaPorNumero; WOS_buscarBackorderPorSKU + WOS_recibirMercaderia + WOS_reporteBackorder + WOS_checkCambios + WOS_getResumenEnvios escanean ambas hojas; WOS_actualizarValidacion aplica a ambas; WOS_migrarPedidosOTs() función one-time que recrea Pedidos_OTs con 26 cols |
| WOS_GmailFlow.js | 2.3 | Fix: _esDJIAprobado y formaPago se declaraban DESPUÉS de usarlas en adminHtml (hoisting var → undefined); movidas al bloque previo a tituloAdmin |
| WOS_GmailFlow.js | 2.2 | WOS_despacharCompleto: email admin muestra "Sin costo — DJI aprobado" (verde) o "A 30 días" según OBS del pedido OT; PR mantiene campo PAGO existente |
| WOS_GmailFlow.js | 2.1 | WOS_despacharCompleto: al despachar OT-, actualiza E: por SKU en col Q (REPUESTOS) del HUB PRO + estado "Repuestos enviados" solo si despacho completo (sin backorder) |
| WOS_GmailFlow.js | 1.9 | _wosLeerPedido usa _getHojaPorNumero(numero); WOS_recuperarThreadIds escanea ambas hojas con hoja-por-numero; WOS_procesarRespuestaManual usa _getHojaPorNumero |
| Despacho_Env.js | 1.2 | HOJA_PEDIDOS_OT, _getHojaPedidosOT(), _esNumeroOT(), _getHojaPorNumero() — resolver de hoja por prefijo OT- vs resellers |
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
| Despacho_Index.html | 3.10 | Fix Lista Compras: aComprar = max(0, bo - enCamino - stockActual); stk se calculaba después de aComprar y nunca se descontaba — SKUs con stock suficiente mostraban "a comprar" en lugar de "Cubierto" |
| Despacho_Index.html | 3.9 | Batch despacho: checkbox incluye Backorder; en abrirBatchDespacho items de BO renderizan con value=0 + hint ámbar + borde ámbar; badge BACKORDER en separador de sección; ítems cancelados/ya-enviados excluidos del batch |
| Despacho_Index.html | 3.8 | Badges WMS en tabla de ítems: columna "Ubicación WMS" con wosPintarBadgesUbicacion(); CSS .wms-bin picking/reserva (tokens nativos); snapshot cargado post-render via _cargarWmsSnapshot() + _actualizarBadgesWms() sin bloquear UI |
| WOS_StockSnapshot.js | 1.1 | WOS_obtenerStockSnapshot() función pública para google.script.run; devuelve stockMap serializable al frontend |
| WOS_StockSnapshot.js | 1.0 | Módulo read-only WMS: _wosCargarStockSnapshot() triple in-memory mapping (Carmen STOCK + UBICACIONES + MASTER TABLA_POSICIONES); sort PICKING-first + alfanumérico natural; _wosStockBySku() helper de consulta |
| Despacho_Env.js | 1.0 | |

## STOCK_MANAGER
| Archivo | Versión | Notas |
|---------|---------|-------|
| SM_Index.html | 5.3 | Layout pasa a 3 niveles ESTANTE-PAÑO-ALTURA (ej "1-A-2"); editor reintroduce acordeón (estante=outer, paño=inner con stepper de alturas); mapa: rows=estantes, cols=paños únicos, celda agrega alturas con dots y total de stock; detalle de celda lista cada bin ESTANTE-PAÑO-ALTURA; _getAllBins genera "ESTANTE-PAÑO-ALTURA" |
| Sm_Código.js | 4.6 | Layout 3 niveles: SM_cargarLayout devuelve {estantes:[{estante,orden,panos:[{pano,orden,alturas}]}]}; SM_guardarLayout(estantes) escribe esquema [ESTANTE,ORDEN_ESTANTE,PAÑO,ORDEN_PAÑO,NUM_ALTURAS] celda por celda con setNumberFormat('@') en cols A y C (evita auto-fecha) |
| SM_Index.html | 5.2 | Fix bug raíz "todos los bins vacíos": _mapaConstruirBinStock leía item.bins (TABLA_POSICIONES, hoja legacy nunca escrita por el WMS actual) en vez de item.ubicaciones (tab UBICACIONES de Carmen, fuente real de Recibir/Contar/Mover); _wmsTabAbierto ahora llama _mapaConstruirBinStock apenas carga stockData (no solo al entrar a Mapa); todas las búsquedas de binKey uniformadas a .toUpperCase() |
| SM_Index.html | 5.1 | Layout sin pasillo: estructura simplificada a ESTANTE-NIVEL (ej "1-3"); editor reemplaza acordeón de pasillos por lista plana de estantes; mapa renderiza rows=estantes, cols=niveles; detalle de celda muestra bin individual; _getAllBins genera "ESTANTE-NIVEL" |
| Sm_Código.js | 4.5 | Fix fecha en layout: SM_guardarLayout fuerza col A a formato texto (@) + escribe celda por celda con setNumberFormat('@') — igual que fix UBICACIONES; SM_cargarLayout saltea filas donde el valor ESTANTE es Date (corruptas) |
| Sm_Código.js | 4.4 | Layout sin pasillo: SM_cargarLayout devuelve {estantes:[{estante,orden,niveles}]}; SM_guardarLayout(estantes) escribe esquema plano; asegurarHojas() actualiza headers LAYOUT_ALMACEN a [ESTANTE,ORDEN_ESTANTE,NUM_NIVELES] |
| SM_Index.html | 5.0 | WMS rediseñado: 3 tareas grandes (Recibir/Contar/Mover) reemplazan los modos abstractos "Por ítem"/"Por sector"; bin picker con autocomplete del layout en todos los campos de ubicación; RECIBIR con qty stepper y guardar en un click; CONTAR muestra ítems del bin con +/− inline y "Guardar conteo completo"; MOVER: buscar repuesto→chips de ubicaciones→bin destino→confirmar (SM_moverStock); layout se pre-carga en background al abrir WMS |
| Sm_Código.js | 4.3 | Fix fecha: guardarUbicacionInicial / guardarConteoUbicacion / SM_moverStock usan setNumberFormat('@') en col B de UBICACIONES — evita que Sheets auto-convierta "1-1" a fecha; SM_repararFormatoUbicaciones() para reparar filas ya corrompidas |
| Sm_Código.js | 4.2 | SM_moverStock(sku,origen,destino,cant): mueve unidades entre bins de UBICACIONES de Carmen atómicamente; guardarConteoUbicacion(ubicacion,items): reescribe el conteo completo de un bin (update+delete+insert en un solo paso) |
| SM_Index.html | 4.0 | Mapa aéreo del almacén: 4to botón "Mapa" en WMS toggle; panel wms-panel-mapa con grid por pasillo/estante; celdas por tipo (picking/sobrestock/cuarentena/mixto/vacío) con dots de niveles; click en celda abre detalle por nivel; modal Editor de Layout para definir pasillos, estantes y niveles de forma visual; SM_cargarLayout / SM_guardarLayout backend |
| Sm_Código.js | 4.1 | SM_cargarLayout(): lee LAYOUT_ALMACEN, arma estructura {pasillos:[{pasillo,estantes:[{estante,niveles}]}]} ordenada; SM_guardarLayout(pasillos): reescribe hoja LAYOUT_ALMACEN completa e invalida caché |
| SM_Index.html | 3.5 | Panel "Compras en tránsito" → "En vuelo / en aduana"; KPI idem; filtro solo muestra esos 2 estados |
| Sm_Código.js | 4.0 | cargarDashboard casTransito: filtro cambiado a CAS_ACTIVOS=['En vuelo','En aduana'] — ya no incluye Comprado/Pagado/Forwarder HK |
| SM_Index.html | 3.4 | Top rotación: segundo panel "por valor" ($); _renderTopRotList helper unifica render de ambos; cargarTopRotacion maneja nuevo retorno {cantidad, valor} |
| Sm_Código.js | 3.9 | obtenerTopRotacion: retorna {cantidad, valor} — ambos top-5 en una sola llamada; valor = suma CANT_SOL × PRECIO por SKU |
| SM_Index.html | 3.3 | renderOTTracking: muestra cantPend por SKU con badge naranja/verde; header de cada OT siempre tiene _wosBadge (wos nunca null) |
| Sm_Código.js | 3.8 | obtenerOTsBloqueadasConCAS: fuente cambiada a Pedidos_OTs WOS (en lugar de HUB OT estado=Espera de repuestos); obtenerTopRotacion: lee Pedidos_resellers+Pedidos_OTs WOS (en lugar de MOVIMIENTOS_STOCK), cuenta CANT_SOL por SKU en últimos 7 días |
| SM_Index.html | 3.2 | Dashboard: KPI y panel renombrados a "Pedidos WOS pendientes" — incluye Pedidos_resellers + Pedidos_OTs |
| Sm_Código.js | 3.7 | cargarDashboard: solicPendientes fusiona Pedidos_resellers + Pedidos_OTs del WOS (ambos tipos de pedido tratan igual al stock); SOLICITUDES_DESPACHO deprecada removida |
| SM_Index.html | 3.1 | Dashboard: panel "Solicitudes pendientes HUB" reemplazado por "Pedidos OT pendientes (WOS)" — lee Pedidos_OTs del WOS (estado, cantPend, diasEspera); renderOTTracking muestra badge de estado WOS por OT (_wosBadge) |
| Sm_Código.js | 3.6 | cargarDashboard: solicPendientes lee Pedidos_OTs del WOS en lugar de SOLICITUDES_DESPACHO (deprecada); SOLICITUDES removida del preload de hojas |
| SM_Index.html | 2.0 | Borrador: buscador usa catálogo completo DJI (1372 repuestos del sheet externo) en lugar de solo STOCK_REPUESTOS; muestra modelo y precio en resultados |
| SM_Index.html | 2.9 | WMS panel Sin mapear: tercer modo con lista de items con stock pero sin ubicacion, badge con conteo, click abre modo Por item pre-cargado |
| SM_Index.html | 3.0 | Cruce externo: nueva seccion Diferencias de items vs sheet externo (tabla SKU/SM/Externo, boton Usar externo sincroniza COMPRAS_DETALLE) |
| Sm_Código.js | 3.5 | obtenerOTsBloqueadasConCAS: agrega campo wos con estado de despacho leido de Pedidos_OTs (WOS); null si aun no generado en WOS |
| Env.js | 1.4 | Agrega WOS_NOTAS_SS_ID para que SM pueda leer Pedidos_OTs |
| Sm_Código.js | 3.4 | _alertarBackordersPendientes: reescrita — cruza SKUs del CAS contra WOS Pedidos_resellers (estado Backorder); cobertura basada en stock actual de STOCK_REPUESTOS (no en items del CAS); email muestra stock disponible vs pendiente por pedido |
| Sm_Código.js | 3.3 | cruzarComprasExternas: compara items de CAS no-deposito contra sheet externo, devuelve diferencias; sincronizarItemsCAS(): reemplaza items en COMPRAS_DETALLE preservando cantidades recibidas |
| Sm_Código.js | 3.2 | Fix stock WOS: elimina llamadas a _actualizarCarmenStock (sobreescribía la fórmula de Carmen col C); agrega restaurarFormulasCarmenStock() para reparar filas con número plano |
| Sm_Código.js | 3.1 | Despachos eliminados: funciones cargarSolicitudesDespacho, limpiarSolicitudesDeOTsCerradas, generarReportePendientesPDF, procesarDespacho, _actualizarEnviadosEnRep, _notificarDespacho removidas; todo el despacho pasa por WOS |
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
