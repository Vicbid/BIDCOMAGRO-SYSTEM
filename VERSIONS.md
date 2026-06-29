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
| Despacho_Código.js | 3.0 | CANT_CANCEL col Z: demanda perdida usa CANT_CANCEL con fallback a CANCELADO viejo | Optimizaciones: elimina funciones duplicadas (_lookupResellerEmail, _emailHead/_emailFoot), unifica cache key CARMEN, enCaminoMap en WOS_cargarStock llama a WOS_getEnCaminoMap(), alinea lógica de compras activas en reporteBackorder, batch writes en WOS_revertirAPreparado, elimina logs debug |
| Despacho_Index.html | 3.0 | Pick ticket: columna Estado + fix multi-página al imprimir (position:fixed→static) |
| Despacho_Index.html | 3.3 | Pick tracker: badge ciclable reemplazado por select directo; _buildPickCajaOpts/_onPickCajaSel/_actualizarPickCajaSelects en lugar de _cyclePickCaja |
| WOS_GmailFlow.js | 1.7 | Nota de entrega agrupa items por caja con fila separadora Caja N + tracking cuando hay multi-bulto; Fix 1.6: fila TOTAL en col 6 |
| Despacho_Env.js | 1.0 | |

## STOCK_MANAGER
| Archivo | Versión | Notas |
|---------|---------|-------|
| SM_Index.html | 1.5 | Registrar CAS desde cruce: precarga ítems del sheet (sin XLS) |
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
