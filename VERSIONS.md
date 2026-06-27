# BIDCOMAGRO-SYSTEM — Control de Versiones

Cada archivo tiene un comentario `// @version X.Y` (JS) o `<!-- @version X.Y -->` (HTML).
Regla: **incrementar la versión cada vez que se edita un archivo**.
- Corrección de bug → X.Y+1 (ej: 1.0 → 1.1)
- Feature nueva → X+1.0 (ej: 1.3 → 2.0)

---

## HUB_PRO
| Archivo | Versión | Notas |
|---------|---------|-------|
| HUB_Código.js | 1.9 | Sin respuesta · Cerrado y Rechazado DJI en todos los checks de estado terminal |
| Index.html | 1.8 | Sin respuesta · Cerrado y Rechazado DJI excluidos de todos los filtros de pendientes/urgentes/KPIs |
| GUIA_FLUJOS.html | 1.2 | Badges de turno (⚽ Bidcom / ⚽ Reseller / ⏳ DJI) en headers de todos los cards; estado Sin respuesta · Cerrado en ambas secciones Reseller |
| Env.js | 1.2 | Agrega SCHEMA.EQUIPOS (NOMBRE, TIPO, PREFIJO, MESES) |

## PORTAL_RESELLER
| Archivo | Versión | Notas |
|---------|---------|-------|
| Index.html | 1.1 | Fix enviarPedido() → generarPedidoRepuestosPortal |
| RS_Pedidos.js | 1.5 | Fórmula CANT_PEND col G incluye -Z (CANT_CANCEL) al crear filas | Fix PDF: total ocupa cols 6-7 merge (160px) en vez de col 7 sola (80px) |
| RS_OTs.js | 1.5 | Fix: CANCELADO ya no cae en paso 3; garantías excluyen CANCELADO y Entregado |
| Index.html | 1.2 | Fix: stepper muestra "Orden cancelada" en lugar de paso 3 para OTs CANCELADAS |
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
| WOS_GmailFlow.js | 1.6 | Fix fila TOTAL: valor iba a col 7 (spacer invisible) en vez de col 6 (Subtotal) |
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
