# BIDCOMAGRO-SYSTEM — Control de Versiones

Cada archivo tiene un comentario `// @version X.Y` (JS) o `<!-- @version X.Y -->` (HTML).
Regla: **incrementar la versión cada vez que se edita un archivo**.
- Corrección de bug → X.Y+1 (ej: 1.0 → 1.1)
- Feature nueva → X+1.0 (ej: 1.3 → 2.0)

---

## HUB_PRO
| Archivo | Versión | Notas |
|---------|---------|-------|
| HUB_Código.js | 1.1 | Agrega HUB_generarPedidoRepuestos + constantes WOS |
| Index.html | 1.1 | Agrega modal REP→WOS, botón, y JS asociado |
| Env.js | 1.0 | Configuración inicial |

## PORTAL_RESELLER
| Archivo | Versión | Notas |
|---------|---------|-------|
| Index.html | 1.1 | Fix enviarPedido() → generarPedidoRepuestosPortal |
| RS_Pedidos.js | 1.4 | Fix PDF: total ocupa cols 6-7 merge (160px) en vez de col 7 sola (80px) |
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
| Despacho_Código.js | 2.7 | Optimizaciones: elimina funciones duplicadas (_lookupResellerEmail, _emailHead/_emailFoot), unifica cache key CARMEN, enCaminoMap en WOS_cargarStock llama a WOS_getEnCaminoMap(), alinea lógica de compras activas en reporteBackorder, batch writes en WOS_revertirAPreparado, elimina logs debug |
| Despacho_Index.html | 2.5 | Fix: ítems Cancelados no se incluyen en el modal de despacho |
| WOS_GmailFlow.js | 1.5 | Seriales: hoja anexo separada en PDF (página 2), nota al pie de la NE |
| Despacho_Env.js | 1.0 | |

## STOCK_MANAGER
| Archivo | Versión | Notas |
|---------|---------|-------|
| SM_Index.html | 1.1 | Validación: bloquea CAS con 0 productos |
| Sm_Código.js | 1.0 | |
| Env.js | 1.0 | |

## LAUNCHER
| Archivo | Versión | Notas |
|---------|---------|-------|
| Launcher_Código.js | 1.0 | |
| Launcher_Index.html | 1.0 | |

## Raíz
| Archivo | Versión | Notas |
|---------|---------|-------|
| SharedLogger.js | 1.0 | |
