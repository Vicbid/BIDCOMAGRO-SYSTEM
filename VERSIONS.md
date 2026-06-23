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
| RS_Pedidos.js | 1.1 | Fix getActiveSheet() → getSheetByName('Pedidos_resellers') |
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
| Despacho_Código.js | 2.4 | Backorder en hilo persistente: Thread ID en WOS_CONFIG, reply a destinatarios actuales |
| Despacho_Index.html | 2.1 | Revierte botón re-enviar confirmación (cancelado) |
| WOS_GmailFlow.js | 1.1 | Fix: elimina bloqueo duro por Thread_ID ausente; fallback a email nuevo ya existía |
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
