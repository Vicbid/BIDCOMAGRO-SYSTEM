# DJI HUB PRO — Contexto del proyecto

## Sheet ID
1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc

## Archivos
- Codigo.gs — backend Apps Script (getActiveSpreadsheet)
- Index.html — frontend SPA

## Mapa COMPLETO de columnas (Ordenes de trabajo)
| Col | Índice | Campo | Notas |
|-----|--------|-------|-------|
| A | 0 | Fecha apertura | |
| B | 1 | Fecha finalización | Se escribe al Finalizar |
| C | 2 | OT | Formato WH/REP/00000 |
| D | 3 | Garantía | IW / OOW |
| E | 4 | Estado | |
| F | 5 | Equipo/Modelo | |
| G | 6 | S/N | |
| H | 7 | Reseller | |
| I | 8 | Ubicación | libre |
| J | 9 | Técnico | |
| K | 10 | Cliente | libre |
| L | 11 | Mensajes reseller | Canal bidireccional Portal↔HUB |
| M | 12 | Informe técnico | |
| N | 13 | Fecha activación | Escritura única desde Portal |
| O | 14 | CAS/FWR | |
| P | 15 | Entrega repuestos | ⚠️ NUNCA TOCAR |
| Q | 16 | Repuestos | Formato: COD \| Desc \| P:X E:X separado por ` ; ` |
| R | 17 | Prioridad | URGENTE / NORMAL |
| S | 18 | Circuito | Taller / Reseller / Reseller Propio |
| T | 19 | Comanda1 | ⚠️ NUNCA TOCAR |
| U | 20 | Sello estado | Fecha del último cambio de estado |
| V | 21 | Checklist QC | JSON |
| W | 22 | Mano de obra | JSON array [{codigo,descripcion,precio}] |
| X | 23 | Notas internas | Texto libre, NO visible para resellers |

## Otras hojas
- `LOGS` — fecha, OT, técnico, email, acción, estadoAnt, estadoNvo, detalle
- `EMAIL_LOGS` — fecha, OT, destinatario, rol, asunto, estado
- `Resellers` — col A=nombre, col J(9)=email, col K(10)=PIN
- `DB_REPUESTOS` — B(1)=código, C(2)=nombre, D(3)=modelos, F(5)=PVP USD, G(6)=FOB
- `EQUIPOS` — col A=nombre, col B=tipo (ej: "Bateria"), col D(3)=meses garantía
- `Usuarios_Internos` — A=nombre, B=email, C=rol(admin/operador), D=esTecnico(si/no)
- `Precios_mano_obra` — A=código/nivel, B=descripción, C=precio sugerido
- `DEUDA_RESELLERS` — backorder de repuestos por reseller

## Usuarios internos
- Victor Almao (admin, técnico)
- Juan Cuaresma (admin, técnico)
- Jeremías Silvestre (operador, técnico)
- Gestion Logistica (sin rol, no técnico) — recibe emails de reposición batería

## Flujos de estado por circuito
**Taller:** Abierto → En Revision → Presupuesto enviado → Presupuesto rechazado → Presupuesto aceptado → Espera de repuestos → En reparacion → Finalizado

**Reseller:** Abierto → En reparacion → Pedido de repuestos → Repuestos enviados → Informe de reparacion → Recepcion piezas dañadas → Aprobacion DJI → Finalizado

**Reseller Propio:** Abierto → Pedido de repuesto para reparar → Reparado y aprobado en el aftersales → Partes dañadas scrapeadas → Finalizado

**Reseller Propio + Batería IW:** Abierto → Aprobado por DJI → Finalizado

## Normalización circuito (col S)
```js
var circUp = String(f[18]||"").trim().toUpperCase();
var tipo = "Taller";
if (circUp === "SI" || circUp === "RESELLER") tipo = "Reseller";
else if (circUp === "RESELLER PROPIO") tipo = "Reseller Propio";
```

## Singleton y helpers
```js
var _SS = null;
function SS() { if (!_SS) _SS = SpreadsheetApp.getActiveSpreadsheet(); return _SS; }
function getSheet(nombre) { return SS().getSheetByName(nombre); }
```

## Reglas INNEGOCIABLES
- **ES5 estricto** — var, function, sin arrow functions, sin const/let
- **NUNCA escribir** col P (idx 15) ni col T (idx 19) desde el sistema
- **GmailApp.sendEmail()** siempre, nunca MailApp
- **Repuestos formato**: `Cod | Desc | P:X E:X` separados por ` ; `
- **DEUDA_RESELLERS**: reescritura en bloque (clearContents + setValues)
- **LockService** en crearNuevaOT para evitar duplicados de concurrencia

## Email config
- EMAIL_SUPERVISOR: soporteagrasdji@bidcom.com.ar
- NOMBRE_REMITENTE: "BIDCOMAGRO · Soporte Técnico DJI Agriculture"
- replyTo: EMAIL_SUPERVISOR

## Lógica de reposición batería
- Al guardar estado "Aprobado por DJI" en circuito Reseller Propio
- Si equipo es tipo "Bateria" (col B de EQUIPOS) → enviar email a Gestion Logistica
- Email con copia al reseller

## Presupuesto OOW
- Solo para OTs OOW en estado "En Revision"
- Descuento 40% para reseller, 25% para cliente final
- Mano de obra seleccionable desde Precios_mano_obra
- Se guarda en col W (idx 22) al guardar la OT

## Funciones backend principales
- cargarTodo() — carga única al inicio, devuelve ordenes + catálogo completo
- actualizarOrden(data) — escribe col por col, llama notificaciones
- crearNuevaOT() — con LockService
- enviarPresupuesto(data) — email con repuestos + MO + descuentos
- enviarEmailReposicionBateria(data) — trigger automático
- enviarNotificaciones(data, estadoAnt, tec) — emails automáticos
- obtenerDatosSupervisor() — backorder predictivo + SLA
- buscarHistorialSN(sn) — reincidentes
- identificarUsuario() — SSO por email de Google