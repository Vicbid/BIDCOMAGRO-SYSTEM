# PORTAL RESELLER — Contexto del proyecto

## Sheet ID
1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc (mismo que HUB)

## Archivos
- RS_Código.js — backend Apps Script
- Index.html — frontend

## Autenticación
- PIN de 4 dígitos, col K de hoja Resellers (índice 10)
- Variable global: `resellerActual` — se llena al autenticarse

## Hojas que usa
- `Ordenes de trabajo` — lectura via _leerOrdenes()
- `Resellers` — col A=nombre, col J(9)=email, col K(10)=PIN
- `EQUIPOS` — col A=nombre, col B=tipo (ej: "Bateria")
- `DB_REPUESTOS` — B(1)=código, C(2)=nombre, F(5)=PVP
- `DEUDA_RESELLERS` — backorder por reseller
- `EMAIL_LOGS` — historial de emails enviados

## Mapa columnas Ordenes de trabajo (lectura)
- f[2]=OT | f[3]=Garantía | f[4]=Estado | f[5]=Equipo
- f[6]=SN | f[7]=Reseller | f[9]=Técnico | f[11]=Mensajes
- f[13]=Fecha activación | f[14]=CAS/FWR | f[16]=Repuestos
- f[17]=Prioridad | f[18]=Circuito

## Flujos de estado por circuito
- Taller: Abierto → En Revision → Presupuesto enviado → Presupuesto rechazado → Presupuesto aceptado → Espera de repuestos → En reparacion → Finalizado
- Reseller: Abierto → En reparacion → Pedido de repuestos → Repuestos enviados → Informe de reparacion → Recepcion piezas dañadas → Aprobacion DJI → Finalizado
- Reseller Propio (no batería): Abierto → Pedido de repuesto para reparar → Reparado y aprobado en el aftersales → Partes dañadas scrapeadas → Finalizado
- Reseller Propio (batería IW): Abierto → Aprobado por DJI → Finalizado

## Detección flujo batería
- Consultar col B de hoja EQUIPOS — si dice "Bateria" → flujo especial
- Solo aplica a Reseller Propio + IW

## Normalización circuito (col S / f[18])
```js
var circRaw = String(f[18]||"").trim().toUpperCase();
var flujo = "Taller";
if (circRaw === "RESELLER" || circRaw === "SI") flujo = "Reseller";
else if (circRaw === "RESELLER PROPIO") flujo = "Reseller Propio";
```

## CONFIG portal
- COL_PIN_RESELLER: 10 (col K)
- COL_FECHA_ACTIVACION: 13 (col N)
- DIAS_AVISO_GARANTIA: 30

## Funciones backend principales
- autenticarReseller(pin) → {ok, nombre}
- consultarEstado(ot, sn) → detalle OT con flujo normalizado
- buscarOTsReseller(nombre) → lista OTs con badges
- agregarComentario(ot, texto, autor)
- agregarFechaActivacion(ot, fechaStr) — escritura única col N
- obtenerGarantias(reseller) → garantías por vencer
- obtenerDeudaReseller(reseller) → repuestos pendientes
- confirmarRecepcionRepuestos(ot, sn) → cambia estado
- obtenerEmailLogsPorReseller(reseller) → historial presupuestos

## Reglas importantes
- ES5 estricto (var, no arrow functions)
- GmailApp.sendEmail(), nunca MailApp
- Singleton getSS() / getSheet()
- Col N (fecha activación) escritura única — no se puede modificar una vez guardada
- Mensajes reseller en col L — bidireccionales con el HUB