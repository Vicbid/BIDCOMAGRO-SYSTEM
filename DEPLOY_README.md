# Deploy Automático - 3 Módulos GAS

Scripts para automatizar el deployment de los 3 módulos a Google Apps Script en un solo comando.

## Scripts Disponibles

### Windows (PowerShell)
```powershell
.\deploy_all.ps1
```

**Requisitos:**
- PowerShell 5.0+
- `clasp` instalado y configurado (`npm install -g @google/clasp`)
- Permisos de ejecución de scripts (puede requerir: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`)

### Linux / macOS (Bash)
```bash
chmod +x deploy_all.sh
./deploy_all.sh
```

**Requisitos:**
- Bash 4.0+
- `clasp` instalado y configurado

## Qué Hace

El script ejecuta automáticamente para los 3 módulos (**HUB_PRO**, **PORTAL_RESELLER**, **STOCK_MANAGER**):

1. **Fase 1: Push**
   - Entra a cada carpeta
   - Ejecuta `clasp push` para subir el código actualizado

2. **Fase 2: Redeploy**
   - Obtiene el último deployment ID usando `clasp deployments`
   - Ejecuta `clasp redeploy <ID>` para activar la nueva versión

## Salida Esperada

```
══════════════════════════════════════════════════════════════
  DEPLOYMENT AUTOMÁTICO - 3 MÓDULOS
══════════════════════════════════════════════════════════════

FASE 1: Pushing código a Google Apps Script...

📦 Procesando: HUB_PRO
   ➜ clasp push...
   ✅ Push completado

📦 Procesando: PORTAL_RESELLER
   ➜ clasp push...
   ✅ Push completado

📦 Procesando: STOCK_MANAGER
   ➜ clasp push...
   ✅ Push completado

FASE 2: Obteniendo deployment IDs y redeployando...

🔄 Redeployando: HUB_PRO
   ➜ Obteniendo deployment ID...
   📌 ID encontrado: AKfycbyXXXXXXXXX
   ➜ clasp redeploy AKfycbyXXXXXXXXX...
   ✅ Redeploy completado

🔄 Redeployando: PORTAL_RESELLER
   ➜ Obteniendo deployment ID...
   📌 ID encontrado: AKfycbyYYYYYYYYY
   ➜ clasp redeploy AKfycbyYYYYYYYYY...
   ✅ Redeploy completado

🔄 Redeployando: STOCK_MANAGER
   ➜ Obteniendo deployment ID...
   📌 ID encontrado: AKfycbyZZZZZZZZZ
   ➜ clasp redeploy AKfycbyZZZZZZZZZ...
   ✅ Redeploy completado

══════════════════════════════════════════════════════════════
  ✅ DEPLOYMENT COMPLETADO
══════════════════════════════════════════════════════════════

Resumen de deployments:

  ✓ HUB_PRO → AKfycbyXXXXXXXXX
  ✓ PORTAL_RESELLER → AKfycbyYYYYYYYYY
  ✓ STOCK_MANAGER → AKfycbyZZZZZZZZZ
```

## Manejo de Errores

- Si `clasp push` falla en algún módulo, el script se detiene y muestra el error
- Si `clasp deployments` no encuentra un ID, el script saltea ese módulo con una advertencia ⚠️
- Si `clasp redeploy` falla, el script se detiene

## Flujo Rápido

```bash
# Windows
cd d:\BIDCOMAGRO-SYSTEM
.\deploy_all.ps1

# Linux/macOS
cd /path/to/BIDCOMAGRO-SYSTEM
./deploy_all.sh
```

## Requisito Previo: Configurar clasp

Si aún no lo has hecho:

```bash
npm install -g @google/clasp
clasp login
```

Luego, cada carpeta debe tener un `.clasp.json`:

```json
{
  "scriptId": "YOUR_SCRIPT_ID",
  "rootDir": "."
}
```

## Notas

- Los scripts se pueden ejecutar múltiples veces sin problemas
- Cada módulo mantiene su propio `.clasp.json` y `appsscript.json`
- El script no modifica ningún archivo, solo ejecuta comandos
- Los deployment IDs son únicos por módulo
