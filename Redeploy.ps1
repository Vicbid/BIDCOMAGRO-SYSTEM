param(
    [switch]$stable,
    [string]$tag = ""
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

$ROOT = "D:\BIDCOMAGRO-SYSTEM"
$MODS = @("HUB_PRO", "PORTAL_RESELLER", "STOCK_MANAGER", "LAUNCHER", "WOS", "ComandasPedidos")

function Get-ModuleHash {
    param([string]$path)
    $files = Get-ChildItem $path -Include "*.js","*.html" -Recurse |
             Where-Object { $_.Name -ne "appsscript.json" } |
             Sort-Object FullName
    $hashes = $files | ForEach-Object { (Get-FileHash $_.FullName -Algorithm MD5).Hash }
    return ($hashes -join "|")
}

Write-Host "--- INICIANDO DEPLOY ---"
$whoami = (clasp show-authorized-user 2>&1 | Select-Object -First 1)
Write-Host "  clasp login: $whoami" -ForegroundColor DarkGray
if ($stable) { Write-Host "  Modo: ESTABLE (commit en main + tag)" -ForegroundColor Magenta }
else          { Write-Host "  Modo: BETA (commit en main)" -ForegroundColor Cyan }

$changedModules = @()

foreach ($m in $MODS) {
    $p = Join-Path $ROOT $m
    if (-not (Test-Path $p)) {
        Write-Host "No encontrado: $m"
        continue
    }

    Write-Host ""
    Write-Host "Modulo: $m"

    $hashFile    = Join-Path $p ".deploy-hash"
    $currentHash = Get-ModuleHash $p
    $savedHash   = ""
    if (Test-Path $hashFile) { $savedHash = (Get-Content $hashFile -Raw).Trim() }

    if ($currentHash -eq $savedHash) {
        Write-Host "  Sin cambios - push y redeploy omitidos." -ForegroundColor Cyan
        continue
    }

    $changedModules += $m

    Set-Location $p
    Write-Host "  Cambios detectados - pusheando..."
    $pushOut = clasp push -f
    $pushOut | ForEach-Object { Write-Host "    $_" }

    if ($pushOut -match "already up to date") {
        Write-Host "  GAS HEAD ya actualizado - forzando redeploy igualmente..." -ForegroundColor Yellow
    }

    # Lee el ID guardado o lo descubre desde clasp deployments
    $deployIdFile = Join-Path $p ".deploy-id"
    $id = ""
    if (Test-Path $deployIdFile) {
        $id = (Get-Content $deployIdFile -Raw).Trim()
        Write-Host "  Deployment ID guardado: $id"
    }

    if (-not $id) {
        # Primera vez: buscar el deployment versionado (excluye @HEAD que es solo para test)
        $claspOut = clasp deployments
        Write-Host "  Buscando deployment en:"
        $claspOut | ForEach-Object { Write-Host "    $_" }

        $versMatches = $claspOut | Select-String -Pattern "- ([a-zA-Z0-9\-_]+) @[0-9]+" -CaseSensitive:$false |
                       Where-Object { $_ -notmatch "@HEAD" }

        if ($versMatches) {
            $id = $versMatches[-1].Matches.Groups[1].Value
            $id | Set-Content $deployIdFile -Encoding utf8 -NoNewline
            Write-Host "  ID encontrado y guardado: $id" -ForegroundColor Yellow
        }
    }

    if ($id) {
        $res = clasp deploy --deploymentId $id -d "Auto-update" 2>&1
        if ($LASTEXITCODE -ne 0 -or ($res -match "error" -and $res -notmatch "Updated")) {
            Write-Host "  Error al actualizar deployment:" -ForegroundColor Red
            $res | ForEach-Object { Write-Host "    $_" }
            Write-Host "  ATENCION: no se creo un deployment nuevo para evitar re-auth." -ForegroundColor Red
            Write-Host "  Verificar el ID en $deployIdFile y correr de nuevo." -ForegroundColor Red
        } else {
            Write-Host "  $m actualizado con exito (mismo deployment ID)" -ForegroundColor Green
        }
    } else {
        # Nunca hubo deployment: crear el primero y guardar el ID
        Write-Host "  Sin deployment previo. Creando el primero..." -ForegroundColor Yellow
        $res = clasp deploy -d "Primer deploy" 2>&1
        $res | ForEach-Object { Write-Host "    $_" }
        $newMatch = ($res | Select-String -Pattern "- ([a-zA-Z0-9\-_]+) @[0-9]+")
        if ($newMatch) {
            $newId = $newMatch.Matches.Groups[1].Value
            $newId | Set-Content $deployIdFile -Encoding utf8 -NoNewline
            Write-Host "  ID del nuevo deployment guardado: $newId" -ForegroundColor Green
            Write-Host "  RECORDATORIO: autorizar el web app en GAS antes de usar." -ForegroundColor Yellow
        }
    }

    $currentHash | Set-Content $hashFile -Encoding utf8 -NoNewline
}

# -- Git commit ------------------------------------------------
Set-Location $ROOT

if ($changedModules.Count -eq 0) {
    Write-Host ""
    Write-Host "--- Sin cambios para commitear ---" -ForegroundColor Cyan
    Write-Host "--- FIN ---"
    exit 0
}

$moduleList = $changedModules -join ", "
$prefix     = if ($stable) { "release" } else { "deploy" }
$commitMsg  = "${prefix}: ${moduleList}"

Write-Host ""
Write-Host "--- GIT ---"

# Rama unica: main
$currentBranch = git rev-parse --abbrev-ref HEAD
if ($currentBranch -ne "main") {
    Write-Host "  Cambiando a main..." -ForegroundColor Yellow
    git checkout main
}
git add VERSIONS.md WOS PORTAL_RESELLER HUB_PRO STOCK_MANAGER LAUNCHER ComandasPedidos 2>$null

if ($stable) {
    # Calcular tag si no se paso uno
    if (-not $tag) {
        $lastTag = git describe --tags --abbrev=0 2>$null
        if ($lastTag -match "^v(\d+)\.(\d+)$") {
            $tag = "v$($Matches[1]).$([int]$Matches[2] + 1)"
        } else {
            $tag = "v1.1"
        }
    }
    git commit -m "release: $tag - $moduleList"
    git tag $tag
    Write-Host "  Release $tag listo en main (tag $tag)" -ForegroundColor Green
    Write-Host "  (pushea con: git push origin main --tags)" -ForegroundColor DarkGray

} else {
    git commit -m $commitMsg
    Write-Host "  Commit en main: $commitMsg" -ForegroundColor Green
    Write-Host "  (pushea con: git push origin main)" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "--- FIN ---"
