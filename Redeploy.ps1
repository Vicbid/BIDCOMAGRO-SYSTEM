$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

$ROOT = "D:\BIDCOMAGRO-SYSTEM"
$MODS = @("HUB_PRO", "PORTAL_RESELLER", "STOCK_MANAGER", "LAUNCHER", "WOS")

function Get-ModuleHash {
    param([string]$path)
    $files = Get-ChildItem $path -Include "*.js","*.html" -Recurse |
             Where-Object { $_.Name -ne "appsscript.json" } |
             Sort-Object FullName
    $hashes = $files | ForEach-Object { (Get-FileHash $_.FullName -Algorithm MD5).Hash }
    return ($hashes -join "|")
}

Write-Host "--- INICIANDO DEPLOY ---"

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

    Set-Location $p
    Write-Host "  Cambios detectados - pusheando..."
    $pushOut = clasp push -f
    $pushOut | ForEach-Object { Write-Host "    $_" }

    if ($pushOut -match "already up to date") {
        Write-Host "  GAS HEAD ya actualizado - forzando redeploy igualmente..." -ForegroundColor Yellow
    }

    $claspOut = clasp deployments
    Write-Host "  Salida de clasp deployments:"
    $claspOut | ForEach-Object { Write-Host "    $_" }

    $allMatches = $claspOut | Select-String -Pattern "- ([a-zA-Z0-9\-_]+) @\d+" -CaseSensitive:$false
    $webMatch   = $claspOut | Select-String -Pattern "- ([a-zA-Z0-9\-_]+) @\d+.*web.*" -CaseSensitive:$false

    $id       = $null
    $criterio = "ninguno"
    if ($webMatch) {
        $id       = $webMatch[0].Matches.Groups[1].Value
        $criterio = "web app detectado"
    } elseif ($allMatches) {
        $id       = $allMatches[-1].Matches.Groups[1].Value
        $criterio = "fallback: ultimo deployment"
    }

    Write-Host "  ID elegido: '$id' (criterio: $criterio)"

    if ($id) {
        $res = clasp redeploy $id 2>&1
        if ($res -match "Read-only" -or $res -match "error") {
            Write-Host "  Bloqueado. Creando nuevo deployment..." -ForegroundColor Yellow
            clasp deploy -d "Auto-update"
        } else {
            Write-Host "  $m actualizado con exito" -ForegroundColor Green
        }
    } else {
        Write-Host "  Sin deployment previo. Creando..." -ForegroundColor Yellow
        clasp deploy -d "Primer deploy"
    }

    $currentHash | Set-Content $hashFile -Encoding utf8 -NoNewline
}

Set-Location $ROOT
Write-Host ""
Write-Host "--- FIN ---"
