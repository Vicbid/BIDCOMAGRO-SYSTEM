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

Write-Host "--- DEPLOY BETA ---" -ForegroundColor Cyan
$whoami = (clasp show-authorized-user 2>&1 | Select-Object -First 1)
Write-Host "  clasp login: $whoami" -ForegroundColor DarkGray

$changedModules = @()

foreach ($m in $MODS) {
    $p = Join-Path $ROOT $m
    if (-not (Test-Path $p)) { Write-Host "No encontrado: $m"; continue }

    Write-Host ""
    Write-Host "Modulo: $m"

    $hashFile    = Join-Path $p ".deploy-hash"
    $currentHash = Get-ModuleHash $p
    $savedHash   = if (Test-Path $hashFile) { (Get-Content $hashFile -Raw).Trim() } else { "" }

    if ($currentHash -eq $savedHash) {
        Write-Host "  Sin cambios." -ForegroundColor Cyan
        continue
    }

    $changedModules += $m

    Set-Location $p
    Write-Host "  Pusheando a GAS..."
    $pushOut = clasp push -f
    $pushOut | ForEach-Object { Write-Host "    $_" }

    $deployIdFile = Join-Path $p ".deploy-id"
    $id = ""
    if (Test-Path $deployIdFile) {
        $id = (Get-Content $deployIdFile -Raw).Trim()
        Write-Host "  Deployment ID guardado: $id"
    }

    if (-not $id) {
        $claspOut    = clasp deployments
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
            Write-Host "  ATENCION: verificar el ID en $deployIdFile" -ForegroundColor Red
        } else { Write-Host "  $m OK" -ForegroundColor Green }
    } else {
        Write-Host "  Sin deployment previo. Creando..." -ForegroundColor Yellow
        $res = clasp deploy -d "Primer deploy" 2>&1
        $res | ForEach-Object { Write-Host "    $_" }
        $newMatch = ($res | Select-String -Pattern "- ([a-zA-Z0-9\-_]+) @[0-9]+")
        if ($newMatch) {
            $newId = $newMatch.Matches.Groups[1].Value
            $newId | Set-Content $deployIdFile -Encoding utf8 -NoNewline
            Write-Host "  ID guardado: $newId - autorizar en GAS antes de usar." -ForegroundColor Yellow
        }
    }

    $currentHash | Set-Content $hashFile -Encoding utf8 -NoNewline
}

# -- Git --
Set-Location $ROOT

if ($changedModules.Count -eq 0) {
    Write-Host ""
    Write-Host "--- Sin cambios ---" -ForegroundColor Cyan
    exit 0
}

$currentBranch = git rev-parse --abbrev-ref HEAD
if ($currentBranch -ne "main") {
    Write-Host "  Cambiando a main..." -ForegroundColor Yellow
    git checkout main
}

git add VERSIONS.md WOS PORTAL_RESELLER HUB_PRO STOCK_MANAGER LAUNCHER ComandasPedidos 2>$null
$msg = "deploy: $($changedModules -join ', ')"
git commit -m $msg
Write-Host "  Commit en main: $msg" -ForegroundColor Green
Write-Host "  (push: git push origin main)" -ForegroundColor DarkGray

Write-Host ""
Write-Host "--- FIN BETA ---"
