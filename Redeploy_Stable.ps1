param([string]$tag = "")

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

Write-Host "--- DEPLOY ESTABLE (forzado) ---" -ForegroundColor Magenta
$whoami = clasp whoami 2>&1
Write-Host "  clasp login: $whoami" -ForegroundColor DarkGray

$changedModules = @()

foreach ($m in $MODS) {
    $p = Join-Path $ROOT $m
    if (-not (Test-Path $p)) { Write-Host "No encontrado: $m"; continue }

    Write-Host ""
    Write-Host "Modulo: $m"

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

    # Actualizar hash para que Beta no repita lo que ya subio Stable
    $currentHash = Get-ModuleHash $p
    $hashFile    = Join-Path $p ".deploy-hash"
    $currentHash | Set-Content $hashFile -Encoding utf8 -NoNewline

    $changedModules += $m
}

# -- Git --
Set-Location $ROOT

# Calcular tag si no se paso uno
if (-not $tag) {
    $lastTag = git describe --tags --abbrev=0 2>$null
    if ($lastTag -match "^v(\d+)\.(\d+)$") {
        $tag = "v$($Matches[1]).$([int]$Matches[2] + 1)"
    } else {
        $tag = "v1.1"
    }
}

$moduleList = $changedModules -join ", "

# Commit en dev
$currentBranch = git rev-parse --abbrev-ref HEAD
if ($currentBranch -ne "dev") {
    Write-Host "  Cambiando a dev..." -ForegroundColor Yellow
    git checkout dev
}
git add VERSIONS.md WOS PORTAL_RESELLER HUB_PRO STOCK_MANAGER LAUNCHER ComandasPedidos 2>$null
git commit -m "release: $moduleList"
Write-Host "  Commit en dev OK" -ForegroundColor Green

# Merge a main + tag
Write-Host "  Mergeando a main con tag $tag..." -ForegroundColor Magenta
git checkout main
git merge dev --no-ff -m "release: $tag - $moduleList"
git tag $tag
git checkout dev

Write-Host "  Tag $tag creado en main" -ForegroundColor Green
Write-Host "  (push: git push origin main dev --tags)" -ForegroundColor DarkGray

Write-Host ""
Write-Host "--- FIN ESTABLE ---"
