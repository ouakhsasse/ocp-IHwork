$ErrorActionPreference = "Stop"

function Resolve-ProjectRoot {
    $currentDirectory = (Get-Location).Path
    $nestedProjectDirectory = Join-Path $currentDirectory "HK_energies-main"
    $scriptProjectDirectory = Resolve-Path (Join-Path $PSScriptRoot "..")

    if (Test-Path (Join-Path $currentDirectory "package.json")) {
        return $currentDirectory
    }

    if (Test-Path (Join-Path $nestedProjectDirectory "package.json")) {
        return $nestedProjectDirectory
    }

    if (Test-Path (Join-Path $scriptProjectDirectory "package.json")) {
        return $scriptProjectDirectory.Path
    }

    throw "Could not find package.json. Run this script from the project root or from the parent folder that contains HK_energies-main."
}

$projectRoot = Resolve-ProjectRoot
Write-Host "Using project directory: $projectRoot"
Set-Location $projectRoot

if (-not (Test-Path (Join-Path $projectRoot "node_modules"))) {
    Write-Host "node_modules not found. Running npm install..."
    & npm install
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

& npm run dev
exit $LASTEXITCODE

