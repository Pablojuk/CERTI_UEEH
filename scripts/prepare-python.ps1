param(
    [string]$PythonExecutable = $env:CERTI_BUILD_PYTHON
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

if (-not $PythonExecutable) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCommand) {
        throw 'No se encontró Python. Defina CERTI_BUILD_PYTHON con la ruta a python.exe.'
    }
    $PythonExecutable = $pythonCommand.Source
}

if (-not (Test-Path -LiteralPath $PythonExecutable -PathType Leaf)) {
    throw "El ejecutable Python no existe: $PythonExecutable"
}

& $PythonExecutable -c 'import PyInstaller' 2>$null
if ($LASTEXITCODE -ne 0) {
    throw 'PyInstaller no está instalado. Ejecute: python -m pip install -r requirements-build.txt'
}

$buildRoot = Join-Path $repoRoot 'build'
$distRoot = Join-Path $buildRoot 'python'
$workRoot = Join-Path $buildRoot 'pyinstaller'
$specPath = Join-Path $buildRoot 'certi-python.spec'
$bundlePath = Join-Path $distRoot 'certi-python'
$expectedExe = Join-Path $bundlePath 'certi-python.exe'
$repoPrefix = $repoRoot.TrimEnd('\') + '\'

foreach ($target in @($bundlePath, $workRoot, $specPath)) {
    $absoluteTarget = [System.IO.Path]::GetFullPath($target)
    if (-not $absoluteTarget.StartsWith($repoPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Ruta de limpieza fuera del repositorio: $absoluteTarget"
    }
    if (Test-Path -LiteralPath $absoluteTarget) {
        Remove-Item -LiteralPath $absoluteTarget -Recurse -Force
    }
}

New-Item -ItemType Directory -Force -Path $buildRoot, $distRoot, $workRoot | Out-Null

$separator = [System.IO.Path]::PathSeparator
$arguments = @(
    '-m', 'PyInstaller',
    '--noconfirm',
    '--clean',
    '--onedir',
    '--name', 'certi-python',
    '--distpath', $distRoot,
    '--workpath', $workRoot,
    '--specpath', $buildRoot,
    '--add-data', "$(Join-Path $repoRoot 'catalogo_asignaturas.json')$separator.",
    '--add-data', "$(Join-Path $repoRoot 'escala_cualitativa.json')$separator.",
    '--add-data', "$(Join-Path $repoRoot 'assets\certificados')${separator}assets\certificados",
    (Join-Path $repoRoot 'procesador_notas.py')
)

$previousNoUserSite = $env:PYTHONNOUSERSITE
$env:PYTHONNOUSERSITE = '1'
try {
    & $PythonExecutable @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller terminó con código $LASTEXITCODE."
    }
} finally {
    if ($null -eq $previousNoUserSite) {
        Remove-Item Env:PYTHONNOUSERSITE -ErrorAction SilentlyContinue
    } else {
        $env:PYTHONNOUSERSITE = $previousNoUserSite
    }
}

if (-not (Test-Path -LiteralPath $expectedExe -PathType Leaf)) {
    throw "No se generó el ejecutable esperado: $expectedExe"
}

$templates = Get-ChildItem -LiteralPath (Join-Path $repoRoot 'assets\certificados') -Filter '*.html'
if ($templates.Count -eq 0) {
    throw 'No se encontraron plantillas de certificados para empaquetar.'
}

Write-Output "Recurso Python preparado: $expectedExe"
