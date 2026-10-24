$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
Push-Location -LiteralPath $projectRoot
try {
    & "$projectRoot\.venv\Scripts\python.exe" tools/build.py
    if ($LASTEXITCODE -ne 0) { throw 'Build thất bại' }
} finally { Pop-Location }
