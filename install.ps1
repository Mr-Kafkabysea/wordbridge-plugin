param([switch]$CheckOnly, [string]$PythonPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'scripts/find_python.ps1')
try {
    $pythonExecutable = Find-WordBridgePython -PythonPath $PythonPath
    $codexCommand = Get-Command codex -ErrorAction Stop
    Write-Host "Python: $pythonExecutable"
    $installArgs = @((Join-Path $PSScriptRoot 'scripts/install_plugin.py'), '--codex', $codexCommand.Source)
    if ($CheckOnly) { $installArgs += '--check' }
    & $pythonExecutable @installArgs
    exit $LASTEXITCODE
} catch {
    Write-Error -ErrorAction Continue $_.Exception.Message
    exit 1
}
