param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$pythonLauncher = Get-Command py -ErrorAction Stop
$codexCommand = Get-Command codex -ErrorAction Stop
$installArgs = @('-3.14', (Join-Path $PSScriptRoot 'scripts/install_plugin.py'), '--codex', $codexCommand.Source)
if ($CheckOnly) { $installArgs += '--check' }
& $pythonLauncher.Source @installArgs
exit $LASTEXITCODE
