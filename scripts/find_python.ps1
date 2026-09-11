# Probe candidates; a launcher existing on PATH does not prove Python is installed.
function Test-WordBridgePython {
    param([string]$Command, [string[]]$Prefix = @())
    $probe = "import json,sys,struct; print(json.dumps({'version':list(sys.version_info[:2]),'bits':struct.calcsize('P')*8,'executable':sys.executable}))"
    try {
        $result = & $Command @Prefix -c $probe 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
        $info = ($result -join "`n") | ConvertFrom-Json -ErrorAction Stop
        if ($info.version.Count -ne 2 -or $info.version[0] -ne 3 -or $info.version[1] -ne 14 -or $info.bits -ne 64) {
            return $null
        }
        if (-not [System.IO.Path]::IsPathRooted($info.executable) -or
            -not (Test-Path -LiteralPath $info.executable -PathType Leaf)) { return $null }
        return [string]$info.executable
    } catch { return $null }
}

function Find-WordBridgePython {
    param([string]$PythonPath)
    if ($PythonPath) {
        if (-not [System.IO.Path]::IsPathRooted($PythonPath) -or
            -not (Test-Path -LiteralPath $PythonPath -PathType Leaf)) {
            throw 'PythonPath must be an existing absolute path to python.exe.'
        }
        $resolved = Test-WordBridgePython -Command $PythonPath
        if (-not $resolved) { throw 'PythonPath must point to a working 64-bit Python 3.14.' }
        return $resolved
    }
    foreach ($name in @('py', 'python', 'python3')) {
        $candidate = Get-Command $name -CommandType Application -ErrorAction SilentlyContinue
        if (-not $candidate) { continue }
        # Do not launch Microsoft Store execution aliases.
        if ($candidate.Source -match '[\\/]Microsoft[\\/]WindowsApps[\\/]') { continue }
        $prefixArgs = @()
        if ($name -eq 'py') { $prefixArgs = @('-3.14') }
        $resolved = Test-WordBridgePython -Command $candidate.Source -Prefix $prefixArgs
        if ($resolved) { return $resolved }
    }
    throw 'No working 64-bit Python 3.14 found. Install it, or pass -PythonPath with the absolute path to python.exe.'
}
