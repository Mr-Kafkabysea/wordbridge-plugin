"""Exercise the PowerShell entry's interpreter selection without installing anything."""
import json
from pathlib import Path
import shutil
import subprocess
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
POWERSHELL = shutil.which('powershell.exe')


def quote(value):
    return "'" + str(value).replace("'", "''") + "'"


@unittest.skipUnless(POWERSHELL and sys.platform == 'win32', 'Windows PowerShell required')
class PythonLauncherTests(unittest.TestCase):
    def run_ps(self, script):
        return subprocess.run([POWERSHELL, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command',
                               "$ErrorActionPreference = 'Stop'; . " + quote(ROOT / 'scripts/find_python.ps1') + "; " + script],
                              capture_output=True, text=True, timeout=30)

    def choose(self, available, working):
        script = """
$global:attempts = @()
function Get-Command {
    param($Name, $CommandType, $ErrorAction)
    if (@(AVAILABLE) -contains $Name) { [pscustomobject]@{Source=$Name} }
}
function Test-WordBridgePython {
    param($Command, $Prefix)
    $global:attempts += $Command + ':' + ($Prefix -join ',')
    if ($Command -eq WORKING) { return 'C:\\verified\\python.exe' }
    return $null
}
try {
    $selected = Find-WordBridgePython
    @{selected=$selected;attempts=$global:attempts} | ConvertTo-Json -Compress
} catch { Write-Output $_.Exception.Message; exit 1 }
""".replace('AVAILABLE', ','.join(quote(x) for x in available)).replace('WORKING', quote(working))
        return self.run_ps(script)

    def test_registered_launcher_is_preferred(self):
        result = self.choose(['py', 'python'], 'py')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)['attempts'], ['py:-3.14'])

    def test_broken_launcher_falls_back_to_python(self):
        result = self.choose(['py', 'python'], 'python')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)['attempts'], ['py:-3.14', 'python:'])

    def test_python3_fallback(self):
        result = self.choose(['python3'], 'python3')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)['attempts'], ['python3:'])

    def test_no_compatible_interpreter_has_actionable_error(self):
        result = self.choose(['py', 'python'], 'unavailable')
        self.assertEqual(result.returncode, 1)
        self.assertIn('-PythonPath', result.stdout)

    def test_explicit_real_python_path_is_probed(self):
        result = self.run_ps('Find-WordBridgePython -PythonPath ' + quote(sys.executable))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(Path(result.stdout.strip()).resolve(), Path(sys.executable).resolve())

    def test_invalid_explicit_path_does_not_fall_back(self):
        result = self.run_ps("try { Find-WordBridgePython -PythonPath 'relative.exe' } catch { exit 7 }")
        self.assertEqual(result.returncode, 7)

    def test_store_alias_not_executed(self):
        script = """
function Get-Command {
    param($Name, $CommandType, $ErrorAction)
    [pscustomobject]@{Source='C:\\Users\\test\\AppData\\Local\\Microsoft\\WindowsApps\\python.exe'}
}
function Test-WordBridgePython { throw 'STORE ALIAS WAS EXECUTED' }
try { Find-WordBridgePython } catch { Write-Output $_.Exception.Message; exit 1 }
"""
        result = self.run_ps(script)
        self.assertEqual(result.returncode, 1)
        self.assertIn('No working', result.stdout)
        self.assertNotIn('STORE ALIAS WAS EXECUTED', result.stdout)


if __name__ == '__main__':
    unittest.main()
