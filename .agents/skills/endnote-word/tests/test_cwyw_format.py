import json
import os
from pathlib import Path
import subprocess
import unittest

ROOT = Path(__file__).resolve().parents[1]
BRIDGE = Path(os.environ.get('ENDNOTE_TEST_BRIDGE', ROOT / 'scripts' / 'edit-citation.ps1'))

class CwywFormatRetryTests(unittest.TestCase):
    def run_format(self, mode):
        bridge = BRIDGE.read_text(encoding='utf-8-sig')
        functions = bridge.split('function Test-PendingCwywCitation {', 1)[1].split('\nfunction Get-CitationXml {', 1)[0]
        code = r'''
$ErrorActionPreference = 'Stop'
$word = [pscustomobject]@{ Runs = 0 }
$word | Add-Member ScriptMethod Run { param($name) $this.Runs++ }
$script:formatRuns = 0
$mode = '%s'
$paragraph = [pscustomobject]@{Text=''}
$targets = @([pscustomobject]@{temporary='{Kourentzes, 2019 #31}'; paragraph=$paragraph})
function Get-Cites {
    $result = '[31]'
    if ($mode -eq 'always' -or ($mode -eq 'first' -and $word.Runs -eq 1)) { $result = '{Kourentzes, 2019 #31}' }
    if ($mode -eq 'prefix' -and $word.Runs -eq 1) { $result = 'see {Kourentzes, 2019 #31}, p. 20' }
    if ($mode -eq 'merged' -and $word.Runs -eq 1) { $result = '{A, 2018 #30; Kourentzes, 2019 #31}' }
    if ($mode -eq 'plain' -and $word.Runs -eq 1) { $paragraph.Text = 'text {Kourentzes, 2019 #31}' } else { $paragraph.Text = 'text [31]' }
    [pscustomobject]@{Result=[pscustomobject]@{Text=$result}}
}
function Test-PendingCwywCitation {%s
try { Invoke-CwywFormat; $outcome = 'ok' } catch { $outcome = $_.Exception.Message }
[pscustomobject]@{outcome=$outcome; runs=$word.Runs; reported=$script:formatRuns} | ConvertTo-Json -Compress
''' % (mode, functions)
        proc = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', code], capture_output=True, text=True, check=True)
        return json.loads(proc.stdout)

    def test_formatted_field_needs_one_pass(self):
        self.assertEqual(self.run_format('done'), {'outcome':'ok', 'runs':1, 'reported':1})

    def test_temporary_results_and_inserted_plain_text_get_one_retry(self):
        for mode in ('first', 'prefix', 'merged', 'plain'):
            with self.subTest(mode=mode):
                self.assertEqual(self.run_format(mode), {'outcome':'ok', 'runs':2, 'reported':2})

    def test_persistent_temporary_field_fails_closed(self):
        result = self.run_format('always')
        self.assertIn('left temporary citations', result['outcome'])
        self.assertEqual((result['runs'], result['reported']), (2, 2))

if __name__ == '__main__':
    unittest.main()
