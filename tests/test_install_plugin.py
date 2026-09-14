"""Installer integration tests: real temporary files, simulated pip and Codex."""
import contextlib
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import install_plugin as installer


class FakeCommands:
    def __init__(self):
        self.installed = None
        self.calls = []
        self.fail_pip = False
        self.fail_add = 0
        self.stale_add = False

    def __call__(self, command, cwd, timeout=600, capture=False):
        self.calls.append(command)
        if command[1:] == ['plugin', 'list', '--json']:
            return json.dumps({'installed': [self.installed] if self.installed else []})
        if command[1:3] == ['-m', 'venv']:
            python = Path(command[3]) / 'Scripts/python.exe'
            python.parent.mkdir(parents=True)
            python.write_bytes(b'test interpreter')
        if command[1:4] == ['-m', 'pip', 'install'] and self.fail_pip:
            raise subprocess.CalledProcessError(1, command)
        if command[1:3] == ['plugin', 'add'] and command[-1] != '--help':
            package = Path(cwd) / 'plugins' / installer.NAME
            version = json.loads((package / '.codex-plugin/plugin.json').read_bytes())['version']
            if not self.stale_add:
                self.installed = {'pluginId': command[-1], 'name': installer.NAME,
                                  'version': version, 'installed': True, 'enabled': True,
                                  'source': {'source': 'local', 'path': str(package)}}
            if not self.stale_add:
                cache = Path(cwd) / '.codex/plugins/cache' / command[-1].split('@')[1] / installer.NAME / version
                for relative in ('.codex-plugin/plugin.json', '.mcp.json', 'skills/wordbridge/SKILL.md'):
                    cached = cache / relative
                    cached.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copyfile(package / relative, cached)
            if self.fail_add:
                self.fail_add -= 1
                raise subprocess.CalledProcessError(1, command)
        return None


class InstallerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name).resolve()
        self.profile = self.base / 'profile'
        self.profile.mkdir()
        self.source = self.base / 'source'
        for relative in installer.FILES:
            path = self.source / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(installer.ROOT / relative, path)
        self.set_version('0.2.2-dev')
        self.enterContext(contextlib.chdir(self.source))
        self.fake = FakeCommands()
        self.addCleanup(patch.stopall)
        # Ancestor config scans must see fixture files, not this developer's profile.
        real_is_file = Path.is_file
        def fixture_is_file(path):
            if path.name == 'config.toml' and path.parent.name == '.codex' and not path.is_relative_to(self.base):
                return False
            return real_is_file(path)
        patch.object(Path, 'is_file', fixture_is_file).start()
        patch.object(installer, 'run', self.fake).start()
        patch.object(installer, 'codex_prefix', return_value=['codex.exe']).start()
        patch.dict(os.environ, {'WINDIR': 'C:\\Windows', 'CODEX_HOME': str(self.profile / '.codex')}).start()
        patch('sys.stdout', new=io.StringIO()).start()
        patch('sys.stderr', new=io.StringIO()).start()

    def set_version(self, version):
        path = self.source / '.codex-plugin/plugin.json'
        manifest = json.loads(path.read_bytes())
        previous = manifest['version']
        manifest['version'] = version
        installer.write_json(path, manifest)
        path = self.source / 'scripts/mcp_server.py'
        path.write_text(path.read_text(encoding='utf-8').replace(
            f'version="{previous}"', f'version="{version}"'), encoding='utf-8')

    def install(self, check=False):
        return installer.install(self.source, self.profile, 'codex.exe', check)

    @property
    def package(self):
        return self.profile / 'plugins' / installer.NAME

    @property
    def catalog(self):
        return self.profile / '.agents/plugins/marketplace.json'

    @property
    def journal(self):
        return self.profile / '.wordbridge/runtime/.wordbridge-transaction.json'

    def current_runtime(self):
        config = json.loads((self.package / '.mcp.json').read_bytes())
        return Path(config['mcpServers']['wordbridge']['args'][0]).parent.parent

    def snapshot(self):
        return {str(p.relative_to(self.profile)): p.read_bytes()
                for p in self.profile.rglob('*') if p.is_file()}

    def test_catalog_preserves_other_plugins(self):
        original = {'name': 'mine', 'interface': {'displayName': 'Mine'},
                    'plugins': [{'name': 'other', 'source': './plugins/other'}]}
        installer.write_json(self.catalog, original)
        self.install()
        updated = installer.read_catalog(self.catalog)
        self.assertEqual(updated['plugins'][0], original['plugins'][0])
        self.assertEqual(updated['interface'], original['interface'])
        self.assertEqual(updated['name'], 'mine')

    def test_duplicate_refused(self):
        with self.assertRaises(ValueError):
            installer.add_entry({'plugins': [{'name': installer.NAME}]})

    def test_bad_catalog_refused(self):
        installer.write_json(self.catalog, {'name': 'bad@name', 'plugins': []})
        with self.assertRaises(ValueError):
            self.install()

    def test_check_only_does_not_write(self):
        self.assertEqual(self.install(True), 'install')
        self.assertEqual(self.snapshot(), {})

    def test_install_separates_runtime_and_package(self):
        self.assertEqual(self.install(), 'install')
        runtime = self.current_runtime()
        self.assertFalse((self.package / '.venv').exists())
        self.assertEqual((self.package / 'skills/wordbridge/SKILL.md').read_bytes(),
                         (runtime / 'skills/wordbridge/SKILL.md').read_bytes())
        self.assertEqual(json.loads((runtime / 'install-receipt.json').read_bytes())['status'],
                         'cli_install_completed')
        self.assertFalse(self.journal.exists())
        self.assertFalse((runtime.parent / '.wordbridge-install.lock').exists())

    def test_same_version_is_noop(self):
        self.install()
        before = self.snapshot()
        self.fake.calls.clear()
        self.assertEqual(self.install(), 'unchanged')
        self.assertEqual(before, self.snapshot())
        self.assertEqual(self.fake.calls, [['codex.exe', 'plugin', 'list', '--json']])

    def test_upgrade_preserves_old_runtime_catalog_and_updates_skill(self):
        self.install()
        old = self.current_runtime()
        old_files = {str(p.relative_to(old)): p.read_bytes() for p in old.rglob('*') if p.is_file()}
        catalog = self.catalog.read_bytes()
        self.set_version('0.3.0-dev')
        (self.source / 'skills/wordbridge/SKILL.md').write_text('new skill', encoding='utf-8')
        self.assertEqual(self.install(), 'upgrade')
        self.assertNotEqual(self.current_runtime(), old)
        self.assertEqual(self.fake.installed['version'], '0.3.0-dev')
        self.assertEqual(self.catalog.read_bytes(), catalog)
        self.assertEqual((self.package / 'skills/wordbridge/SKILL.md').read_text(), 'new skill')
        self.assertEqual(old_files, {str(p.relative_to(old)): p.read_bytes()
                                    for p in old.rglob('*') if p.is_file()})
        before = self.snapshot()
        self.assertEqual(self.install(), 'unchanged')
        self.assertEqual(before, self.snapshot())

    def test_upgrade_legacy_fixed_runtime(self):
        self.install()
        old = self.current_runtime()
        legacy = old.parent / installer.NAME
        old.rename(legacy)  # Fixture only: no real virtualenv or running process.
        cache = self.profile / '.codex/plugins/cache/personal' / installer.NAME / '0.2.2-dev'
        for root in (legacy, self.package, cache):
            path = root / '.mcp.json'
            path.write_text(path.read_text().replace(str(old).replace('\\', '\\\\'),
                                                    str(legacy).replace('\\', '\\\\')))
        installer.write_json(legacy / 'install-receipt.json',
                             {'version': '0.2.2-dev', 'runtime': str(legacy),
                              'status': 'cli_install_completed'})
        self.set_version('0.3.0-dev')
        self.assertEqual(self.install(), 'upgrade')
        self.assertTrue(legacy.is_dir())

    def test_upgrade_check_only_and_downgrade_do_not_write(self):
        self.install()
        before = self.snapshot()
        self.set_version('0.3.0-dev')
        self.assertEqual(self.install(True), 'upgrade')
        self.assertEqual(before, self.snapshot())
        self.set_version('0.1.0')
        with self.assertRaisesRegex(ValueError, 'Downgrade'):
            self.install()
        self.assertEqual(before, self.snapshot())

    def test_dependency_failure_does_not_register(self):
        self.fake.fail_pip = True
        with self.assertRaises(subprocess.CalledProcessError):
            self.install()
        self.assertFalse(self.catalog.exists())
        self.assertFalse(self.package.exists())
        self.fake.fail_pip = False
        self.assertEqual(self.install(), 'install')

    def test_upgrade_dependency_failure_keeps_old(self):
        self.install()
        old = self.current_runtime()
        catalog = self.catalog.read_bytes()
        self.set_version('0.3.0-dev')
        self.fake.fail_pip = True
        with self.assertRaises(subprocess.CalledProcessError):
            self.install()
        self.assertEqual(self.current_runtime(), old)
        self.assertEqual(self.catalog.read_bytes(), catalog)
        self.assertEqual(self.fake.installed['version'], '0.2.2-dev')
        self.assertFalse(self.journal.exists())

    def test_cli_failure_after_mutation_restores_old(self):
        self.install()
        old = self.current_runtime()
        self.set_version('0.3.0-dev')
        self.fake.fail_add = 1
        with self.assertRaises(subprocess.CalledProcessError):
            self.install()
        self.assertEqual(self.current_runtime(), old)
        self.assertEqual(self.fake.installed['version'], '0.2.2-dev')
        self.assertFalse(self.journal.exists())
        self.assertEqual(self.install(), 'upgrade')

    def test_stale_cli_version_detected_and_restored(self):
        self.install()
        old = self.current_runtime()
        self.set_version('0.3.0-dev')
        self.fake.stale_add = True
        with self.assertRaisesRegex(ValueError, 'registration'):
            self.install()
        self.assertEqual(self.current_runtime(), old)
        self.assertFalse(self.journal.exists())

    def test_rollback_failure_leaves_journal_and_blocks_retry(self):
        self.install()
        self.set_version('0.3.0-dev')
        self.fake.fail_add = 2
        with self.assertRaisesRegex(RuntimeError, 'Recovery incomplete'):
            self.install()
        self.assertTrue(self.journal.exists())
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, 'interrupted'):
            self.install()
        self.assertEqual(before, self.snapshot())

    def test_first_install_cli_failure_is_not_claimed_rolled_back(self):
        self.fake.fail_add = 1
        with self.assertRaisesRegex(RuntimeError, 'outcome uncertain'):
            self.install()
        self.assertTrue(self.journal.exists())

    def test_receipt_corruption_refused(self):
        self.install()
        (self.current_runtime() / 'install-receipt.json').write_text('{}')
        with self.assertRaisesRegex(ValueError, 'inconsistent'):
            self.install()

    def test_existing_registered_config_allowed_but_standalone_refused(self):
        self.install()
        config = self.profile / '.codex/config.toml'
        config.parent.mkdir(parents=True, exist_ok=True)
        config.write_text('[plugins."wordbridge-plugin@personal"]\nenabled = true\n')
        self.assertEqual(self.install(), 'unchanged')
        config.write_text('[mcp_servers.wordbridge]\ncommand = "old"\n')
        with self.assertRaisesRegex(ValueError, 'standalone'):
            self.install()

    def test_disabled_plugin_refused_without_enabling(self):
        self.install()
        self.fake.installed['enabled'] = False
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, 'registration'):
            self.install()
        self.assertEqual(before, self.snapshot())

    def test_other_marketplace_and_foreign_runtime_refused(self):
        self.install()
        data = installer.read_catalog(self.catalog)
        data['plugins'][0]['source']['path'] = './elsewhere'
        installer.write_json(self.catalog, data)
        with self.assertRaisesRegex(ValueError, 'managed local'):
            self.install()

    def test_active_lock_refused(self):
        path = self.profile / '.wordbridge/runtime/.wordbridge-install.lock'
        path.parent.mkdir(parents=True)
        path.write_text('123')
        with self.assertRaisesRegex(ValueError, 'running or interrupted'):
            self.install()

    def test_junction_inside_profile_refused(self):
        target = self.profile / 'plugins'
        target.mkdir()
        with patch.object(Path, 'is_junction', lambda p: p == target):
            with self.assertRaisesRegex(ValueError, 'Linked'):
                self.install()

    def test_semver_order_and_validation(self):
        values = ['0.2.2-dev', '0.2.2', '0.3.0-dev.2', '0.3.0-dev.10', '0.3.0', '0.10.0', '1.0.0']
        self.assertEqual(sorted(reversed(values), key=installer.version_key), values)
        for value in ('0.01.0', '0.3.0-dev..1', '0.3.0-dev.01', '../bad', '1.2'):
            with self.assertRaises(ValueError):
                installer.version_key(value)

    def test_cache_mismatch_refused_even_at_same_version(self):
        self.install()
        cache = self.profile / '.codex/plugins/cache/personal' / installer.NAME / '0.2.2-dev'
        (cache / 'skills/wordbridge/SKILL.md').write_text('stale skill')
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, 'cache differs'):
            self.install()
        self.assertEqual(before, self.snapshot())

    def test_package_switch_failure_restores_old_without_cli(self):
        self.install()
        old = self.current_runtime()
        self.set_version('0.3.0-dev')
        rename = Path.rename
        def fail_stage(path, target):
            if path.name.startswith('.wordbridge-stage-'):
                raise OSError('simulated package in use')
            return rename(path, target)
        with patch.object(Path, 'rename', fail_stage):
            with self.assertRaisesRegex(OSError, 'in use'):
                self.install()
        self.assertEqual(self.current_runtime(), old)
        self.assertFalse(self.journal.exists())
        self.assertEqual(self.fake.installed['version'], '0.2.2-dev')

    def test_marketplace_changed_during_prepare_is_preserved(self):
        self.install()
        old = self.current_runtime()
        self.set_version('0.3.0-dev')
        prepare = installer.prepare_runtime
        def concurrent_edit(*args):
            prepare(*args)
            catalog = installer.read_catalog(self.catalog)
            catalog['plugins'].append({'name': 'added-by-user'})
            installer.write_json(self.catalog, catalog)
        with patch.object(installer, 'prepare_runtime', concurrent_edit):
            with self.assertRaisesRegex(ValueError, 'changed during preparation'):
                self.install()
        self.assertEqual(self.current_runtime(), old)
        self.assertEqual(installer.read_catalog(self.catalog)['plugins'][-1]['name'], 'added-by-user')

    def test_handshake_failure_never_switches(self):
        self.install()
        old = self.current_runtime()
        self.set_version('0.3.0-dev')
        def fail_handshake(command, *args, **kwargs):
            if str(command[-1]).endswith('check_plugin_runtime.py'):
                raise subprocess.CalledProcessError(1, command)
            return self.fake(command, *args, **kwargs)
        with patch.object(installer, 'run', fail_handshake):
            with self.assertRaises(subprocess.CalledProcessError):
                self.install()
        self.assertEqual(self.current_runtime(), old)
        self.assertEqual(self.fake.installed['version'], '0.2.2-dev')

    def test_source_version_mismatch_refused(self):
        path = self.source / '.codex-plugin/plugin.json'
        data = json.loads(path.read_bytes())
        data['version'] = '0.3.0'
        installer.write_json(path, data)
        with self.assertRaisesRegex(ValueError, 'versions differ'):
            self.install()
        self.assertEqual(self.snapshot(), {})


if __name__ == '__main__':
    unittest.main()
