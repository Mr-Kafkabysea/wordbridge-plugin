"""Install or upgrade the local Windows Codex plugin without accessing Word."""
import argparse
import copy
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tomllib
from uuid import uuid4

from build_plugin import FILES, ROOT, package_version

NAME = 'wordbridge-plugin'
ALIASES = {NAME, 'wordbridge-mcp', 'wordbridge'}


def version_key(version):
    """SemVer precedence for the release/prerelease versions supported by our package."""
    match = re.fullmatch(r'(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-([0-9A-Za-z.-]+))?', version)
    if not match:
        raise ValueError(f'Invalid version: {version}')
    suffix = match[4]
    parts = []
    if suffix is not None:
        for part in suffix.split('.'):
            if not part or (part.isdigit() and len(part) > 1 and part[0] == '0'):
                raise ValueError(f'Invalid prerelease: {version}')
            parts.append((0, int(part)) if part.isdigit() else (1, part))
    return tuple(int(match[i]) for i in (1, 2, 3)), suffix is None, tuple(parts)


def read_catalog(path):
    if not path.exists():
        return {'name': 'personal', 'interface': {'displayName': 'Personal'}, 'plugins': []}
    data = json.loads(path.read_text(encoding='utf-8-sig'))
    if not isinstance(data, dict) or not isinstance(data.get('name'), str) or not re.fullmatch(r'[A-Za-z0-9_-]+', data['name']):
        raise ValueError('Invalid personal marketplace name')
    if not isinstance(data.get('plugins'), list) or not all(isinstance(p, dict) for p in data['plugins']):
        raise ValueError('Invalid personal marketplace plugins')
    return data


def add_entry(catalog):
    result = copy.deepcopy(catalog)
    if any(p.get('name') in ALIASES for p in result['plugins']):
        raise ValueError('WordBridge already listed; inspect the existing installation first.')
    result['plugins'].append({
        'name': NAME, 'source': {'source': 'local', 'path': './plugins/' + NAME},
        'policy': {'installation': 'AVAILABLE', 'authentication': 'ON_INSTALL'},
        'category': 'Productivity',
    })
    return result


def guard_path(path, profile):
    """Reject linked/reparse installation paths, including links inside the profile."""
    path = Path(path)
    if not path.is_absolute() or not path.is_relative_to(profile):
        raise ValueError(f'Installation path outside user profile: {path}')
    for part in [path, *path.parents]:
        if part == profile:
            break
        if part.is_symlink() or part.is_junction():
            raise ValueError(f'Linked installation path refused: {part}')
    if not path.resolve().is_relative_to(profile):
        raise ValueError(f'Installation path escapes user profile: {path}')
    return path


def check_legacy(profile, workspace, allowed_plugin=None):
    codex_root = Path(os.environ.get('CODEX_HOME', str(profile / '.codex')))
    if codex_root.resolve() != (profile / '.codex').resolve():
        raise ValueError('Custom CODEX_HOME is not supported.')
    for path in [profile / '.codex/skills/wordbridge', profile / '.agents/skills/wordbridge']:
        if path.exists():
            raise ValueError(f'Existing standalone WordBridge skill: {path}')
    configs = {profile / '.codex/config.toml'}
    configs.update(p / '.codex/config.toml' for p in [workspace, *workspace.parents])
    for path in configs:
        if not path.is_file():
            continue
        data = tomllib.loads(path.read_text(encoding='utf-8-sig'))
        if any('wordbridge' in k.lower() for k in data.get('mcp_servers', {})):
            raise ValueError(f'Existing standalone MCP configuration: {path}')
        for key in data.get('plugins', {}):
            if 'wordbridge' in key.lower() and (path != profile / '.codex/config.toml' or key != allowed_plugin):
                raise ValueError(f'Conflicting WordBridge plugin configuration: {path}')


def codex_prefix(path):
    path = Path(path).resolve(strict=True)
    if path.suffix.lower() == '.ps1':
        powershell = shutil.which('powershell.exe')
        if not powershell:
            raise ValueError('Windows PowerShell is required for this Codex launcher')
        return [powershell, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(path)]
    if path.suffix.lower() != '.exe':
        raise ValueError('Use a Codex .exe or .ps1 launcher')
    return [str(path)]


def run(command, cwd, timeout=600, capture=False):
    result = subprocess.run(command, cwd=cwd, check=True, timeout=timeout,
                            capture_output=capture, text=True, encoding='utf-8')
    return result.stdout if capture else None


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def replace_bytes(path, data):
    temp = path.with_name(path.name + '.' + uuid4().hex + '.tmp')
    temp.write_bytes(data)
    os.replace(temp, path)


def read_optional(path):
    return path.read_bytes() if path.exists() else None


def registered_plugin(prefix, profile):
    data = json.loads(run(prefix + ['plugin', 'list', '--json'], profile, 30, capture=True))
    if not isinstance(data, dict) or not isinstance(data.get('installed'), list):
        raise ValueError('Unexpected Codex plugin list response')
    matches = [p for p in data['installed'] if p.get('name') in ALIASES]
    if len(matches) > 1:
        raise ValueError('Multiple WordBridge registrations; resolve them before updating.')
    return matches[0] if matches else None


def verify_registration(prefix, profile, catalog, package, version):
    plugin = registered_plugin(prefix, profile)
    expected_id = NAME + '@' + catalog['name']
    if (not plugin or plugin.get('pluginId') != expected_id or
            plugin.get('version') != version or plugin.get('installed') is not True or
            plugin.get('enabled') is not True or
            plugin.get('source', {}).get('source') != 'local' or
            Path(plugin.get('source', {}).get('path', '')).resolve() != package.resolve()):
        raise ValueError('Codex registration/version/source mismatch; client reload is not verified.')
    cache = guard_path(profile / '.codex/plugins/cache' / catalog['name'] / NAME / version, profile)
    for relative in ('.codex-plugin/plugin.json', '.mcp.json', 'skills/wordbridge/SKILL.md'):
        cached = guard_path(cache / relative, profile)
        if not cached.is_file() or cached.read_bytes() != (package / relative).read_bytes():
            raise ValueError(f'Codex cache differs from the installed package: {relative}')
    return plugin


def inspect_install(profile, catalog, package):
    entries = [p for p in catalog['plugins'] if p.get('name') in ALIASES]
    legacy = profile / '.wordbridge/runtime' / NAME
    if not entries and not package.exists() and not legacy.exists():
        return None
    if len(entries) != 1 or entries[0].get('name') != NAME:
        raise ValueError('Incomplete or conflicting WordBridge installation.')
    if entries[0].get('source') != {'source': 'local', 'path': './plugins/' + NAME}:
        raise ValueError('WordBridge marketplace must point at the managed local package.')
    if not package.is_dir():
        raise ValueError('Installed package is missing.')
    for relative in ('.mcp.json', '.codex-plugin/plugin.json', 'skills/wordbridge/SKILL.md'):
        guard_path(package / relative, profile)
    config_bytes = (package / '.mcp.json').read_bytes()
    config = json.loads(config_bytes)['mcpServers']['wordbridge']
    args = config.get('args')
    if not isinstance(args, list) or len(args) != 1:
        raise ValueError('Unrecognized installed MCP arguments.')
    script = guard_path(Path(args[0]), profile)
    runtime = script.parent.parent
    if (script.name != 'mcp_server.py' or script.parent.name != 'scripts' or
            runtime.parent != profile / '.wordbridge/runtime' or
            not (runtime.name == NAME or runtime.name.startswith(NAME + '-'))):
        raise ValueError('Unrecognized managed runtime location.')
    if config.get('command') != str(runtime / '.venv/Scripts/python.exe'):
        raise ValueError('Installed Python does not match the runtime.')
    for relative in ('install-receipt.json', '.mcp.json', '.venv/Scripts/python.exe',
                     '.codex-plugin/plugin.json', 'scripts/mcp_server.py', 'skills/wordbridge/SKILL.md'):
        if not guard_path(runtime / relative, profile).is_file():
            raise ValueError(f'Installed runtime is incomplete: {relative}')
    receipt_bytes = (runtime / 'install-receipt.json').read_bytes()
    receipt = json.loads(receipt_bytes)
    version = package_version(runtime)
    version_key(version)
    manifest = json.loads((package / '.codex-plugin/plugin.json').read_bytes())
    if (receipt.get('status') != 'cli_install_completed' or receipt.get('version') != version or
            receipt.get('runtime') != str(runtime) or manifest.get('name') != NAME or
            manifest.get('version') != version or (runtime / '.mcp.json').read_bytes() != config_bytes or
            (runtime / 'skills/wordbridge/SKILL.md').read_bytes() != (package / 'skills/wordbridge/SKILL.md').read_bytes()):
        raise ValueError('Installed receipt, MCP, Skill or manifest is inconsistent.')
    return {'version': version, 'runtime': str(runtime), 'receipt': receipt_bytes,
            'config': config_bytes, 'manifest': (package / '.codex-plugin/plugin.json').read_bytes(),
            'skill': (package / 'skills/wordbridge/SKILL.md').read_bytes()}


def prepare_catalog_package(target, package, payload):
    package.parent.mkdir(parents=True, exist_ok=True)
    package.mkdir()
    for relative, data in payload.items():
        path = package / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
    (package / '.mcp.json').write_bytes((target / '.mcp.json').read_bytes())


def prepare_runtime(target, payload, version):
    target.mkdir()
    for relative, data in payload.items():
        path = target / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
    if package_version(target) != version:
        raise ValueError('Source version changed while preparing the snapshot.')
    run([sys.executable, '-m', 'venv', str(target / '.venv')], target)
    python = target / '.venv/Scripts/python.exe'
    run([str(python), '-m', 'pip', 'install', '--disable-pip-version-check',
         '-r', str(target / 'requirements.txt')], target)
    run([str(python), '-m', 'pip', 'check'], target)
    run([str(python), str(target / 'scripts/check_plugin_runtime.py')], target, 60)
    write_json(target / '.mcp.json', {'mcpServers': {'wordbridge': {
        'command': str(python), 'args': [str(target / 'scripts/mcp_server.py')],
        'env': {'WINDIR': os.environ['WINDIR']},
    }}})
    write_json(target / 'install-receipt.json', {
        'version': version, 'runtime': str(target), 'status': 'prepared',
    })


def install(source, profile, codex, check_only=False):
    source, profile = Path(source).resolve(), Path(profile).resolve()
    version = package_version(source)
    candidate_key = version_key(version)
    catalog_path = guard_path(profile / '.agents/plugins/marketplace.json', profile)
    package = guard_path(profile / 'plugins' / NAME, profile)
    runtime_root = guard_path(profile / '.wordbridge/runtime', profile)
    journal = guard_path(runtime_root / '.wordbridge-transaction.json', profile)
    lock = guard_path(runtime_root / '.wordbridge-install.lock', profile)
    if journal.exists() or lock.exists():
        raise ValueError(f'An update is running or interrupted. Inspect {journal} and {lock}; no changes made.')
    original = read_optional(catalog_path)
    catalog = read_catalog(catalog_path)
    old = inspect_install(profile, catalog, package)
    selector = NAME + '@' + catalog['name']
    check_legacy(profile, Path.cwd().resolve(), selector if old else None)
    prefix = codex_prefix(codex)
    if old:
        verify_registration(prefix, profile, catalog, package, old['version'])
    elif registered_plugin(prefix, profile):
        raise ValueError('WordBridge is registered without a recognized managed installation.')
    print(f"Installed: {old['version'] if old else '(none)'} -> Package: {version}", flush=True)
    if old and candidate_key < version_key(old['version']):
        raise ValueError('Downgrade refused. Use a newer release package.')
    if old and candidate_key == version_key(old['version']):
        print('Already up to date; no files, dependencies or Codex settings changed.')
        return 'unchanged'
    action = 'upgrade' if old else 'install'
    # Validate the complete snapshot before creating any files.
    payload = {}
    for relative in FILES:
        path = (source / relative).resolve(strict=True)
        if not path.is_relative_to(source) or not path.is_file():
            raise ValueError(f'Unsafe package member: {relative}')
        payload[relative] = path.read_bytes()
    run(prefix + ['plugin', 'add', '--help'], source, 30, capture=True)
    print(f'Action: {action}. MCP and Skill will be installed together.', flush=True)
    if check_only:
        print('Preflight passed. No files changed; no dependencies installed.')
        return action
    token = uuid4().hex
    target = guard_path(runtime_root / f'{NAME}-{version}-{token}', profile)
    stage = guard_path(package.with_name('.wordbridge-stage-' + token), profile)
    backup = guard_path(package.with_name('.wordbridge-backup-' + token), profile)
    failed = guard_path(package.with_name('.wordbridge-failed-' + token), profile)
    runtime_root.mkdir(parents=True, exist_ok=True)
    with lock.open('x', encoding='utf-8') as stream:
        stream.write(str(os.getpid()))
    try:
        # Lock acquired: recheck the state seen during preflight.
        if read_optional(catalog_path) != original or inspect_install(profile, catalog, package) != old:
            raise ValueError('Installation changed during preflight; retry after inspection.')
        prepare_runtime(target, payload, version)
        prepare_catalog_package(target, stage, payload)
        if read_optional(catalog_path) != original or inspect_install(profile, catalog, package) != old:
            raise ValueError('Installation changed during preparation; existing installation not overwritten.')
        if old:
            verify_registration(prefix, profile, catalog, package, old['version'])
        publish(profile, catalog_path, catalog, original, package, stage, backup, failed,
                target, journal, prefix, old, version)
    except Exception:
        print(f'Operation failed. New files retained at {target}; inspect any transaction journal before retrying.',
              file=sys.stderr)
        raise
    finally:
        lock.unlink()
    print(f'{action.capitalize()} completed: {version}. Open a NEW Codex task to load MCP and Skill; '
          'existing tasks may still use the old service. Verify word_status there.')
    return action


def publish(profile, catalog_path, catalog, original, package, stage, backup, failed,
            target, journal, prefix, old, version):
    """Keep a recovery record across the non-atomic filesystem/CLI boundary."""
    selector = NAME + '@' + catalog['name']
    write_json(journal, {'version': version, 'previous_version': old['version'] if old else None,
                         'runtime': str(target), 'package': str(package), 'backup': str(backup),
                         'staged_package': str(stage), 'failed_package': str(failed),
                         'restore_command': ['codex', 'plugin', 'add', selector]})
    switched = False
    cli_started = False
    published_catalog = None
    catalog_backup = catalog_path.with_name('marketplace.before-wordbridge-' + uuid4().hex + '.json')
    try:
        if old:
            package.rename(backup)
        stage.rename(package)
        switched = True
        if not old:
            catalog_path.parent.mkdir(parents=True, exist_ok=True)
            if original is not None:
                catalog_backup.write_bytes(original)
            updated = add_entry(catalog)
            published_catalog = (json.dumps(updated, ensure_ascii=False, indent=2) + '\n').encode('utf-8')
            replace_bytes(catalog_path, published_catalog)
        cli_started = True
        run(prefix + ['plugin', 'add', selector], profile, 120)
        verify_registration(prefix, profile, catalog, package, version)
        write_json(target / 'install-receipt.json', {
            'version': version, 'runtime': str(target), 'status': 'cli_install_completed',
            'previous_runtime': old['runtime'] if old else None,
            'previous_package': str(backup) if old else None,
        })
        journal.unlink()
    except Exception as error:
        try:
            if old:
                # Never touch the previous runtime or move its virtualenv.
                if switched:
                    package.rename(failed)
                if backup.exists():
                    backup.rename(package)
                if cli_started:
                    run(prefix + ['plugin', 'add', selector], profile, 120)
                    verify_registration(prefix, profile, catalog, package, old['version'])
                if inspect_install(profile, catalog, package) != old:
                    raise ValueError('Previous installation could not be verified.')
                journal.unlink()
                print('Previous package and CLI registration restored. Existing tasks require a reload.',
                      file=sys.stderr)
            elif not cli_started:
                if switched:
                    package.rename(failed)
                if published_catalog is not None:
                    if read_optional(catalog_path) != published_catalog:
                        raise ValueError('Marketplace changed externally; not restored automatically.')
                    if original is None:
                        catalog_path.unlink()
                    else:
                        replace_bytes(catalog_path, original)
                journal.unlink()
            else:
                raise ValueError('First-install CLI outcome uncertain; inspect registration before recovery.')
        except Exception as recovery_error:
            raise RuntimeError(f'Update failed: {error}. Recovery incomplete: {recovery_error}. '
                               f'Preserved recovery record: {journal}') from error
        raise


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--codex', required=True)
    parser.add_argument('--check', action='store_true', help='Show install/upgrade/no-op without changes')
    args = parser.parse_args()
    if sys.platform != 'win32' or sys.version_info[:2] != (3, 14) or sys.maxsize <= 2**32:
        parser.error('Requires Windows and 64-bit Python 3.14')
    if not os.environ.get('WINDIR'):
        parser.error('WINDIR is missing')
    try:
        install(ROOT, Path.home(), args.codex, args.check)
    except (OSError, ValueError, KeyError, TypeError, RuntimeError, subprocess.SubprocessError) as error:
        print(f'Install/update failed: {error}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
