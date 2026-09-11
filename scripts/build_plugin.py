"""Build a source-only plugin archive. Never installs, downloads, or calls Word."""

import json
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
# Explicit allowlist: new files require a packaging review.
FILES = (
    ".codex-plugin/plugin.json", ".mcp.json", "README.md", "CHANGELOG.md",
    "AGENTS.md", "requirements.txt", "skills/wordbridge/SKILL.md",
    "docs/插件打包与更新.md", "docs/第三方依赖.md", "docs/MCP确认表单兼容性问题.md",
    "test-documents/README.md", "install.ps1", "scripts/install_plugin.py",
    "scripts/build_plugin.py", "scripts/check_plugin_runtime.py",
    "scripts/append_text.py", "scripts/check_active_document.py",
    "scripts/check_mcp.py", "scripts/check_word_connection.py",
    "scripts/connection_popup.py", "scripts/connection_status.py",
    "scripts/mcp_server.py", "scripts/read_selection.py", "scripts/status_window.py",
)


def package_version(root):
    manifest = json.loads((root / FILES[0]).read_text(encoding="utf-8"))
    version = manifest["version"]
    if manifest["name"] != "wordbridge-plugin":
        raise ValueError("Unexpected plugin name")
    if not re.fullmatch(r"\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?", version):
        raise ValueError("Invalid package version")
    source = (root / "scripts/mcp_server.py").read_text(encoding="utf-8")
    if f'version="{version}"' not in source:
        raise ValueError("Plugin and MCP versions differ")
    return version


def build(root=ROOT, output_dir=None):
    root = Path(root).resolve()
    version = package_version(root)
    payload = {}
    for relative in FILES:
        source = (root / relative).resolve(strict=True)
        if not source.is_relative_to(root) or not source.is_file():
            raise ValueError(f"Unsafe package member: {relative}")
        payload[relative] = source.read_bytes()
    output_dir = Path(output_dir) if output_dir else root / "dist"
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / f"wordbridge-plugin-{version}-plugin.zip"
    # Never silently overwrite an earlier artifact, including an existing release.
    with ZipFile(target, "x", compression=ZIP_DEFLATED) as archive:
        for relative, data in payload.items():
            archive.writestr(f"wordbridge-plugin/{relative}", data)
    return target


if __name__ == "__main__":
    print(build())
