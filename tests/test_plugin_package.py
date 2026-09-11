import json
from pathlib import Path
import sys
import tempfile
import unittest
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import build_plugin


class PluginPackageTests(unittest.TestCase):
    def test_versions_match(self):
        self.assertEqual(build_plugin.package_version(build_plugin.ROOT), "0.3.0-dev")

    def test_package_contains_exact_allowlist_and_skill(self):
        with tempfile.TemporaryDirectory() as directory:
            target = build_plugin.build(output_dir=directory)
            with ZipFile(target) as archive:
                self.assertEqual(set(archive.namelist()),
                                 {f"wordbridge-plugin/{p}" for p in build_plugin.FILES})
                self.assertEqual(archive.read("wordbridge-plugin/skills/wordbridge/SKILL.md"),
                                 (build_plugin.ROOT / "skills/wordbridge/SKILL.md").read_bytes())
            with self.assertRaises(FileExistsError):
                build_plugin.build(output_dir=directory)

    def test_no_personal_or_runtime_members(self):
        for member in build_plugin.FILES:
            self.assertFalse(set(Path(member).parts) & {"local", ".venv", ".git", "__pycache__"})
            self.assertFalse(member.endswith(".docx"))

    def test_mcp_paths_are_package_relative(self):
        config = json.loads((build_plugin.ROOT / ".mcp.json").read_text(encoding="utf-8"))
        server = config["mcpServers"]["wordbridge"]
        self.assertEqual(server["command"], "${CLAUDE_PLUGIN_ROOT}/.venv/Scripts/python.exe")
        self.assertEqual(server["args"], ["${CLAUDE_PLUGIN_ROOT}/scripts/mcp_server.py"])
