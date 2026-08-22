"""Unit tests for the non-GUI logic of ConfigEditor.py.

The editor is instantiated without running __init__ (which builds the
Tk interface); only the attributes needed by each method under test are
set up, so the tests run headless.
"""

import importlib.util
import json
import sys
from pathlib import Path
from unittest import mock

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]


class _StubVar:
    """Minimal stand-in for tkinter variables (get/set)."""

    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


@pytest.fixture(scope="module")
def config_editor_module():
    spec = importlib.util.spec_from_file_location(
        "config_editor_under_test", REPO_ROOT / "ConfigEditor.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture
def editor(config_editor_module, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    instance = config_editor_module.ConfigEditor.__new__(
        config_editor_module.ConfigEditor
    )
    instance.version_configs = {
        "5.0": {"config_file": "OfficebackupSingleConfig.json"},
        "6.3": {"config_file": "OBU6.3.json"},
        "6.3Core": {"config_file": "OBU6.3Core.json"},
    }
    instance.current_version = "6.3"
    instance.config_data = {}
    instance.original_config = {}
    instance.history = []
    instance.history_index = -1
    instance.key_name_mode = "simple"
    instance.key_name_button = None
    instance.status_var = _StubVar()
    instance.update_config_ui = mock.Mock()
    return instance


class TestGetDefaultConfig:
    def test_63_defaults_include_openlist_params(self, editor):
        editor.current_version = "6.3"
        defaults = editor.get_default_config()
        assert defaults["openlist_upload_mode"] == "standard"
        assert defaults["upload_to_openlist_enable"] is True
        assert defaults["interval"] == 60

    def test_50_defaults_include_123pan_params(self, editor):
        editor.current_version = "5.0"
        defaults = editor.get_default_config()
        assert "client_id" in defaults
        assert "upload_to_123pan_enable" in defaults
        assert "openlist_url" not in defaults

    def test_63core_defaults_have_no_cloud_params(self, editor):
        editor.current_version = "6.3Core"
        defaults = editor.get_default_config()
        assert "openlist_url" not in defaults
        assert "client_id" not in defaults
        assert defaults["show_console_window_at_startup"] is True


class TestKeyDisplayName:
    def test_simple_mode_maps_to_chinese(self, editor):
        editor.key_name_mode = "simple"
        assert editor.get_key_display_name("interval") == "轮询间隔(秒)"

    def test_original_mode_returns_key(self, editor):
        editor.key_name_mode = "original"
        assert editor.get_key_display_name("interval") == "interval"

    def test_unknown_key_returns_key(self, editor):
        editor.key_name_mode = "simple"
        assert editor.get_key_display_name("mystery_key") == "mystery_key"


class TestLoadConfig:
    def test_creates_default_config_when_file_missing(self, editor, tmp_path):
        editor.load_config()
        saved = json.loads((tmp_path / "OBU6.3.json").read_text(encoding="utf-8"))
        assert saved == editor.get_default_config()
        assert editor.config_data == editor.get_default_config()
        assert editor.history_index == 0
        editor.update_config_ui.assert_called_once()

    def test_fills_missing_keys_and_filters_stale_keys(self, editor, tmp_path):
        (tmp_path / "OBU6.3.json").write_text(
            json.dumps({"interval": 30, "legacy_key_from_old_version": 1}),
            encoding="utf-8",
        )
        editor.load_config()
        saved = json.loads((tmp_path / "OBU6.3.json").read_text(encoding="utf-8"))
        assert saved["interval"] == 30
        assert "legacy_key_from_old_version" not in saved
        assert set(saved) == set(editor.get_default_config())
        assert "过滤" in editor.status_var.get()


class TestSaveConfig:
    def test_save_config_to_file_writes_json(self, editor, tmp_path):
        editor.save_config_to_file({"a": 1, "路径": "C:\\x"}, str(tmp_path / "out.json"))
        assert json.loads((tmp_path / "out.json").read_text(encoding="utf-8")) == {
            "a": 1,
            "路径": "C:\\x",
        }

    def test_save_config_filters_unknown_keys(self, editor, tmp_path):
        editor.config_data = dict(editor.get_default_config(), stale_key="x")
        editor.save_config()
        saved = json.loads((tmp_path / "OBU6.3.json").read_text(encoding="utf-8"))
        assert "stale_key" not in saved
        assert "stale_key" not in editor.config_data


class TestHistory:
    def test_on_config_change_appends_history_and_saves(self, editor, tmp_path):
        editor.config_data = dict(editor.get_default_config())
        editor.history = [dict(editor.config_data)]
        editor.history_index = 0

        editor.on_config_change("interval", 120)

        assert editor.config_data["interval"] == 120
        assert editor.history_index == 1
        saved = json.loads((tmp_path / "OBU6.3.json").read_text(encoding="utf-8"))
        assert saved["interval"] == 120

    def test_on_config_change_ignores_unknown_key(self, editor):
        editor.config_data = dict(editor.get_default_config())
        editor.history = [dict(editor.config_data)]
        editor.history_index = 0

        editor.on_config_change("not_a_real_key", "value")

        assert "not_a_real_key" not in editor.config_data
        assert editor.history_index == 0

    def test_undo_and_redo_restore_states(self, editor):
        first = dict(editor.get_default_config())
        second = dict(first, interval=99)
        editor.config_data = dict(second)
        editor.history = [dict(first), dict(second)]
        editor.history_index = 1

        editor.undo()
        assert editor.config_data["interval"] == first["interval"]
        editor.redo()
        assert editor.config_data["interval"] == 99

    def test_undo_at_start_is_noop(self, editor):
        editor.config_data = dict(editor.get_default_config())
        editor.history = [dict(editor.config_data)]
        editor.history_index = 0
        editor.undo()
        assert editor.history_index == 0


class TestAutoCompleteTargetPath:
    def test_appends_backup_suffix(self, editor):
        # forward-slash paths normalize identically on Windows and POSIX
        editor.config_data = {"accurate_backup_source_path": "/data/project"}
        var = _StubVar("/backups")
        editor.auto_complete_target_path(var)
        assert var.get() == "\\backups\\project-backup"

    def test_requires_source_path(self, editor, config_editor_module, monkeypatch):
        editor.config_data = {"accurate_backup_source_path": ""}
        info = mock.Mock()
        monkeypatch.setattr(config_editor_module, "messagebox", mock.Mock(showinfo=info))
        var = _StubVar("D:\\backups")
        editor.auto_complete_target_path(var)
        info.assert_called_once()
        assert var.get() == "D:\\backups"
