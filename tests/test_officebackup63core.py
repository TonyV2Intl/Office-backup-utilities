"""Unit tests for Officebackup6.3Core.py (loaded via the conftest harness)."""

import hashlib
import json
import os
import stat
from types import SimpleNamespace

SCRIPT = "Officebackup6.3Core.py"
CONFIG_FILE = "OBU6.3Core.json"


def base_config(**overrides):
    config = {"save_log": True}
    config.update(overrides)
    return config


class TestConfigLoading:
    def test_creates_default_config_when_missing(self, load_script, tmp_path):
        module = load_script(SCRIPT)
        saved = json.loads((tmp_path / CONFIG_FILE).read_text(encoding="utf-8"))
        assert set(saved) == set(module.default_config)
        assert saved["interval"] == 60
        assert saved["backup_timeout"] == 600

    def test_fills_missing_keys_and_writes_back(self, load_script, tmp_path):
        load_script(SCRIPT, config={"interval": 10})
        saved = json.loads((tmp_path / CONFIG_FILE).read_text(encoding="utf-8"))
        assert saved["interval"] == 10
        assert saved["ppt_backup_path"] == "C:\\Officebackup\\pptbackup"
        assert saved["save_log"] is True


class TestLogHandling:
    def test_previous_log_is_archived(self, load_script, tmp_path):
        (tmp_path / "OBUlatest.log").write_text("old session", encoding="utf-8")
        load_script(SCRIPT, config=base_config(archive_previous_log=True))
        assert (tmp_path / "OBUprevious.log").read_text(encoding="utf-8") == "old session"

    def test_log_print_increments_runid_and_writes_log(self, load_script, tmp_path, capsys):
        module = load_script(SCRIPT, config=base_config())
        start = module.runid
        module.log_print("core log line")
        assert module.runid == start + 1
        assert "core log line" in capsys.readouterr().out
        assert "core log line" in (tmp_path / "OBUlatest.log").read_text(encoding="utf-8")


class TestHelpers:
    def test_calculate_md5(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        target = tmp_path / "file.bin"
        target.write_bytes(b"core content")
        assert module.calculate_md5(str(target)) == hashlib.md5(b"core content").hexdigest()
        assert module.calculate_md5(str(tmp_path / "missing.bin")) is None

    def test_remove_readonly(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        target = tmp_path / "readonly.txt"
        target.write_text("data", encoding="utf-8")
        os.chmod(target, stat.S_IRUSR | stat.S_IRGRP | stat.S_IROTH)
        module.remove_readonly(str(target))
        assert os.stat(target).st_mode & 0o200

    def test_timeout_decorator_passes_through_results(self, load_script):
        module = load_script(SCRIPT, config=base_config())

        @module.timeout(seconds=600, config_key="backup_timeout")
        def add(a, b):
            return a + b

        assert add(4, 5) == 9

    def test_timeout_decorator_returns_none_on_exception(self, load_script, capsys):
        module = load_script(SCRIPT, config=base_config())

        @module.timeout(seconds=600)
        def boom():
            raise RuntimeError("expected failure")

        assert boom() is None
        assert "Error in boom: expected failure" in capsys.readouterr().out


class TestBackupOpenFiles:
    def test_backs_up_open_document(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        source = tmp_path / "notes.docx"
        source.write_bytes(b"word-bytes")
        backup_dir = tmp_path / "backups"
        module.win32.Dispatch.return_value = SimpleNamespace(
            Documents=[SimpleNamespace(FullName=str(source))]
        )

        module.save_open_word_files(str(backup_dir))

        assert (backup_dir / "notes.docx").read_bytes() == b"word-bytes"
        module.win32.Dispatch.assert_called_with("Word.Application")

    def test_unchanged_file_is_skipped_by_md5(self, load_script, tmp_path, capsys):
        module = load_script(SCRIPT, config=base_config())
        source = tmp_path / "notes.docx"
        source.write_bytes(b"word-bytes")
        backup_dir = tmp_path / "backups"
        module.win32.Dispatch.return_value = SimpleNamespace(
            Documents=[SimpleNamespace(FullName=str(source))]
        )

        module.save_open_word_files(str(backup_dir))
        capsys.readouterr()
        module.save_open_word_files(str(backup_dir))

        assert "skipped backup" in capsys.readouterr().out

    def test_com_error_logs_application_not_detected(self, load_script, tmp_path, capsys):
        module = load_script(SCRIPT, config=base_config())
        com_error = type("com_error", (Exception,), {})
        module.win32.GetObject.side_effect = com_error("no app")
        module.save_open_WPS_files(str(tmp_path / "backups"))
        assert "No WPS ppt available now (KWPP application not detected)" in capsys.readouterr().out


class TestAccurateBackup:
    def test_copies_tree_and_disables_itself(self, load_script, tmp_path):
        source = tmp_path / "source"
        source.mkdir()
        (source / "doc.txt").write_text("payload", encoding="utf-8")
        target = tmp_path / "target"
        module = load_script(
            SCRIPT,
            config=base_config(
                accurate_backup_enable=True,
                accurate_backup_source_path=str(source),
                accurate_backup_target_path=str(target),
            ),
        )

        module.accurate_backup()

        assert (target / "doc.txt").read_text(encoding="utf-8") == "payload"
        saved = json.loads((tmp_path / CONFIG_FILE).read_text(encoding="utf-8"))
        assert saved["accurate_backup_enable"] is False
        assert module.accurate_backup_running is False

    def test_empty_paths_disable_accurate_backup_for_session(self, load_script):
        module = load_script(
            SCRIPT,
            config=base_config(accurate_backup_enable=True, accurate_backup_source_path=""),
        )
        assert module.config["accurate_backup_enable"] is False
