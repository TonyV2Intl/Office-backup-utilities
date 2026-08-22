"""Unit tests for Officebackup6.3.py (loaded via the conftest harness)."""

import asyncio
import hashlib
import json
import os
import stat
import sys
import types
from types import SimpleNamespace
from unittest import mock

import pytest

SCRIPT = "Officebackup6.3.py"
CONFIG_FILE = "OBU6.3.json"


def base_config(**overrides):
    config = {
        "upload_to_openlist_enable": False,
        "hide_tray_icon": True,
        "save_log": True,
    }
    config.update(overrides)
    return config


class TestConfigLoading:
    def test_creates_default_config_when_missing(self, load_script, tmp_path):
        module = load_script(SCRIPT)
        saved = json.loads((tmp_path / CONFIG_FILE).read_text(encoding="utf-8"))
        # note: config aliases default_config here, which the session then
        # mutates (e.g. force-disabling upload), so compare against the file
        assert set(saved) == set(module.default_config)
        assert saved["interval"] == 60
        assert saved["upload_to_openlist_enable"] is True
        assert module.sleeptime == 60
        assert module.ppt_save_folder == "C:\\Officebackup\\pptbackup"

    def test_fills_missing_keys_and_writes_back(self, load_script, tmp_path):
        module = load_script(SCRIPT, config={"interval": 5, "hide_tray_icon": True})
        saved = json.loads((tmp_path / CONFIG_FILE).read_text(encoding="utf-8"))
        assert saved["interval"] == 5
        assert saved["word_backup_path"] == "C:\\Officebackup\\wordbackup"
        assert set(module.default_config) <= set(saved)

    def test_invalid_json_falls_back_to_defaults(self, load_script, tmp_path):
        (tmp_path / CONFIG_FILE).write_text("{not valid json", encoding="utf-8")
        module = load_script(SCRIPT)
        saved = json.loads((tmp_path / CONFIG_FILE).read_text(encoding="utf-8"))
        assert set(saved) == set(module.default_config)
        assert saved["ppt_backup_path"] == "C:\\Officebackup\\pptbackup"


class TestLogHandling:
    def test_previous_log_is_archived(self, load_script, tmp_path):
        (tmp_path / "OBUlatest.log").write_text("old session", encoding="utf-8")
        load_script(SCRIPT, config=base_config(archive_previous_log=True))
        assert (tmp_path / "OBUprevious.log").read_text(encoding="utf-8") == "old session"
        assert "Session starts at" in (tmp_path / "OBUlatest.log").read_text(encoding="utf-8")

    def test_previous_log_is_deleted_when_archiving_disabled(self, load_script, tmp_path):
        (tmp_path / "OBUlatest.log").write_text("old session", encoding="utf-8")
        load_script(SCRIPT, config=base_config(archive_previous_log=False))
        assert not (tmp_path / "OBUprevious.log").exists()
        assert "old session" not in (tmp_path / "OBUlatest.log").read_text(encoding="utf-8")

    def test_log_print_increments_runid_and_writes_log(self, load_script, tmp_path, capsys):
        module = load_script(SCRIPT, config=base_config())
        start = module.runid
        module.log_print("hello world", source="unittest")
        assert module.runid == start + 1
        assert "hello world" in capsys.readouterr().out
        assert "-unittest] hello world" in (tmp_path / "OBUlatest.log").read_text(encoding="utf-8")


class TestOpenListConfigNormalization:
    def test_url_and_target_folder_are_normalized(self, load_script):
        module = load_script(
            SCRIPT,
            config=base_config(
                upload_to_openlist_enable=True,
                openlist_url="http://server.example/",
                openlist_username="user",
                openlist_password="pw",
                openlist_target_folder="backups/office/",
            ),
        )
        assert module.openlist_url == "http://server.example"
        assert module.openlist_target_folder == "/backups/office"
        assert module.config["upload_to_openlist_enable"] is True

    def test_empty_target_folder_defaults_to_root_and_disables_upload(self, load_script):
        module = load_script(
            SCRIPT,
            config=base_config(
                upload_to_openlist_enable=True,
                openlist_url="http://server.example",
                openlist_username="",
                openlist_target_folder="",
            ),
        )
        assert module.openlist_target_folder == "/"
        # username is empty -> upload is force-disabled for the session
        assert module.config["upload_to_openlist_enable"] is False


class TestCalculateMd5:
    def test_returns_expected_digest(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        target = tmp_path / "file.bin"
        target.write_bytes(b"some binary content" * 1000)
        expected = hashlib.md5(target.read_bytes()).hexdigest()
        assert module.calculate_md5(str(target)) == expected

    def test_returns_none_for_missing_file(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        assert module.calculate_md5(str(tmp_path / "missing.bin")) is None


class TestRemoveReadonly:
    def test_adds_write_permission(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        target = tmp_path / "readonly.txt"
        target.write_text("data", encoding="utf-8")
        os.chmod(target, stat.S_IRUSR | stat.S_IRGRP | stat.S_IROTH)
        module.remove_readonly(str(target))
        assert os.stat(target).st_mode & 0o200

    def test_missing_file_is_ignored(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        module.remove_readonly(str(tmp_path / "missing.txt"))  # must not raise


class TestUploadQueue:
    def test_add_and_duplicate_detection(self, load_script):
        module = load_script(SCRIPT, config=base_config())
        module.openlist_ready = True  # skip lazy OpenList initialization
        assert module.add_to_upload_queue("a.pptx", "/tmp/a.pptx") is True
        assert module.is_file_in_upload_queue("a.pptx") is True
        assert module.add_to_upload_queue("a.pptx", "/tmp/a.pptx") is False
        assert module.upload_queue == [("a.pptx", "/tmp/a.pptx")]

    def test_is_file_on_openlist_uses_remote_cache(self, load_script):
        module = load_script(SCRIPT, config=base_config())
        module.openlist_remote_files = {"remote.docx"}
        assert module.is_file_on_openlist("remote.docx") is True
        assert module.is_file_on_openlist("local.docx") is False

    def test_check_file_exists_respects_upload_toggle(self, load_script):
        module = load_script(SCRIPT, config=base_config())
        module.openlist_remote_files = {"remote.docx"}
        module.config["upload_to_openlist_enable"] = False
        assert module.check_file_exists_on_openlist("remote.docx") is False
        module.config["upload_to_openlist_enable"] = True
        assert module.check_file_exists_on_openlist("remote.docx") is True
        assert module.check_file_exists_on_openlist("other.docx") is False


class TestTimeoutDecorator:
    def test_returns_function_result(self, load_script):
        module = load_script(SCRIPT, config=base_config())

        @module.timeout(seconds=600)
        def add(a, b):
            return a + b

        assert add(2, 3) == 5

    def test_returns_none_and_logs_on_exception(self, load_script, capsys):
        module = load_script(SCRIPT, config=base_config())

        @module.timeout(seconds=600)
        def boom():
            raise ValueError("expected failure")

        assert boom() is None
        assert "Error in boom: expected failure" in capsys.readouterr().out

    def test_empty_config_value_falls_back_to_default(self, load_script):
        module = load_script(SCRIPT, config=base_config(backup_timeout=""))

        @module.timeout(seconds=600, config_key="backup_timeout")
        def fast():
            return "done"

        assert fast() == "done"


class TestBackupOpenFiles:
    def _fake_document(self, path):
        return SimpleNamespace(FullName=str(path))

    def test_backs_up_open_presentation(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        source = tmp_path / "deck.pptx"
        source.write_bytes(b"presentation-bytes")
        backup_dir = tmp_path / "backups"
        module.win32.Dispatch.return_value = SimpleNamespace(
            Presentations=[self._fake_document(source)]
        )

        module.save_open_ppt_files(str(backup_dir))

        backup_file = backup_dir / "deck.pptx"
        assert backup_file.read_bytes() == b"presentation-bytes"
        module.win32.Dispatch.assert_called_with("PowerPoint.Application")

    def test_unchanged_file_is_skipped_by_md5(self, load_script, tmp_path, capsys):
        module = load_script(SCRIPT, config=base_config())
        source = tmp_path / "deck.pptx"
        source.write_bytes(b"presentation-bytes")
        backup_dir = tmp_path / "backups"
        module.win32.Dispatch.return_value = SimpleNamespace(
            Presentations=[self._fake_document(source)]
        )

        module.save_open_ppt_files(str(backup_dir))
        capsys.readouterr()
        module.save_open_ppt_files(str(backup_dir))

        assert "skipped backup (MD5 match)" in capsys.readouterr().out

    def test_changed_file_is_backed_up_again(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        source = tmp_path / "deck.pptx"
        source.write_bytes(b"version-1")
        backup_dir = tmp_path / "backups"
        module.win32.Dispatch.return_value = SimpleNamespace(
            Presentations=[self._fake_document(source)]
        )

        module.save_open_ppt_files(str(backup_dir))
        source.write_bytes(b"version-2-changed")
        module.save_open_ppt_files(str(backup_dir))

        assert (backup_dir / "deck.pptx").read_bytes() == b"version-2-changed"

    def test_no_open_files_logs_normal_request(self, load_script, tmp_path, capsys):
        module = load_script(SCRIPT, config=base_config())
        module.win32.Dispatch.return_value = SimpleNamespace(Presentations=[])
        module.save_open_ppt_files(str(tmp_path / "backups"))
        assert "No ppt available now (Normal request)" in capsys.readouterr().out

    def test_com_error_logs_application_not_detected(self, load_script, tmp_path, capsys):
        module = load_script(SCRIPT, config=base_config())
        com_error = type("com_error", (Exception,), {})
        module.win32.Dispatch.side_effect = com_error("no app")
        module.save_open_word_files(str(tmp_path / "backups"))
        assert "No doc available now (Word application not detected)" in capsys.readouterr().out

    def test_backed_up_file_is_queued_for_upload(self, load_script, tmp_path):
        module = load_script(SCRIPT, config=base_config())
        module.config["upload_to_openlist_enable"] = True
        module.openlist_ready = True  # skip lazy OpenList initialization
        source = tmp_path / "deck.pptx"
        source.write_bytes(b"presentation-bytes")
        backup_dir = tmp_path / "backups"
        module.win32.Dispatch.return_value = SimpleNamespace(
            Presentations=[self._fake_document(source)]
        )

        with mock.patch.object(module, "upload_to_openlist"):
            module.save_open_ppt_files(str(backup_dir))

        assert module.is_file_in_upload_queue("deck.pptx") is True


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
        assert module.config["accurate_backup_enable"] is False
        saved = json.loads((tmp_path / CONFIG_FILE).read_text(encoding="utf-8"))
        assert saved["accurate_backup_enable"] is False
        assert module.accurate_backup_running is False

    def test_missing_source_path_waits_for_next_request(self, load_script, tmp_path, capsys):
        module = load_script(
            SCRIPT,
            config=base_config(
                accurate_backup_enable=True,
                accurate_backup_source_path=str(tmp_path / "does-not-exist"),
                accurate_backup_target_path=str(tmp_path / "target"),
            ),
        )

        module.accurate_backup()

        assert "wait for the next request" in capsys.readouterr().out
        assert module.config["accurate_backup_enable"] is True
        assert module.accurate_backup_running is False

    def test_empty_paths_disable_accurate_backup_for_session(self, load_script):
        module = load_script(
            SCRIPT,
            config=base_config(accurate_backup_enable=True, accurate_backup_source_path=""),
        )
        assert module.config["accurate_backup_enable"] is False


class _FakeAiohttpResponse:
    def __init__(self, payload):
        self._payload = payload

    async def json(self):
        return self._payload

    async def __aenter__(self):
        return self

    async def __aexit__(self, *exc):
        return False


class _FakeAiohttpSession:
    def __init__(self, payloads, calls):
        self._payloads = payloads
        self._calls = calls

    async def __aenter__(self):
        return self

    async def __aexit__(self, *exc):
        return False

    def put(self, url, data=None, headers=None):
        self._calls.append({"url": url, "data": data, "headers": headers})
        return _FakeAiohttpResponse(self._payloads[len(self._calls) - 1])


def _install_fake_aiohttp(monkeypatch, payloads):
    calls = []
    fake_aiohttp = types.ModuleType("aiohttp")
    fake_aiohttp.ClientSession = lambda: _FakeAiohttpSession(payloads, calls)
    monkeypatch.setitem(sys.modules, "aiohttp", fake_aiohttp)
    return calls


class TestChunkedStreamUpload:
    def _fake_client(self):
        client = SimpleNamespace(
            token="jwt-token",
            headers={"User-Agent": "OBU-test"},
            endpoint="http://server.example/",
        )
        client.upload = mock.AsyncMock(return_value="standard-upload")
        return client

    def test_empty_file_uses_standard_upload(self, load_script, tmp_path, monkeypatch):
        module = load_script(SCRIPT, config=base_config())
        _install_fake_aiohttp(monkeypatch, payloads=[])
        empty = tmp_path / "empty.pptx"
        empty.write_bytes(b"")
        client = self._fake_client()

        result = asyncio.run(
            module.chunked_stream_upload(client, "/backups/empty.pptx", str(empty))
        )

        assert result == "standard-upload"
        client.upload.assert_awaited_once_with("/backups/empty.pptx", str(empty))

    def test_single_chunk_upload_sends_content_range(self, load_script, tmp_path, monkeypatch):
        module = load_script(SCRIPT, config=base_config())
        calls = _install_fake_aiohttp(monkeypatch, payloads=[{"code": 200}])
        local = tmp_path / "deck.pptx"
        local.write_bytes(b"0123456789")
        client = self._fake_client()

        result = asyncio.run(
            module.chunked_stream_upload(client, "/backups/deck 1.pptx", str(local))
        )

        assert result is True
        assert len(calls) == 1
        call = calls[0]
        assert call["url"] == "http://server.example/api/fs/put"
        assert call["data"] == b"0123456789"
        assert call["headers"]["Content-Range"] == "bytes 0-9/10"
        assert call["headers"]["Authorization"] == "jwt-token"
        assert call["headers"]["File-Path"] == "/backups/deck%201.pptx"

    def test_failed_chunk_raises(self, load_script, tmp_path, monkeypatch):
        module = load_script(SCRIPT, config=base_config())
        _install_fake_aiohttp(
            monkeypatch, payloads=[{"code": 500, "message": "storage failure"}]
        )
        local = tmp_path / "deck.pptx"
        local.write_bytes(b"0123456789")

        with pytest.raises(Exception, match="storage failure"):
            asyncio.run(
                module.chunked_stream_upload(
                    self._fake_client(), "/backups/deck.pptx", str(local)
                )
            )
