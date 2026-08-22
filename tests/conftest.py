"""Shared pytest fixtures for testing the OBU backup scripts.

The backup scripts (Officebackup6.3.py / Officebackup6.3Core.py) are
Windows-only, top-level scripts: they import win32com/pystray/PIL, touch
ctypes.windll and end in an infinite main loop. To unit test their functions
the loader below execs the script source up to (excluding) the main loop,
inside a temporary working directory, with the Windows-only modules replaced
by mocks.
"""

import ctypes
import json
import sys
import types
from pathlib import Path
from unittest import mock

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]

# Everything from this line onwards is the main loop of the script and is
# excluded when loading the script for tests.
MAIN_SENTINEL = "print('Program initialization completed"

CONFIG_FILES = {
    "Officebackup6.3.py": "OBU6.3.json",
    "Officebackup6.3Core.py": "OBU6.3Core.json",
}

FAKE_MODULE_NAMES = [
    "win32com",
    "win32com.client",
    "pystray",
    "PIL",
    "PIL.Image",
    "alist",
]


@pytest.fixture
def load_script(tmp_path, monkeypatch):
    """Return a loader that execs a backup script (minus its main loop).

    The loader chdirs into ``tmp_path`` so config/log files are created
    there, optionally pre-writes a config file, mocks all Windows-only
    modules and returns the resulting module namespace.
    """

    def _load(script_name, config=None):
        monkeypatch.chdir(tmp_path)

        fakes = {name: mock.MagicMock(name=name) for name in FAKE_MODULE_NAMES}
        for name, fake in fakes.items():
            monkeypatch.setitem(sys.modules, name, fake)
        monkeypatch.setattr(ctypes, "windll", mock.MagicMock(), raising=False)

        if config is not None:
            config_file = tmp_path / CONFIG_FILES[script_name]
            config_file.write_text(
                json.dumps(config, ensure_ascii=False), encoding="utf-8"
            )

        script_path = REPO_ROOT / script_name
        source = script_path.read_text(encoding="utf-8")
        cut = source.find(MAIN_SENTINEL)
        assert cut != -1, f"main-loop sentinel not found in {script_name}"

        module = types.ModuleType("obu_under_test")
        module.__file__ = str(script_path)
        code = compile(source[:cut], str(script_path), "exec")
        exec(code, module.__dict__)
        return module

    return _load
