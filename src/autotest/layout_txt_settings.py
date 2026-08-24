# -*- coding: utf-8 -*-
"""Layout generator GUI settings stored outside the application folder."""

import json
import os
import tempfile
from pathlib import Path


CONFIG_ENV = "AUTOTEST_LAYOUT_GUI_CONFIG"


class LayoutGuiSettingsError(Exception):
    """The persisted GUI settings could not be read or written."""


def default_settings_path():
    override = os.environ.get(CONFIG_ENV, "").strip()
    if override:
        return Path(override).expanduser()

    app_data = os.environ.get("APPDATA", "").strip()
    if os.name == "nt" and app_data:
        return Path(app_data) / "AUTO_TEST_BATCH" / "layout_txt_gui.json"
    return Path.home() / ".auto_test_batch" / "layout_txt_gui.json"


def load_settings(path=None):
    target = Path(path) if path is not None else default_settings_path()
    if not target.is_file():
        return {}
    try:
        with target.open("r", encoding="utf-8") as stream:
            data = json.load(stream)
    except (OSError, ValueError) as exc:
        raise LayoutGuiSettingsError("設定ファイルを読めません: %s" % exc)
    if not isinstance(data, dict):
        raise LayoutGuiSettingsError("設定ファイルの最上位はJSONオブジェクトにしてください。")
    return data


def save_settings(settings, path=None):
    if not isinstance(settings, dict):
        raise LayoutGuiSettingsError("保存する設定は辞書形式で指定してください。")
    target = Path(path) if path is not None else default_settings_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        handle, temp_name = tempfile.mkstemp(
            prefix=target.name + ".", suffix=".tmp", dir=str(target.parent))
        try:
            with os.fdopen(handle, "w", encoding="utf-8", newline="\n") as stream:
                json.dump(settings, stream, ensure_ascii=False, indent=2, sort_keys=True)
                stream.write("\n")
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temp_name, str(target))
        except Exception:
            try:
                os.unlink(temp_name)
            except OSError:
                pass
            raise
    except (OSError, TypeError, ValueError) as exc:
        raise LayoutGuiSettingsError("設定ファイルを保存できません: %s" % exc)
    return target
