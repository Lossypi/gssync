import json
from pathlib import Path
import pytest
from gssync.config import load_config, save_config, Config


def test_load_config_returns_defaults_when_no_file(tmp_path, monkeypatch):
    monkeypatch.setattr("gssync.config.CONFIG_FILE", tmp_path / "config.json")
    config = load_config()
    assert config.spreadsheet_url == ""
    assert config.file_path == ""
    assert config.file_format == "xlsx"


def test_save_and_load_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr("gssync.config.CONFIG_DIR", tmp_path)
    monkeypatch.setattr("gssync.config.CONFIG_FILE", tmp_path / "config.json")
    cfg = Config(spreadsheet_url="https://example.com", file_path="C:/data/f.xlsx", file_format="json")
    save_config(cfg)
    loaded = load_config()
    assert loaded.spreadsheet_url == "https://example.com"
    assert loaded.file_path == "C:/data/f.xlsx"
    assert loaded.file_format == "json"


def test_save_creates_directory(tmp_path, monkeypatch):
    new_dir = tmp_path / "nested" / "gssync"
    monkeypatch.setattr("gssync.config.CONFIG_DIR", new_dir)
    monkeypatch.setattr("gssync.config.CONFIG_FILE", new_dir / "config.json")
    save_config(Config())
    assert (new_dir / "config.json").exists()
