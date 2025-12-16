import os
import tempfile
import json
from pathlib import Path
import Heimdall.Main as Main


def test_save_and_load_settings(tmp_path):
    tmp_file = tmp_path / 'settings.json'
    Main.SETTINGS_FILE = tmp_file
    data = {'a': 1, 'b': 'c'}
    assert Main.save_settings(data) is True
    loaded = Main.load_settings()
    assert loaded == data


def test_search_files(tmp_path):
    d = tmp_path / 'docs'
    d.mkdir()
    (d / 'notes.txt').write_text('hello')
    (d / 'image.png').write_text('x')
    res = Main.search_files('notes', start_dirs=[str(d)], extensions=['.txt'])
    assert any('notes.txt' in p for p in res)


def test_generate_chart(tmp_path):
    out = tmp_path / 'chart.png'
    p = Main.generate_chart([1, 2, 3], filename=str(out.name))
    assert p is not None
    assert (tmp_path / str(out.name)).exists() or Path(p).exists()
