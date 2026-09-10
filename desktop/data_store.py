"""Simpan/baca data project (bangunan/floor/data) ke fail JSON tempatan."""
import json
import os

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tnbes_data.json")


def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"teamName": "", "projectName": "", "buildings": []}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
