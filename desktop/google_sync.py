"""
Tarik data terus dari Google Sheet (backend yang SAMA dengan app web HTML),
dan susun semula struktur Project > Team > Phase > Building > Floor supaya
boleh terus dihantar ke excel_writer.py untuk di-compile jadi Excel.

Guna URL Apps Script Web App yang SAMA yang anda tampal dalam app.html
(pemalar GOOGLE_SHEETS_WEBAPP_URL).
"""
import json
import urllib.request
import urllib.error

CONFIG_FILE = "google_sync_config.json"


def load_webapp_url():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get("webapp_url", "")
    except Exception:
        return ""


def save_webapp_url(url):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"webapp_url": url}, f)


def fetch_all(webapp_url):
    """Panggil action=dump_all, pulangkan dict {key: json_string_value}."""
    url = f"{webapp_url}?action=dump_all"
    with urllib.request.urlopen(url, timeout=30) as resp:
        body = json.loads(resp.read().decode("utf-8"))
    if not body.get("ok"):
        raise RuntimeError(f"Google Sheets API error: {body.get('error')}")
    raw = body["data"]
    # Setiap value disimpan sbg JSON string (ikut cara sset() app web menulisnya)
    parsed = {}
    for k, v in raw.items():
        try:
            parsed[k] = json.loads(v)
        except (json.JSONDecodeError, TypeError):
            parsed[k] = v
    return parsed


def build_hierarchy(data):
    """
    Bina semula struktur penuh dari key-value flat Google Sheet:
      projects, teams, phases, buildings, floors:{buildingId}
    Pulangkan: list of buildings siap pakai utk export_project_to_one_workbook:
      [{"name", "location", "project", "floors":[{"name","data"}, ...]}, ...]
    """
    projects = {p["id"]: p for p in data.get("projects", [])}
    teams = {t["id"]: t for t in data.get("teams", [])}
    phases = {ph["id"]: ph for ph in data.get("phases", [])}
    buildings = data.get("buildings", [])

    result = []
    for b in buildings:
        phase = phases.get(b.get("phaseId"))
        team = teams.get(phase.get("teamId")) if phase else None
        project = projects.get(team.get("projectId")) if team else None

        floors_raw = data.get(f"floors:{b['id']}", [])
        floors = [{"name": f.get("name", "Floor"), "data": f.get("data", {})} for f in floors_raw]

        result.append({
            "name": b.get("name", "Building"),
            "location": b.get("location", ""),
            "project": project.get("name", "") if project else "",
            "team": team.get("name", "") if team else "",
            "phase": phase.get("name", "") if phase else "",
            "floors": floors,
        })
    return result
