"""
Tarik data dari Google Sheet (yang dikemaskini oleh app web HTML) dan
compile jadi SATU fail Excel — guna engine Excel Python yang stabil
(tak akan crash macam versi JS/browser untuk template yang kompleks).

Cara guna:
    python pull_and_compile.py
"""
import sys
from openpyxl import Workbook

from google_sync import load_webapp_url, save_webapp_url, fetch_all, build_hierarchy
from excel_writer import export_project_to_one_workbook, export_floor_to_workbook


def main():
    print("=" * 60)
    print("TNBES Energy Audit — Pull dari Google Sheet & Compile Excel")
    print("=" * 60)

    webapp_url = load_webapp_url()
    if not webapp_url:
        webapp_url = input("\nTampal URL Google Apps Script Web App (.../exec): ").strip()
        save_webapp_url(webapp_url)
    else:
        print(f"\nGuna URL tersimpan: {webapp_url}")
        change = input("Nak tukar URL? (y/N): ").strip().lower()
        if change == "y":
            webapp_url = input("URL baharu: ").strip()
            save_webapp_url(webapp_url)

    print("\nMenarik data dari Google Sheet...")
    try:
        data = fetch_all(webapp_url)
    except Exception as e:
        print(f"\n❌ Gagal sambung ke Google Sheet: {e}")
        sys.exit(1)

    buildings = build_hierarchy(data)
    if not buildings:
        print("\n⚠️  Tiada bangunan dijumpai dalam Google Sheet.")
        sys.exit(0)

    print(f"\nDijumpai {len(buildings)} bangunan:")
    for i, b in enumerate(buildings):
        n_floors = len(b["floors"])
        print(f"  [{i+1}] {b['name']}  (Project: {b['project'] or '-'}, "
              f"Team: {b['team'] or '-'}, Phase: {b['phase'] or '-'}, {n_floors} floor)")

    print("\nPilihan:")
    print("  [A] Compile SEMUA bangunan -> 1 fail Excel")
    print("  [nombor] Compile 1 bangunan sahaja (contoh: 1)")
    choice = input("\nPilih (A / nombor): ").strip().lower()

    if choice == "a" or choice == "":
        selected = buildings
        out_name = "Compiled Overall Data.xlsx"
    else:
        try:
            idx = int(choice) - 1
            selected = [buildings[idx]]
            out_name = f"{buildings[idx]['name']} - Compiled.xlsx".replace("/", "-")
        except (ValueError, IndexError):
            print("Pilihan tidak sah.")
            sys.exit(1)

    print("\nMenjana fail Excel...")
    wb = export_project_to_one_workbook(selected)
    wb.save(out_name)
    print(f"\n✅ Selesai! Fail disimpan: {out_name}")
    print(f"   Sheet dijana: {len(wb.sheetnames)}")


if __name__ == "__main__":
    main()
