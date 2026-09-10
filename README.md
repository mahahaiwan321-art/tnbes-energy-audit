# TNBES Energy Audit Progress Dashboard

Dua bahagian, satu data:

```
/web       -> App web (HTML) untuk KEY-IN DATA — dipublish via GitHub Pages
/desktop   -> App Python untuk COMPILE / DOWNLOAD Excel — dijalankan tempatan
```

**Kenapa dipisahkan?** Key-in data lebih senang guna app web (boleh diakses
mana-mana peranti, tiada install). Tapi bila nak compile ke fail Excel yang
kompleks (template dengan banyak formula & merged cell), engine Python
(`openpyxl`) jauh lebih stabil daripada engine JS/browser (`SheetJS`) —
tak crash walau template rumit macam AHU.

Kedua-dua bahagian **share DATA yang SAMA** melalui satu Google Sheet
(bertindak sebagai "database").

## Aliran Kerja

1. **Key-in data** — buka link app web (GitHub Pages), isi data audit
   (Lighting, ACMV, Office, Lift, dll). Setiap kali simpan, data terus
   masuk ke Google Sheet.
2. **Compile ke Excel** — di komputer, jalankan `desktop/pull_and_compile.py`.
   Script ni akan:
   - Tarik semua data terkini dari Google Sheet yang sama,
   - Susun semula ikut Project > Team > Phase > Building > Floor,
   - Jana SATU fail Excel (template asal, formula & formatting dikekalkan).

## Setup — Bahagian Web (`/web`)

1. Ikuti panduan Google Sheets backend (`web/Code.gs`) — deploy sebagai
   Apps Script Web App, dapatkan URL `.../exec`.
2. Buka `web/index.html`, cari `GOOGLE_SHEETS_WEBAPP_URL`, tampal URL tersebut.
3. Publish `web/index.html` (namakan `index.html`, letak di root repo atau
   guna GitHub Pages dengan folder `/web` sebagai source) supaya boleh
   diakses via link.

## Setup — Bahagian Desktop (`/desktop`)

```
cd desktop
pip install -r requirements.txt
python pull_and_compile.py
```

Kali pertama run, ia akan minta URL Google Apps Script Web App yang SAMA
(dari langkah web di atas) — akan disimpan automatik untuk kali seterusnya.

## Struktur GitHub Repo (disyorkan)

```
repo-root/
├── README.md              <- fail ni
├── web/
│   ├── index.html         <- app key-in data (GitHub Pages source)
│   └── Code.gs             <- backend Google Apps Script (rujukan)
└── desktop/
    ├── pull_and_compile.py <- entry point (jalankan ni)
    ├── google_sync.py
    ├── excel_writer.py
    ├── config.py
    ├── templates_data.py
    ├── requirements.txt
    └── README.md
```

Untuk GitHub Pages: Settings → Pages → Source: Deploy from branch →
Branch: main, Folder: `/web`.
