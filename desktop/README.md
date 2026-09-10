# TNBES Energy Audit — Desktop Tools (Python)

Dua cara guna folder ini:

## A) Compile Excel dari data app WEB (disyorkan)

Kalau anda key-in data guna app web (HTML, sync ke Google Sheet), guna ini
untuk tarik data tu dan jana fail Excel:

```
pip install -r requirements.txt
python pull_and_compile.py
```

Akan minta URL Google Apps Script Web App (sekali sahaja, disimpan lepas
tu). Kemudian pilih nak compile bangunan mana / semua sekali, dan fail
Excel terus dijana — guna engine `openpyxl` yang stabil (tak crash pada
template kompleks).

## B) App key-in data berasingan (Tkinter, simpan local)

Kalau anda TAK guna app web dan nak key-in terus dalam Python (data
disimpan dalam `tnbes_data.json` tempatan sahaja, bukan Google Sheet):

```
python app.py
```

---

## Setup (sekali sahaja)

1. Pasang Python 3.9+ (kalau belum ada): https://www.python.org/downloads/
2. Buka Command Prompt / Terminal di dalam folder ini, jalankan:
   ```
   pip install -r requirements.txt
   ```

## Jalankan App

```
python app.py
```

## Cara Guna

0. **Nama Project & Team** — isi di bahagian atas app (akan auto-tulis ke
   header setiap sheet Excel).
1. **Tambah Bangunan** — klik "+ Tambah" di panel Bangunan, masukkan nama
   DAN location (alamat/lokasi bangunan). Boleh edit balik bila-bila masa
   guna butang "Edit".
2. **Tambah Floor** — pilih bangunan dulu, klik "+ Tambah" di panel Floor.
3. **Isi Data** — pilih floor, klik tab kategori (Lighting, Split Unit, AHU, dll),
   klik "+ Tambah Rekod" untuk buka borang, isi, klik Simpan.
   - Dropdown automatik untuk Lighting Equipment, Type of Lamps, Type of Control,
     Status Reflection, Type of MS Standard, dll.
   - **Pilihan "Other"**: kalau item yang anda nak takde dalam senarai (contoh
     jenis lampu "T8"), pilih "Other" dalam dropdown — kotak teks di sebelah
     akan aktif, taip terus (contoh "T8"), automatik disimpan bila anda klik Simpan.
   - Field nombor (Lux, Temperature, dll) hanya terima nombor.
   - Field masa (Weekdays Start/End, dll) guna spinner jam:minit (24 jam).
4. **AHU khas**: satu rekod AHU = satu unit AHU. Kalau ada 4 unit AHU dalam
   satu floor, tambah 4 rekod — nanti bila export, akan jadi **4 sheet berasingan**.
5. **Export**:
   - "⬇ Export Floor Ini ke Excel" — muat turun 1 fail xlsx untuk floor yang
     sedang dipilih sahaja.
   - "💾 Compile Semua Bangunan -> 1 Excel" — gabungkan SEMUA bangunan & floor
     ke dalam SATU fail xlsx (setiap kategori/AHU jadi sheet berasingan,
     dinamakan ikut Bangunan + Floor + Kategori).

Data disimpan automatik ke `tnbes_data.json` (dalam folder sama) setiap kali
anda tambah/edit/padam rekod — jadi kalau app ditutup, data tak hilang.

## Struktur Fail

- `app.py` — GUI utama (Tkinter)
- `config.py` — senarai kategori, field, pilihan dropdown, dan **cell-mapping
  tepat** bagi setiap kategori (rujuk sini kalau nak audit/ubah posisi cell)
- `excel_writer.py` — engine yang tulis data ke template Excel
- `templates_data.py` — template Excel (base64) tertanam terus dari
  `Template__1_.xlsx` anda — 8 kategori: lighting, lighting_sampling,
  acmv_split, acmv_ahu, acmv_chiller, office, lift, datalogger
- `data_store.py` — simpan/baca data ke `tnbes_data.json`

## Nota Penting / Limitasi Template

Semasa saya audit semula `Template__1_.xlsx` untuk buat app ini, saya jumpa
beberapa sheet **struktur cell dia dah berbeza** daripada yang saya map
sebelum ini (dalam app web HTML anda):

- **Office/General Equipment**: header sekarang di row 7-8, jadi data
  sebenar mula row **9** (bukan row 7 macam sebelum ini).
- **Lift**: bukan borang menegak lagi — sekarang ada **3 jadual berasingan**
  (General Info row 6+, Measurement Log row 27+, Operating Profile row 45+),
  semuanya dipaut ikut "Lift ID / Tag No." di lajur A.
- **Chiller/Pump**: sama macam Lift — **2 jadual berasingan** (General Info
  row 9+, Measurement Log row 27+), dipaut ikut "Pump ID / Tag No.".
- **Data Logger**: sekarang jadual biasa (row 7+, satu row = satu logger),
  bukan borang menegak macam sebelum ini.

App Python ini (dan cell-mapping dalam `config.py`) sudah ikut struktur
**TERKINI** ini. **App web HTML yang saya bagi sebelum ini masih guna
struktur LAMA untuk Lift/Chiller/Office/Data Logger** — kalau anda nak
guna app web tu lagi, bagitahu saya, saya boleh selaraskan dia sekali.

Satu lagi limitasi template (bukan salah app): kotak "Outside Air (OA)"
dalam sheet AHU cuma ada **1 cell** untuk Temp+RH+CO2 (bercantum, bukan
3 cell berasingan) — jadi field tu jadi 1 kotak teks bebas dalam borang.
