"""
TNBES Energy Audit — Data Collection (Desktop App / Tkinter)
-------------------------------------------------------------
Bangunan -> Floor -> Kategori (Lighting, ACMV Split/AHU/Chiller, Office,
Lift, Data Logger) -> rekod data. Setiap rekod ditulis ke cell Excel yang
TEPAT (ikut template asal anda) apabila di-export.

Jalankan: python app.py
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re

from config import (
    CATEGORIES, sub_types_for, category_name, subtype_name,
    fields_for, FIELD_DROPDOWNS, NUMERIC_ONLY_FIELDS, TIME_FIELDS, OTHER_TRIGGER_VALUES,
)
from data_store import load_data, save_data
from excel_writer import export_floor_to_workbook, export_project_to_one_workbook
from openpyxl import Workbook

# Field ringkas utk dipaparkan dalam senarai (Treeview) bagi setiap kategori
SUMMARY_FIELDS = {
    "lighting:main": ["Level", "Room / Area", "Lighting Equipment"],
    "lighting:sampling": ["Level", "Room / Area"],
    "acmv:split": ["Level", "Room / Area", "Brand", "Model"],
    "acmv:ahu": ["Equipment", "AHU Tag No.", "Zone", "Level"],
    "acmv:chiller": ["Pump ID / Tag No.", "Building / Location", "Manufacturer"],
    "office:main": ["Level", "Room / Area", "Type of Equipment"],
    "lift:main": ["Lift ID / Tag No.", "Building / Location"],
    "datalogger:main": ["Logger ID", "Brand / Model", "Installation Location"],
}


def is_time_str(s):
    return bool(re.match(r"^\d{1,2}:\d{2}$", str(s).strip()))


# ---------------------------------------------------------------------------
class ScrollableFrame(ttk.Frame):
    """Frame dengan scrollbar menegak — utk borang yang banyak field (contoh AHU)."""
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # scroll guna roda tetikus
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)


# ---------------------------------------------------------------------------
class RecordDialog(tk.Toplevel):
    """Dialog borang untuk tambah/edit SATU rekod bagi satu kategori."""
    def __init__(self, parent, cat_id, sub_id, fields, initial=None):
        super().__init__(parent)
        self.title(f"{category_name(cat_id)} — {subtype_name(cat_id, sub_id)}")
        self.geometry("560x640")
        self.result = None
        self.fields = fields
        self.vars = {}

        scrollable = ScrollableFrame(self)
        scrollable.pack(fill="both", expand=True, padx=10, pady=10)
        form = scrollable.inner

        initial = initial or {}

        for i, f in enumerate(fields):
            row = ttk.Frame(form)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=f, width=32, anchor="w", wraplength=230).pack(side="left")

            val = initial.get(f, "")

            if f in FIELD_DROPDOWNS:
                options = FIELD_DROPDOWNS[f]
                other_trigger = next((o for o in options if o in OTHER_TRIGGER_VALUES), None)
                # Kalau nilai sedia ada BUKAN salah satu pilihan dalam list,
                # ia mesti nilai custom yg pernah ditaip sebelum ini (Other).
                is_custom_initial = bool(val) and (val not in options) and other_trigger is not None

                var = tk.StringVar(value=(other_trigger if is_custom_initial else val))
                cb = ttk.Combobox(row, textvariable=var, values=options, width=22, state="readonly")
                cb.pack(side="left")

                if other_trigger:
                    custom_var = tk.StringVar(value=(val if is_custom_initial else ""))
                    custom_entry = ttk.Entry(row, textvariable=custom_var, width=20)
                    custom_entry.pack(side="left", padx=(4, 0), fill="x", expand=True)

                    def _toggle(event=None, cvar=var, centry=custom_entry, trigger=other_trigger):
                        if cvar.get() == trigger:
                            centry.configure(state="normal")
                        else:
                            centry.configure(state="disabled")

                    cb.bind("<<ComboboxSelected>>", _toggle)
                    _toggle()  # set keadaan awal ikut nilai semasa
                    self.vars[f] = (var, custom_var, other_trigger)
                else:
                    self.vars[f] = var

            elif f in TIME_FIELDS:
                # keypad masa 24 jam: 2 Spinbox (jam / minit)
                hh_val, mm_val = "", ""
                if is_time_str(val):
                    hh_val, mm_val = val.split(":")
                hh_var = tk.StringVar(value=hh_val)
                mm_var = tk.StringVar(value=mm_val)
                spin_h = tk.Spinbox(row, from_=0, to=23, width=3, format="%02.0f",
                                     textvariable=hh_var, wrap=True)
                spin_h.pack(side="left")
                ttk.Label(row, text=":").pack(side="left")
                spin_m = tk.Spinbox(row, from_=0, to=59, width=3, format="%02.0f",
                                     textvariable=mm_var, wrap=True)
                spin_m.pack(side="left")
                self.vars[f] = (hh_var, mm_var)

            elif f in NUMERIC_ONLY_FIELDS:
                vcmd = (self.register(self._validate_numeric), "%P")
                var = tk.StringVar(value=val)
                ent = ttk.Entry(row, textvariable=var, validate="key", validatecommand=vcmd)
                ent.pack(side="left", fill="x", expand=True)
                self.vars[f] = var

            else:
                var = tk.StringVar(value=val)
                ent = ttk.Entry(row, textvariable=var)
                ent.pack(side="left", fill="x", expand=True)
                self.vars[f] = var

        btns = ttk.Frame(self)
        btns.pack(fill="x", pady=8, padx=10)
        ttk.Button(btns, text="Simpan", command=self._on_save).pack(side="right", padx=4)
        ttk.Button(btns, text="Batal", command=self.destroy).pack(side="right")

        self.transient(parent)
        self.grab_set()

    @staticmethod
    def _validate_numeric(proposed):
        if proposed == "":
            return True
        return bool(re.match(r"^-?\d*\.?\d*$", proposed))

    def _on_save(self):
        out = {}
        for f in self.fields:
            v = self.vars[f]
            if isinstance(v, tuple) and len(v) == 2:  # time field (hh, mm)
                hh, mm = v[0].get().strip(), v[1].get().strip()
                if hh != "" and mm != "":
                    out[f] = f"{int(hh):02d}:{int(mm):02d}"
                else:
                    out[f] = ""
            elif isinstance(v, tuple) and len(v) == 3:  # dropdown dgn "Other" (combo_var, custom_var, trigger)
                combo_var, custom_var, trigger = v
                chosen = combo_var.get()
                if chosen == trigger:
                    out[f] = custom_var.get().strip() or trigger
                else:
                    out[f] = chosen
            else:
                out[f] = v.get()
        self.result = out
        self.destroy()


# ---------------------------------------------------------------------------
class TNBESApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TNBES Energy Audit — Data Collection")
        self.geometry("1100x680")

        self.data = load_data()
        if "buildings" not in self.data:
            self.data["buildings"] = []

        self.current_building_idx = None
        self.current_floor_idx = None

        self._build_ui()
        self._refresh_buildings()

    # ---------------- UI construction ----------------
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Label(top, text="Nama Project:").pack(side="left")
        self.project_var = tk.StringVar(value=self.data.get("projectName", ""))
        project_entry = ttk.Entry(top, textvariable=self.project_var, width=24)
        project_entry.pack(side="left", padx=(4, 12))
        project_entry.bind("<FocusOut>", lambda e: self._save())

        ttk.Label(top, text="Nama Team:").pack(side="left")
        self.team_var = tk.StringVar(value=self.data.get("teamName", ""))
        team_entry = ttk.Entry(top, textvariable=self.team_var, width=24)
        team_entry.pack(side="left", padx=6)
        team_entry.bind("<FocusOut>", lambda e: self._save())

        ttk.Button(top, text="💾 Compile Semua Bangunan -> 1 Excel",
                   command=self._export_all).pack(side="right", padx=4)

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=8, pady=6)

        # --- Panel Bangunan ---
        b_frame = ttk.LabelFrame(main, text="Bangunan")
        b_frame.pack(side="left", fill="y", padx=(0, 6))
        self.building_list = tk.Listbox(b_frame, width=22, exportselection=False)
        self.building_list.pack(fill="y", expand=True, padx=4, pady=4)
        self.building_list.bind("<<ListboxSelect>>", self._on_select_building)
        bb = ttk.Frame(b_frame)
        bb.pack(fill="x")
        ttk.Button(bb, text="+ Tambah", command=self._add_building).pack(side="left", expand=True, fill="x")
        ttk.Button(bb, text="Edit", command=self._edit_building).pack(side="left", expand=True, fill="x")
        ttk.Button(bb, text="Padam", command=self._delete_building).pack(side="left", expand=True, fill="x")

        # --- Panel Floor ---
        f_frame = ttk.LabelFrame(main, text="Floor")
        f_frame.pack(side="left", fill="y", padx=(0, 6))
        self.floor_list = tk.Listbox(f_frame, width=22, exportselection=False)
        self.floor_list.pack(fill="y", expand=True, padx=4, pady=4)
        self.floor_list.bind("<<ListboxSelect>>", self._on_select_floor)
        fb = ttk.Frame(f_frame)
        fb.pack(fill="x")
        ttk.Button(fb, text="+ Tambah", command=self._add_floor).pack(side="left", expand=True, fill="x")
        ttk.Button(fb, text="Padam", command=self._delete_floor).pack(side="left", expand=True, fill="x")

        # --- Panel Kategori (Notebook tabs) + rekod ---
        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True)

        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill="both", expand=True)

        self.tab_info = []  # list of (cat_id, sub_id, treeview)
        for cat in CATEGORIES:
            for sub in cat["subTypes"]:
                tab = ttk.Frame(self.notebook)
                label = subtype_name(cat["id"], sub["id"]) if len(cat["subTypes"]) > 1 else category_name(cat["id"])
                self.notebook.add(tab, text=label)

                cols = SUMMARY_FIELDS.get(f"{cat['id']}:{sub['id']}", fields_for(cat["id"], sub["id"])[:3])
                tree = ttk.Treeview(tab, columns=cols, show="headings", height=16)
                for c in cols:
                    tree.heading(c, text=c)
                    tree.column(c, width=150)
                tree.pack(fill="both", expand=True, padx=6, pady=6)

                btns = ttk.Frame(tab)
                btns.pack(fill="x", padx=6, pady=(0, 6))
                ttk.Button(btns, text="+ Tambah Rekod",
                           command=lambda c=cat["id"], s=sub["id"]: self._add_record(c, s)).pack(side="left")
                ttk.Button(btns, text="Edit",
                           command=lambda c=cat["id"], s=sub["id"]: self._edit_record(c, s)).pack(side="left", padx=4)
                ttk.Button(btns, text="Padam",
                           command=lambda c=cat["id"], s=sub["id"]: self._delete_record(c, s)).pack(side="left")

                self.tab_info.append((cat["id"], sub["id"], tree, cols))

        bottom = ttk.Frame(right)
        bottom.pack(fill="x", pady=6)
        ttk.Button(bottom, text="⬇ Export Floor Ini ke Excel",
                   command=self._export_floor).pack(side="left")

    # ---------------- Data helpers ----------------
    def _save(self):
        self.data["teamName"] = self.team_var.get()
        self.data["projectName"] = self.project_var.get()
        save_data(self.data)

    def _current_building(self):
        if self.current_building_idx is None:
            return None
        return self.data["buildings"][self.current_building_idx]

    def _current_floor(self):
        b = self._current_building()
        if b is None or self.current_floor_idx is None:
            return None
        return b["floors"][self.current_floor_idx]

    # ---------------- Bangunan ----------------
    def _refresh_buildings(self):
        self.building_list.delete(0, "end")
        for b in self.data["buildings"]:
            self.building_list.insert("end", b["name"])
        self.floor_list.delete(0, "end")
        self._refresh_all_tabs()

    def _add_building(self):
        result = _ask_building_info(self)
        if not result:
            return
        self.data["buildings"].append({"name": result["name"], "location": result["location"], "floors": []})
        self._save()
        self._refresh_buildings()

    def _edit_building(self):
        b = self._current_building()
        if b is None:
            messagebox.showinfo("Info", "Pilih bangunan dahulu.")
            return
        result = _ask_building_info(self, initial_name=b["name"], initial_location=b.get("location", ""))
        if not result:
            return
        b["name"] = result["name"]
        b["location"] = result["location"]
        self._save()
        self._refresh_buildings()
        self.building_list.selection_set(self.current_building_idx)

    def _delete_building(self):
        if self.current_building_idx is None:
            return
        if not messagebox.askyesno("Padam Bangunan", "Padam bangunan ini beserta semua floor & data?"):
            return
        del self.data["buildings"][self.current_building_idx]
        self.current_building_idx = None
        self.current_floor_idx = None
        self._save()
        self._refresh_buildings()

    def _on_select_building(self, event):
        sel = self.building_list.curselection()
        if not sel:
            return
        self.current_building_idx = sel[0]
        self.current_floor_idx = None
        b = self._current_building()
        self.floor_list.delete(0, "end")
        for f in b["floors"]:
            self.floor_list.insert("end", f["name"])
        self._refresh_all_tabs()

    # ---------------- Floor ----------------
    def _add_floor(self):
        b = self._current_building()
        if b is None:
            messagebox.showinfo("Info", "Pilih/tambah bangunan dahulu.")
            return
        name = _ask_text(self, "Floor Baharu", "Nama floor (contoh: Level 1):")
        if not name:
            return
        b["floors"].append({"name": name, "data": {}})
        self._save()
        self.floor_list.insert("end", name)

    def _delete_floor(self):
        b = self._current_building()
        if b is None or self.current_floor_idx is None:
            return
        if not messagebox.askyesno("Padam Floor", "Padam floor ini beserta semua data?"):
            return
        del b["floors"][self.current_floor_idx]
        self.current_floor_idx = None
        self._save()
        self.floor_list.delete(0, "end")
        for f in b["floors"]:
            self.floor_list.insert("end", f["name"])
        self._refresh_all_tabs()

    def _on_select_floor(self, event):
        sel = self.floor_list.curselection()
        if not sel:
            return
        self.current_floor_idx = sel[0]
        self._refresh_all_tabs()

    # ---------------- Rekod (dalam tab kategori) ----------------
    def _refresh_all_tabs(self):
        for cat_id, sub_id, tree, cols in self.tab_info:
            tree.delete(*tree.get_children())
            floor = self._current_floor()
            if floor is None:
                continue
            rows = floor["data"].get(f"{cat_id}:{sub_id}", [])
            for i, row in enumerate(rows):
                values = [row.get(c, "") for c in cols]
                tree.insert("", "end", iid=str(i), values=values)

    def _get_tree_for(self, cat_id, sub_id):
        for c, s, tree, cols in self.tab_info:
            if c == cat_id and s == sub_id:
                return tree
        return None

    def _add_record(self, cat_id, sub_id):
        floor = self._current_floor()
        if floor is None:
            messagebox.showinfo("Info", "Pilih bangunan & floor dahulu.")
            return
        fields = fields_for(cat_id, sub_id)
        dlg = RecordDialog(self, cat_id, sub_id, fields)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        key = f"{cat_id}:{sub_id}"
        floor["data"].setdefault(key, []).append(dlg.result)
        self._save()
        self._refresh_all_tabs()

    def _selected_record_index(self, cat_id, sub_id):
        tree = self._get_tree_for(cat_id, sub_id)
        sel = tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def _edit_record(self, cat_id, sub_id):
        floor = self._current_floor()
        if floor is None:
            return
        idx = self._selected_record_index(cat_id, sub_id)
        if idx is None:
            messagebox.showinfo("Info", "Pilih rekod dahulu.")
            return
        key = f"{cat_id}:{sub_id}"
        rows = floor["data"].get(key, [])
        fields = fields_for(cat_id, sub_id)
        dlg = RecordDialog(self, cat_id, sub_id, fields, initial=rows[idx])
        self.wait_window(dlg)
        if dlg.result is None:
            return
        rows[idx] = dlg.result
        self._save()
        self._refresh_all_tabs()

    def _delete_record(self, cat_id, sub_id):
        floor = self._current_floor()
        if floor is None:
            return
        idx = self._selected_record_index(cat_id, sub_id)
        if idx is None:
            messagebox.showinfo("Info", "Pilih rekod dahulu.")
            return
        if not messagebox.askyesno("Padam Rekod", "Padam rekod ini?"):
            return
        key = f"{cat_id}:{sub_id}"
        del floor["data"][key][idx]
        self._save()
        self._refresh_all_tabs()

    # ---------------- Export ----------------
    def _export_floor(self):
        b = self._current_building()
        floor = self._current_floor()
        if b is None or floor is None:
            messagebox.showinfo("Info", "Pilih bangunan & floor dahulu.")
            return
        wb = Workbook()
        wb.remove(wb.active)
        any_data = export_floor_to_workbook(
            wb, floor["data"], b["name"], floor["name"],
            project_name=self.project_var.get(), location=b.get("location", "")
        )
        if not any_data:
            messagebox.showinfo("Info", "Tiada data untuk floor ini lagi.")
            return
        default_name = f"{self.team_var.get() or 'Team'} - {b['name']} - {floor['name']}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        wb.save(path)
        messagebox.showinfo("Selesai", f"Fail Excel disimpan:\n{path}")

    def _export_all(self):
        if not self.data["buildings"]:
            messagebox.showinfo("Info", "Tiada bangunan lagi.")
            return
        wb = export_project_to_one_workbook(self.data["buildings"], project_name=self.project_var.get())
        default_name = f"{self.team_var.get() or 'Team'} - Compiled Overall Data.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        wb.save(path)
        messagebox.showinfo("Selesai", f"Fail Excel (compiled) disimpan:\n{path}")


def _ask_text(parent, title, prompt):
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("320x120")
    win.transient(parent)
    win.grab_set()
    ttk.Label(win, text=prompt).pack(pady=(14, 4))
    var = tk.StringVar()
    entry = ttk.Entry(win, textvariable=var, width=30)
    entry.pack(pady=4)
    entry.focus_set()
    result = {}

    def on_ok():
        result["value"] = var.get().strip()
        win.destroy()

    ttk.Button(win, text="OK", command=on_ok).pack(pady=8)
    win.bind("<Return>", lambda e: on_ok())
    parent.wait_window(win)
    return result.get("value")


def _ask_building_info(parent, initial_name="", initial_location=""):
    win = tk.Toplevel(parent)
    win.title("Bangunan")
    win.geometry("340x180")
    win.transient(parent)
    win.grab_set()

    ttk.Label(win, text="Nama bangunan:").pack(pady=(14, 2))
    name_var = tk.StringVar(value=initial_name)
    name_entry = ttk.Entry(win, textvariable=name_var, width=34)
    name_entry.pack(pady=2)
    name_entry.focus_set()

    ttk.Label(win, text="Location (alamat / lokasi):").pack(pady=(10, 2))
    loc_var = tk.StringVar(value=initial_location)
    ttk.Entry(win, textvariable=loc_var, width=34).pack(pady=2)

    result = {}

    def on_ok():
        name = name_var.get().strip()
        if not name:
            messagebox.showinfo("Info", "Nama bangunan diperlukan.")
            return
        result["value"] = {"name": name, "location": loc_var.get().strip()}
        win.destroy()

    ttk.Button(win, text="OK", command=on_ok).pack(pady=10)
    win.bind("<Return>", lambda e: on_ok())
    parent.wait_window(win)
    return result.get("value")


if __name__ == "__main__":
    app = TNBESApp()
    app.mainloop()
