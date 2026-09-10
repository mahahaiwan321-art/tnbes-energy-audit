"""
TNBES Energy Audit — konfigurasi kategori, field, dan cell-mapping template Excel.
Semua mapping cell di sini disalin terus daripada app web (app.html) supaya
output Excel Python ini IDENTIK dengan app web tersebut.
"""

# ---------------------------------------------------------------------------
# Kategori & sub-jenis (padan dengan CATEGORIES dalam app.html)
# ---------------------------------------------------------------------------
CATEGORIES = [
    {"id": "lighting", "name": "Lighting", "subTypes": [
        {"id": "main", "name": "Main Data"},
        {"id": "sampling", "name": "Sheet 2 — Sampling"},
    ]},
    {"id": "acmv", "name": "ACMV (Split Unit, AHU, Chiller)", "subTypes": [
        {"id": "split", "name": "Split Unit"},
        {"id": "ahu", "name": "AHU"},
        {"id": "chiller", "name": "Chiller / Pump"},
    ]},
    {"id": "office", "name": "Office & General Equipment", "subTypes": [
        {"id": "main", "name": "Main Data"},
    ]},
    {"id": "lift", "name": "Lift System", "subTypes": [
        {"id": "main", "name": "Main Data"},
    ]},
    {"id": "datalogger", "name": "Data Logger", "subTypes": [
        {"id": "main", "name": "Main Data"},
    ]},
]

def sub_types_for(cat_id):
    for c in CATEGORIES:
        if c["id"] == cat_id:
            return c["subTypes"]
    return []

def category_name(cat_id):
    for c in CATEGORIES:
        if c["id"] == cat_id:
            return c["name"]
    return cat_id

def subtype_name(cat_id, sub_id):
    for s in sub_types_for(cat_id):
        if s["id"] == sub_id:
            return s["name"]
    return sub_id


# ---------------------------------------------------------------------------
# Maps a category id -> TEMPLATES key
# ---------------------------------------------------------------------------
def template_key_for(cat_id, sub_id):
    if cat_id == "acmv":
        if sub_id == "split": return "acmv_split"
        if sub_id == "ahu": return "acmv_ahu"
        if sub_id == "chiller": return "acmv_chiller"
    if cat_id == "lighting":
        if sub_id == "sampling": return "lighting_sampling"
        return "lighting"
    if cat_id == "lift": return "lift"
    return cat_id  # office, datalogger


# ---------------------------------------------------------------------------
# FIELDS — senarai field bagi setiap kategori:sub, dalam urutan yang sama
# dengan template Excel
# ---------------------------------------------------------------------------
FIELDS = {
    "lighting:main": [
        "Level","Room / Area","Lighting Equipment","No. of Lamps","Type of Lamps",
        "Watts (Lamp)","No of Fitting","No of Broken Lamp","Type of Control","Status Reflection",
        "Weekdays Start","Weekdays End","Sat & Sun Start","Sat & Sun End",
        "Lux 1","Lux 2","Lux 3","Lux 4","Lux 5",
        "Temperature (°C)","Relative Humidity (%)","PPM","Type of MS Standard",
    ],
    "lighting:sampling": [
        "Level","Room / Area","No. of Lamps","Type of Lighting","Voltage (Measured)",
        "Current (Measured)","Watt (Measured)","Lux Level (Measured)","Power Factor",
    ],
    "acmv:split": [
        "Level","Room / Area","No of Units","Horsepower (HP)","Brand","Type of AC","Model",
        "Refrigerant","Cooling Capacity (Btu/hr)","Total AC Load (kW)","Temperature (°C)",
        "Relative Humidity (%)","CO2 (ppm)","Weekdays Start","Weekdays End","Sat & Sun Start",
        "Sat & Sun End","Hours/Day","Hours/Week","Savings %",
    ],
    "acmv:ahu": [
        "Equipment","Brand","Model","Horsepower (HP)","Rated Input Power (kW)","Rated CC (btu/hr or kW)",
        "Zone","Level","AHU Tag No.","Area Served / Description",
        "Press Drop Filter 1 (mbar)","Press Drop Filter 2 (mbar)","Press Drop Filter 3 (mbar)",
        "Press Drop Coil (mbar)","Press Drop Overall (mbar)",
        "AHU Fan VSD status (%)","AHU Fan Speed (RPM)",
        "Outside Air (OA) - Temp/RH/CO2",
        "Mixture Air (Mix) Temp (°C)","Mixture Air (Mix) RH (%)",
        "Return Air (RA) Temp (°C)","Return Air (RA) RH (%)","Return Air (RA) CO2 (ppm)",
        "Exhaust Air Temp (°C)","Exhaust Air RH (%)",
        "Supply Air (SA) Temp (°C)","Supply Air (SA) RH (%)","Supply Air (SA) Press. (mbar)",
        "Set Points RA (°C)","Set Points SA (°C)",
        "Mix Air Duct Distance - Mix (cm)","Mix Air Duct Velocity 1 - Mix (m/s)",
        "Mix Air Duct Velocity 2 - Mix (m/s)","Mix Air Duct Velocity 3 - Mix (m/s)","Mix Air Duct Velocity 4 - Mix (m/s)",
        "Mix Air Duct Distance - Supply (cm)","Mix Air Duct Velocity 1 - Supply (m/s)",
        "Mix Air Duct Velocity 2 - Supply (m/s)","Mix Air Duct Velocity 3 - Supply (m/s)","Mix Air Duct Velocity 4 - Supply (m/s)",
        "Mix Air Duct Width (m)","Mix Air Duct Height (m)",
        "Exhaust Air Duct Distance - R1 (cm)","Exhaust Air Duct Velocity 1 - R1 (m/s)",
        "Exhaust Air Duct Velocity 2 - R1 (m/s)","Exhaust Air Duct Velocity 3 - R1 (m/s)","Exhaust Air Duct Velocity 4 - R1 (m/s)",
        "Exhaust Air Duct Distance - R2 (cm)","Exhaust Air Duct Velocity 1 - R2 (m/s)",
        "Exhaust Air Duct Velocity 2 - R2 (m/s)","Exhaust Air Duct Velocity 3 - R2 (m/s)","Exhaust Air Duct Velocity 4 - R2 (m/s)",
        "Exhaust Air Duct Distance - R3 (cm)","Exhaust Air Duct Velocity 1 - R3 (m/s)",
        "Exhaust Air Duct Velocity 2 - R3 (m/s)","Exhaust Air Duct Velocity 3 - R3 (m/s)","Exhaust Air Duct Velocity 4 - R3 (m/s)",
        "Exhaust Air Duct Width (m)","Exhaust Air Duct Height (m)",
        "Return Air Duct Distance - R1 (cm)","Return Air Duct Velocity 1 - R1 (m/s)",
        "Return Air Duct Velocity 2 - R1 (m/s)","Return Air Duct Velocity 3 - R1 (m/s)","Return Air Duct Velocity 4 - R1 (m/s)",
        "Return Air Duct Distance - R2 (cm)","Return Air Duct Velocity 1 - R2 (m/s)",
        "Return Air Duct Velocity 2 - R2 (m/s)","Return Air Duct Velocity 3 - R2 (m/s)","Return Air Duct Velocity 4 - R2 (m/s)",
        "Return Air Duct Distance - R3 (cm)","Return Air Duct Velocity 1 - R3 (m/s)",
        "Return Air Duct Velocity 2 - R3 (m/s)","Return Air Duct Velocity 3 - R3 (m/s)","Return Air Duct Velocity 4 - R3 (m/s)",
        "Return Air Duct Width (m)","Return Air Duct Height (m)",
        "Outside Air Duct Distance - R1 (cm)","Outside Air Duct Velocity 1 - R1 (m/s)",
        "Outside Air Duct Velocity 2 - R1 (m/s)","Outside Air Duct Velocity 3 - R1 (m/s)","Outside Air Duct Velocity 4 - R1 (m/s)",
        "Outside Air Duct Distance - R2 (cm)","Outside Air Duct Velocity 1 - R2 (m/s)",
        "Outside Air Duct Velocity 2 - R2 (m/s)","Outside Air Duct Velocity 3 - R2 (m/s)","Outside Air Duct Velocity 4 - R2 (m/s)",
        "Outside Air Duct Distance - R3 (cm)","Outside Air Duct Velocity 1 - R3 (m/s)",
        "Outside Air Duct Velocity 2 - R3 (m/s)","Outside Air Duct Velocity 3 - R3 (m/s)","Outside Air Duct Velocity 4 - R3 (m/s)",
        "Outside Air Duct Width (m)","Outside Air Duct Height (m)",
        "Total Load DB Current R/Vry (A)","Total Load DB Cos phi R/Vry","Total Load DB Voltage R/Vry (V)",
        "Total Load DB Current Y/Vrb (A)","Total Load DB Cos phi Y/Vrb","Total Load DB Voltage Y/Vrb (V)",
        "Total Load DB Current B/Vyb (A)","Total Load DB Cos phi B/Vyb","Total Load DB Voltage B/Vyb (V)",
        "AHU Fan Motor DB Current R/Vry (A)","AHU Fan Motor DB Cos phi R/Vry","AHU Fan Motor DB Voltage R/Vry (V)",
        "AHU Fan Motor DB Current Y/Vrb (A)","AHU Fan Motor DB Cos phi Y/Vrb","AHU Fan Motor DB Voltage Y/Vrb (V)",
        "AHU Fan Motor DB Current B/Vyb (A)","AHU Fan Motor DB Cos phi B/Vyb","AHU Fan Motor DB Voltage B/Vyb (V)",
        "BAS RA Temp (°C)","BAS Valve Opening (%)",
        "Chilled Water Valve Opening (%)","Chilled Water Temp In/Out (°C)","Chilled Water SA Pressure (Pa)","Chilled Water SA CO2 (ppm)",
        "Remarks",
    ],
    "acmv:chiller": [
        "Pump ID / Tag No.","Building / Location","Pump Application","Pump Type","Manufacturer","Model",
        "Installation Year","Operating Schedule","Control Method","Rated Flow Rate (m³/h)","Rated Head (m)",
        "Motor Rated Power (kW)","Motor Rated RPM","Motor Efficiency (%)",
        "Date","Time","Operating Condition","Voltage (V)","Current (A)","Power (kW)","Power Factor",
        "Frequency (Hz)","RPM","Flow Rate (m³/h)","Suction Pressure","Discharge Pressure","Remarks",
    ],
    "office:main": [
        "Level","Room / Area","Type of Equipment","Brand","No of Unit","Voltage (V)","Current (A)","Power (W)",
        "Manual / Auto","Operating Time","Running (W)","Standby (W)","kW/hours","Hours/Day","Hours/Week",
        "Weekly Energy (kWh/week)",
    ],
    "lift:main": [
        "Lift ID / Tag No.",
        "Building / Location","Manufacturer","Model","Installation Year","Rated Capacity (kg/persons)",
        "No. of Floors Served","Rated Speed (m/s)","Motor/Drive Rated Power (kW)","Drive Type","Operating Schedule",
        "Date","Start Time","End Time","Duration (min)","Logging Interval (sec)",
        "Voltage (V)","Current (A)","Power (kW)","Energy (kWh)","Power Factor",
        "Time","Lift Condition","Direction","Load Condition","Power (kW) - Profile","Remarks",
    ],
    "datalogger:main": [
        "Logger ID","Brand / Model","Serial No.","Calibration Due Date","Installation Location",
        "Installation Date & Time","Dismantling Date & Time","Logging Interval","Logger Installed By","Remarks",
    ],
}

def fields_for(cat_id, sub_id):
    return FIELDS.get(f"{cat_id}:{sub_id}", [])


# ---------------------------------------------------------------------------
# Dropdown options
# ---------------------------------------------------------------------------
LIGHTING_EQUIPMENT_OPTIONS = [
    "(SM) Surface Mounted","(CR) Ceiling Recessed","(BC) Bare Channel","(SP) Sportlight / Spotlight",
    "(DL) Downlight","(DC) Decorative Light","(IL) Inspection Light","(SL) Surgical Light",
    "(EE) Emergency Exit Light","(EL) Emergency Light","(HB) High Bay Light","(LB) Low Bay Light",
    "(TL) Track Light","(FDL) Floodlight","(ST) Street Light / Lantern","(BL) Bollard Light",
    "(WL) Wall Light / Wall Sconce","(PN) Pendant Light","(BH) Bulkhead Light",
    "(STL) Strip Light / Tape Light","(IGL) Inground / Buried Light","(UWL) Underwater Light","(SLR) Solar Light",
    "Other",
]

TYPE_OF_LAMP_OPTIONS = [
    "(I) Incandescent Lamp","(FL) Fluorescent Lamp","(CFL) Compact Fluorescent Lamp",
    "(LED) Light Emitting Diode","(PL) Plasma Lamp","(HG) Halogen Lamp","(MH) Metal Halide Lamp",
    "(HPS / SON) High-Pressure Sodium Lamp","(LPS / SOX) Low-Pressure Sodium Lamp","(MV) Mercury Vapor Lamp",
    "(IND) Induction Lamp","(OLED) Organic Light Emitting Diode","(SMART) Smart LED / RGBW LED","(O) Others",
]

MS_STANDARD_OPTIONS = [
    "Interior walkway and car-park","Hotel bedroom","Lift interior",
    "Corridor","Passageways","Stairs",
    "Escalator","Travellator","Entrance and exit",
    "Staff changing room","Locker and cleaner room","Cloak room","Lavatories","Stores",
    "Entrance hall","Lobbies","Waiting room","Inquiry desk","Gate house",
    "Infrequent reading and writing",
    "General offices","Shops and stores","Reading and writing",
    "Drawing office","Restroom",
    "Restaurant","Canteen","Cafeteria","Kitchen","Lounge","Bathroom","Toilet","Bedroom",
    "Class room","Library",
    "Shop/supermarket/department store","Museum and gallery","Proof reading",
    "Exacting drawing","Detailed and precise work",
    "Other",
]

TYPE_OF_CONTROL_OPTIONS = ["", "Manual", "Auto", "Other"]
STATUS_REFLECTION_OPTIONS = ["", "Yes", "No", "Other"]
TYPE_OF_AC_OPTIONS = ["", "Wall Mounted", "Cassette", "Ducted", "Floor Standing", "Other"]
CONDITION_OPTIONS = ["", "Good", "Fair", "Poor", "Faulty/Not Working", "Other"]

# Nilai dalam dropdown yang bermaksud "sila taip sendiri" -> bila dipilih,
# GUI akan buka satu kotak teks tambahan utk custom input.
OTHER_TRIGGER_VALUES = {"Other", "Others", "(O) Others"}

# Field -> dropdown option list (dipakai oleh GUI untuk render Combobox)
FIELD_DROPDOWNS = {
    "Lighting Equipment": LIGHTING_EQUIPMENT_OPTIONS,
    "Type of Lamps": TYPE_OF_LAMP_OPTIONS,
    "Type of MS Standard": MS_STANDARD_OPTIONS,
    "Type of Control": TYPE_OF_CONTROL_OPTIONS,
    "Status Reflection": STATUS_REFLECTION_OPTIONS,
    "Type of AC": TYPE_OF_AC_OPTIONS,
    "Manual / Auto": TYPE_OF_CONTROL_OPTIONS,
    "Operating Condition": CONDITION_OPTIONS,
}

# Field -> "numeric" (keypad nombor sahaja)
NUMERIC_ONLY_FIELDS = {
    "No. of Lamps","No of Fitting","No of Broken Lamp",
    "Lux 1","Lux 2","Lux 3","Lux 4","Lux 5",
    "Temperature (°C)","Relative Humidity (%)","PPM",
}

# Field -> "time" (keypad masa 24 jam, format HH:MM)
TIME_FIELDS = {"Weekdays Start","Weekdays End","Sat & Sun Start","Sat & Sun End"}


# ---------------------------------------------------------------------------
# TEMPLATE_LAYOUT — cell-mapping tepat bagi setiap kategori:sub
# ---------------------------------------------------------------------------
TEMPLATE_LAYOUT = {
    "lighting:main": {
        "mode": "table",
        "start_row": 13,
        "columns": ["A","B","C","D","E","F","G","H","J","K","L","M","N","O",
                    "R","S","T","U","V","X","Y","Z","AA"],
    },
    "lighting:sampling": {
        "mode": "table",
        "start_row": 13,
        "columns": ["B","C","D","E","F","G","H","I","J"],
        "serial_column": "A",
    },
    "acmv:split": {
        "mode": "table",
        "start_row": 13,
        "columns": ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","W"],
    },
    "acmv:chiller": {
        # Sheet ada 2 jadual berasingan (General Info + Measurement Log),
        # dipaut ikut "Pump ID / Tag No." (lajur A pada kedua-dua jadual).
        # data_count = bilangan field SELAIN link_field bagi jadual tersebut.
        "mode": "linked-tables",
        "link_field": "Pump ID / Tag No.",
        "tables": [
            {"start_row": 9,  "columns": ["A","B","C","D","E","F","G","H","I","J","K","L","M","N"], "data_count": 13},
            {"start_row": 27, "columns": ["A","B","C","D","E","F","G","H","I","J","K","L","M","N"], "data_count": 13},
        ],
    },
    "office:main": {
        "mode": "table",
        "start_row": 9,
        "columns": ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"],
    },
    "lift:main": {
        # Sheet ada 3 jadual berasingan (General Info + Measurement Log +
        # Operating Profile), dipaut ikut "Lift ID / Tag No." (lajur A semua jadual).
        "mode": "linked-tables",
        "link_field": "Lift ID / Tag No.",
        "tables": [
            {"start_row": 6,  "columns": ["A","B","C","D","E","F","G","H","I","J","K"], "data_count": 10},
            {"start_row": 27, "columns": ["A","B","C","D","E","F","G","H","I","J","K"], "data_count": 10},
            {"start_row": 45, "columns": ["A","B","C","D","E","F","G"], "data_count": 6},
        ],
    },
    "datalogger:main": {
        "mode": "table",
        "start_row": 7,
        "columns": ["A","B","C","D","E","F","G","H","I","J"],
    },
    "acmv:ahu": {
        # Satu AHU = SATU sheet penuh (klon template). Kalau > 1 AHU dalam
        # floor yang sama, setiap satu jadi sheet berasingan.
        "mode": "grid-per-entry",
        "cells": {
            "Equipment": "A10", "Brand": "C10", "Model": "E10", "Horsepower (HP)": "H10",
            "Rated Input Power (kW)": "K10", "Rated CC (btu/hr or kW)": "O10",
            "Zone": "S10", "Level": "U10", "AHU Tag No.": "W10", "Area Served / Description": "S12",
            "Press Drop Filter 1 (mbar)": "A18", "Press Drop Filter 2 (mbar)": "B18", "Press Drop Filter 3 (mbar)": "C18",
            "Press Drop Coil (mbar)": "D18", "Press Drop Overall (mbar)": "E18",
            "AHU Fan VSD status (%)": "F18", "AHU Fan Speed (RPM)": "H18",
            "Outside Air (OA) - Temp/RH/CO2": "I18",
            "Mixture Air (Mix) Temp (°C)": "M18", "Mixture Air (Mix) RH (%)": "N18",
            "Return Air (RA) Temp (°C)": "P18", "Return Air (RA) RH (%)": "Q18", "Return Air (RA) CO2 (ppm)": "R18",
            "Exhaust Air Temp (°C)": "S18", "Exhaust Air RH (%)": "T18",
            "Supply Air (SA) Temp (°C)": "V18", "Supply Air (SA) RH (%)": "W18", "Supply Air (SA) Press. (mbar)": "X18",
            "Set Points RA (°C)": "Y18", "Set Points SA (°C)": "Z18",

            "Mix Air Duct Distance - Mix (cm)": "B22", "Mix Air Duct Velocity 1 - Mix (m/s)": "C22",
            "Mix Air Duct Velocity 2 - Mix (m/s)": "D22", "Mix Air Duct Velocity 3 - Mix (m/s)": "E22",
            "Mix Air Duct Velocity 4 - Mix (m/s)": "F22",
            "Mix Air Duct Distance - Supply (cm)": "B24", "Mix Air Duct Velocity 1 - Supply (m/s)": "C24",
            "Mix Air Duct Velocity 2 - Supply (m/s)": "D24", "Mix Air Duct Velocity 3 - Supply (m/s)": "E24",
            "Mix Air Duct Velocity 4 - Supply (m/s)": "F24",
            "Mix Air Duct Width (m)": "C26", "Mix Air Duct Height (m)": "E26",

            "Exhaust Air Duct Distance - R1 (cm)": "I22", "Exhaust Air Duct Velocity 1 - R1 (m/s)": "J22",
            "Exhaust Air Duct Velocity 2 - R1 (m/s)": "K22", "Exhaust Air Duct Velocity 3 - R1 (m/s)": "L22",
            "Exhaust Air Duct Velocity 4 - R1 (m/s)": "M22",
            "Exhaust Air Duct Distance - R2 (cm)": "I23", "Exhaust Air Duct Velocity 1 - R2 (m/s)": "J23",
            "Exhaust Air Duct Velocity 2 - R2 (m/s)": "K23", "Exhaust Air Duct Velocity 3 - R2 (m/s)": "L23",
            "Exhaust Air Duct Velocity 4 - R2 (m/s)": "M23",
            "Exhaust Air Duct Distance - R3 (cm)": "I24", "Exhaust Air Duct Velocity 1 - R3 (m/s)": "J24",
            "Exhaust Air Duct Velocity 2 - R3 (m/s)": "K24", "Exhaust Air Duct Velocity 3 - R3 (m/s)": "L24",
            "Exhaust Air Duct Velocity 4 - R3 (m/s)": "M24",
            "Exhaust Air Duct Width (m)": "J26", "Exhaust Air Duct Height (m)": "L26",

            "Return Air Duct Distance - R1 (cm)": "B29", "Return Air Duct Velocity 1 - R1 (m/s)": "C29",
            "Return Air Duct Velocity 2 - R1 (m/s)": "D29", "Return Air Duct Velocity 3 - R1 (m/s)": "E29",
            "Return Air Duct Velocity 4 - R1 (m/s)": "F29",
            "Return Air Duct Distance - R2 (cm)": "B30", "Return Air Duct Velocity 1 - R2 (m/s)": "C30",
            "Return Air Duct Velocity 2 - R2 (m/s)": "D30", "Return Air Duct Velocity 3 - R2 (m/s)": "E30",
            "Return Air Duct Velocity 4 - R2 (m/s)": "F30",
            "Return Air Duct Distance - R3 (cm)": "B31", "Return Air Duct Velocity 1 - R3 (m/s)": "C31",
            "Return Air Duct Velocity 2 - R3 (m/s)": "D31", "Return Air Duct Velocity 3 - R3 (m/s)": "E31",
            "Return Air Duct Velocity 4 - R3 (m/s)": "F31",
            "Return Air Duct Width (m)": "C33", "Return Air Duct Height (m)": "E33",

            "Outside Air Duct Distance - R1 (cm)": "I29", "Outside Air Duct Velocity 1 - R1 (m/s)": "J29",
            "Outside Air Duct Velocity 2 - R1 (m/s)": "K29", "Outside Air Duct Velocity 3 - R1 (m/s)": "L29",
            "Outside Air Duct Velocity 4 - R1 (m/s)": "M29",
            "Outside Air Duct Distance - R2 (cm)": "I30", "Outside Air Duct Velocity 1 - R2 (m/s)": "J30",
            "Outside Air Duct Velocity 2 - R2 (m/s)": "K30", "Outside Air Duct Velocity 3 - R2 (m/s)": "L30",
            "Outside Air Duct Velocity 4 - R2 (m/s)": "M30",
            "Outside Air Duct Distance - R3 (cm)": "I31", "Outside Air Duct Velocity 1 - R3 (m/s)": "J31",
            "Outside Air Duct Velocity 2 - R3 (m/s)": "K31", "Outside Air Duct Velocity 3 - R3 (m/s)": "L31",
            "Outside Air Duct Velocity 4 - R3 (m/s)": "M31",
            "Outside Air Duct Width (m)": "J33", "Outside Air Duct Height (m)": "L33",

            "Total Load DB Current R/Vry (A)": "Q22", "Total Load DB Cos phi R/Vry": "R22", "Total Load DB Voltage R/Vry (V)": "S22",
            "Total Load DB Current Y/Vrb (A)": "Q23", "Total Load DB Cos phi Y/Vrb": "R23", "Total Load DB Voltage Y/Vrb (V)": "S23",
            "Total Load DB Current B/Vyb (A)": "Q24", "Total Load DB Cos phi B/Vyb": "R24", "Total Load DB Voltage B/Vyb (V)": "S24",

            "AHU Fan Motor DB Current R/Vry (A)": "Q28", "AHU Fan Motor DB Cos phi R/Vry": "R28", "AHU Fan Motor DB Voltage R/Vry (V)": "S28",
            "AHU Fan Motor DB Current Y/Vrb (A)": "Q29", "AHU Fan Motor DB Cos phi Y/Vrb": "R29", "AHU Fan Motor DB Voltage Y/Vrb (V)": "S29",
            "AHU Fan Motor DB Current B/Vyb (A)": "Q30", "AHU Fan Motor DB Cos phi B/Vyb": "R30", "AHU Fan Motor DB Voltage B/Vyb (V)": "S30",

            "BAS RA Temp (°C)": "P33", "BAS Valve Opening (%)": "R33",

            "Chilled Water Valve Opening (%)": "V24", "Chilled Water Temp In/Out (°C)": "W24",
            "Chilled Water SA Pressure (Pa)": "Y24", "Chilled Water SA CO2 (ppm)": "Z24",
            "Remarks": "V27",
        },
    },
}


# ---------------------------------------------------------------------------
# HEADER_CELLS — cell utk tulis Project / Building / Location (dipaparkan di
# bahagian atas kebanyakan sheet). Key ialah TEMPLATES key (bukan cat:sub),
# sebab header ni milik template/sheet itu sendiri, bukan kategori data.
# Sheet yang tiada dalam dict ni (contoh "lift") memang tak ada header
# berasingan - info tersebut sudah tertanam sebagai field biasa dlm jadual.
# ---------------------------------------------------------------------------
HEADER_CELLS = {
    "lighting":          {"Project": "B5", "Building": "B6", "Location": "B7"},
    "lighting_sampling": {"Project": "C5", "Building": "C6", "Location": "C7"},
    "acmv_split":        {"Project": "B5", "Building": "B6", "Location": "B7"},
    "acmv_ahu":          {"Project": "D5", "Building": "D6", "Location": "D7"},
    "acmv_chiller":      {"Project": "B4", "Building": "D4", "Location": "F4"},
    "office":            {"Project": "B3", "Building": "B4"},  # tiada field Location dlm template ni
    "datalogger":        {"Project": "B4", "Building": "D4", "Location": "F4"},
}

