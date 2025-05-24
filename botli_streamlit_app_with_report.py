import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import fitz
import os

# Load data files
@st.cache_data
def load_data():
    ag = pd.read_csv("Lighting_AboveGround.csv")
    bg = pd.read_csv("Lighting_BelowGround.csv")
    return ag, bg

above_ground, below_ground = load_data()

def get_room_type(room):
    df = above_ground if room.startswith("L") else below_ground
    row = df[df["Room Number"].str.upper().str.strip() == room]
    if row.empty:
        return None, "Room not found"
    room_type = row.iloc[0].get("Type of room", "").strip()
    return room_type, None if room_type else "Room type missing"

def check_documents(room):
    df = above_ground if room.startswith("L") else below_ground
    row = df[df["Room Number"].str.upper().str.strip() == room]
    if row.empty:
        return None
    return row.iloc[0].get("מסמכים סופקו", "").strip() == "כן"

def get_schedule_date(room):
    df = above_ground if room.startswith("L") else below_ground
    row = df[df["Room Number"].str.upper().str.strip() == room]
    if row.empty:
        return None, None, "Room not found"
    planned_date_str = row.iloc[0].get("Commissioning planned date", "").strip()
    if not planned_date_str:
        return None, None, "No planned date found"
    planned_date = datetime.strptime(planned_date_str, "%d-%b-%y").date()
    today = datetime.today().date()
    delta = (today - planned_date).days
    if delta == 0:
        status = "במועד"
    elif delta > 0:
        status = f"מאוחרת ב־{delta} ימים"
    else:
        status = f"מוקדמת ב־{abs(delta)} ימים"
    return planned_date, today, status

def evaluate_lux(room_type, measured_lux):
    lux_table = {
        "חדר ישיבות": 500,
        "מסדרון": 200,
        "משרד": 400,
        "חדר חשמל": 400,
        "חניון": 75,
        "חוץ": 30,
        "רמפה": 300
    }
    required = lux_table.get(room_type, 0)
    if measured_lux >= required:
        return "רמת ההארה תקינה."
    elif 0 < required - measured_lux <= 10:
        return "סטייה קלה – תירשם הערה לידיעת המתכנן."
    else:
        return "רמת ההארה אינה תקינה – נדרש תיקון או אישור המתכנן."

def get_power_sources(room):
    if room.startswith("L"):
        file_name = "SLD1-L3-EL-001.pdf"
    elif room.startswith("P"):
        file_name = f"SLD1-P{room[1]}-001.pdf"
    else:
        return []
    if not os.path.exists(file_name):
        return []
    doc = fitz.open(file_name)
    text = "".join([page.get_text() for page in doc])
    return list(set(line.strip() for line in text.splitlines() if "EP-" in line and line.strip().startswith("EP-")))

def generate_report(room, room_type, planned, today, status, lux_result, sources, participants, dark_measured=None, dark_measured=dark_measured if darker_area == "כן" else None):
    wb = load_workbook("דוח מסירה.xlsx")
    ws = wb.active
    ws["A1"] = f"{room} - {room_type}"
    ws["B3"] = str(planned)
    ws["B4"] = str(today)
    ws["C4"] = status
    ws["B6"] = "✓"
    ws["B7"] = "✓"
    ws["B8"] = "✓"
    ws["B9"] = participants.replace("–", ":").replace("\n", "; ")
    ws["B22"] = "בדיקת תאורה"
    ws["C22"] = lux_result
    ws["B34"] = ", ".join(sources)
    if dark_measured is not None:
        ws["C23"] = f"{dark_measured} לוקס באזור חשוך"
