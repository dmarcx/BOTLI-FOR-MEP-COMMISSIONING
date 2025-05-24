import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import fitz
import os
from streamlit_webrtc import webrtc_streamer, AudioProcessorBase, ClientSettings
import speech_recognition as sr
import av

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

    synonyms = {
        "MEETING ROOM": "חדר ישיבות",
        "MEETINGROOM": "חדר ישיבות",
        "OFFICE": "משרד",
        "CORRIDOR": "מסדרון",
        "ELECTRICAL ROOM": "חדר חשמל",
        "PARKING": "חניון",
        "OUTDOOR": "חוץ",
        "RAMP": "רמפה"
    }
    room_type = synonyms.get(room_type.upper().replace(" ", ""), room_type)

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
        status = f"מאוחרת ב‏{delta} ימים"
    else:
        status = f"מוקדמת ב‏{abs(delta)} ימים"
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
    required = lux_table.get(room_type)

    if required is None:
        st.warning(f"אזהרה: לא נמצאה דרישת לוקס עבור סוג החדר '{room_type}' — ייתכן שהערך שגוי או לא נתמך.")
        return "סוג חדר לא מזוהה – דרישות התאורה אינן ידועות."

    deviation = measured_lux - required

    if measured_lux == 0:
        return "לא נמדדה עוצמת הארה – נדרש להזין ערך."
    elif deviation >= 0:
        return "רמת ההארה תקינה."
    elif -10 <= deviation < 0:
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

def get_lighting_fixtures(room):
    df = above_ground if room.startswith("L") else below_ground
    rows = df[df["Room Number"].str.upper().str.strip() == room]
    if rows.empty:
        return ["לא נמצאו נתונים"]
    fixtures = []
    for _, row in rows.iterrows():
        type_c = str(row.get("Type", "")).strip()
        model_d = str(row.get("Model", "")).strip()
        quantity_e = str(row.get("Quantity", "")).strip()
        if type_c or model_d:
            fixtures.append(f"{type_c} {model_d} – כמות: {quantity_e}")
    return fixtures if fixtures else ["לא צויין"]

def generate_report(room, room_type, planned, today, status, lux_result, dark_result, sources, participants, remarks):
    wb = load_workbook("דוח מסירה.xlsx")
    ws = wb.active
    ws["A1"] = f"חדר {room} ({room_type})"
    ws["B3"] = str(planned)
    ws["B4"] = str(today)
    ws["C4"] = status
    ws["B6"] = lux_result
    ws["B7"] = dark_result if dark_result else "לא נמדד"
    ws["B8"] = ", ".join(sources)
    for i, p in enumerate(participants, start=12):
        ws[f"B{i}"] = p
    start_row = 22
    for i, remark in enumerate(remarks):
        ws[f"B{start_row + i}"] = remark
    file_name = f"report_{room}.xlsx"
    wb.save(file_name)
    return file_name

# Audio recognition using WebRTC
st.sidebar.header("🔊 דיבור לזיהוי טקסט")
with st.sidebar.expander("הקלד או דבר במקום להקליד"):
    st.markdown("כאן תוכל להשתמש ברכיב זיהוי דיבור (דורש הפעלת מיקרופון בדפדפן)")
    webrtc_ctx = webrtc_streamer(
        key="speech-to-text",
        client_settings=ClientSettings(
            media_stream_constraints={"audio": True, "video": False},
            rtc_configuration={"iceServers": [{"urls": ["stun:stun.l.google.com:19302"]}]},
        ),
        audio_receiver_size=1024,
        translations=None
    )
    st.caption("הערה: לצורך זיהוי הדיבור המלא נדרש לנתח את הקלט בצד שרת או להשתמש ב־API חיצוני")
