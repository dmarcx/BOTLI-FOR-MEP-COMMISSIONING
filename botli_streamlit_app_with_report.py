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

    # Normalize known room type synonyms
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

# Streamlit UI
st.title("BOTLI – בדיקת תאורה")

room = st.text_input("הזן מספר חדר (לדוגמה L3001):")
if room:
    room = room.upper().strip()
    if len(room) == 5 and room[0] in ["L", "P"] and room[1:].isdigit():
        room_type, error = get_room_type(room)
        if error:
            st.error(error)
        else:
            st.success(f"הבדיקה מתבצעת על חדר מספר {room} מסוג {room_type}.")
            if check_documents(room):
                st.info("כל המסמכים הוגשו.")
                planned, today, status = get_schedule_date(room)
                st.write(f"התאריך המתוכנן הוא {planned}, היום {today} — הבדיקה {status}.")
                if st.checkbox("האם ניתן להתקדם לביצוע הבדיקה בפועל?"):
                    if st.checkbox("האם קיים מד תאורה זמין לביצוע הבדיקה?"):
                        measured = st.number_input("הזן את רמת ההארה שנמדדה (בלוקס):", min_value=0)
                        if measured:
                            lux_result = evaluate_lux(room_type, measured)
                            st.info(lux_result)
                    else:
                        st.warning("נדרש מד תאורה לביצוע הבדיקה.")
                else:
                    st.warning("לא ניתן להמשיך ללא אישור התקדמות.")
            else:
                st.error("נדרש אישור שכל המסמכים הוגשו. לא ניתן להמשיך.")
    else:
        st.error("מספר החדר אינו תקני. יש להזין קלט במבנה L1234 או P1234.")
