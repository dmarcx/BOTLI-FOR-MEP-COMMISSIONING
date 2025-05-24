
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
    room_type = row.iloc[0].get("Type of Room", "").strip()
    return room_type, None if room_type else "Room type missing"

def check_documents(room):
    df = above_ground if room.startswith("L") else below_ground
    row = df[df["Room Number"].str.upper().str.strip() == room]
    if row.empty:
        return None
    return row.iloc[0].get("Documents Supplied", "").strip() == "כן"

def get_schedule_date(room):
    df = above_ground if room.startswith("L") else below_ground
    row = df[df["Room Number"].str.upper().str.strip() == room]
    if row.empty:
        return None, None, "Room not found"
    planned_date_str = row.iloc[0].get("Planned Date", "").strip()
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
    doc = fitz.open(file_name)
    text = "".join([page.get_text() for page in doc])
    return list(set(line.strip() for line in text.splitlines() if "EP-" in line and line.strip().startswith("EP-")))

def generate_report(room, room_type, planned, today, status, lux_result, sources):
    wb = load_workbook("דוח מסירה.xlsx")
    ws = wb.active
    ws["A1"] = f"{room} - {room_type}"
    ws["B3"] = str(planned)
    ws["B4"] = str(today)
    ws["C4"] = status
    ws["B6"] = "✓"
    ws["B7"] = "✓"
    ws["B8"] = "✓"
    ws["B22"] = "בדיקת תאורה"
    ws["C22"] = lux_result
    ws["B34"] = ", ".join(sources)
    output_path = f"report_{room}.xlsx"
    wb.save(output_path)
    return output_path

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
        st.markdown("""📏 **הנחיה:** מדוד את רמת ההארה במרכז החדר בגובה 80 ס"מ. ודא שאין אור חיצוני שמפריע.""")
                        measured = st.number_input("הזן את רמת ההארה שנמדדה (בלוקס):", min_value=0)
                        if measured:
                            lux_result = evaluate_lux(room_type, measured)
                            st.info(lux_result)
                            sources = get_power_sources(room)
                            st.write("מקורות אספקה שנמצאו בתוכנית:")
                            for s in sources:
                                st.write(f"🔌 {s}")
                            if st.checkbox("האם השילוט בפועל תואם לתכנון?"):
                                if st.checkbox("האם האור כבה לאחר הפעלת מאמ"ת?"):
                                    st.success("בדיקת התאורה הסתיימה בהצלחה.")
                                    if st.button("📄 הפק דו"ח מסירה"):
                                        file = generate_report(room, room_type, planned, today, status, lux_result, sources)
                                        with open(file, "rb") as f:
                                            st.download_button("📥 הורד את הדו"ח", data=f, file_name=file)
                                else:
                                    st.warning("נדרש לאמת את פעולת מאמ"ת.")
                            else:
                                st.warning("נדרש לתקן את השילוט או לעדכן את התכנון.")
                    else:
                        st.stop()
                else:
                    st.stop()
            else:
                st.error("נדרש אישור שכל המסמכים הוגשו. לא ניתן להמשיך.")
    else:
        st.error("הקלט שסופק אינו כולל אות אחת ואחריה 4 ספרות. לא ניתן להמשיך.")
