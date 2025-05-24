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
    return row.iloc[0].get("מסמכים סופקו", "").strip() == "כן" if not row.empty else None

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
            st.success(f"סוג חדר: {room_type} | מספר חדר: {room}")
            if check_documents(room):
                planned, today, status = get_schedule_date(room)
                st.info(f"התאריך המתוכנן הוא {planned}, היום {today} — הבדיקה {status}.")
                if st.checkbox("האם ניתן להתקדם לביצוע הבדיקה בפועל?"):
                    if st.checkbox("האם קיים מד תאורה זמין לביצוע הבדיקה?"):
                        fixtures = get_lighting_fixtures(room)
                        st.markdown("### 💡 בדיקת גופי תאורה")
                        for fx in fixtures:
                            st.info(fx)
                        match_fixtures = st.radio("האם אלו גופי התאורה והכמות הקיימים בפועל?", ("כן", "לא"))
                        if match_fixtures == "לא":
                            actual_fixtures = st.text_area("אנא הזן את סוגי הגופים והכמויות כפי שנמצאו בפועל (שורה לכל פריט)")
                            if actual_fixtures.strip():
                                fixtures = [line.strip() for line in actual_fixtures.splitlines() if line.strip()]
                                st.success("הרשימה עודכנה בהצלחה.")

                        st.markdown("### 🧑‍🤝‍🧑 מי המשתתפים בבדיקה ומה תפקידם?")
                        participants_text = st.text_area("אנא הזן רשימת משתתפים בפורמט שם – תפקיד, שורה לכל משתתף")
                        participants = []
                        if participants_text.strip():
                            participants = [line.strip() for line in participants_text.splitlines() if line.strip()]
                            more = st.radio("האם זו הרשימה המלאה?", ("כן", "לא"))
                            while more == "לא":
                                additional = st.text_area("הוסף משתתפים נוספים", key=f"more_{len(participants)}")
                                if additional.strip():
                                    participants += [line.strip() for line in additional.splitlines() if line.strip()]
                                    more = st.radio("האם כעת זו הרשימה המלאה?", ("כן", "לא"), key=f"confirm_{len(participants)}")
                        if not participants:
                            st.warning("נדרשת רשימת משתתפים להמשך.")
                            st.stop()

                        measured = st.number_input("הזן את רמת ההארה שנמדדה (בלוקס):", min_value=0)
                        if measured:
                            lux_result = evaluate_lux(room_type, measured)
                            st.info(lux_result)
                            dark_result = ""
                            if st.radio("האם קיימים אזורים חשוכים יותר בחדר?", ("לא", "כן")) == "כן":
                                dark_measure = st.number_input("הזן את רמת ההארה באזור החשוך (בלוקס):", min_value=0)
                                if dark_measure:
                                    dark_result = evaluate_lux(room_type, dark_measure)
                                    st.info("באזור החשוך: " + dark_result)

                        sources = get_power_sources(room)
                        st.markdown("### ⚡ מקורות אספקה שנמצאו בתוכנית:")
                        if sources:
                            for s in sources:
                                st.write(f"🔌 {s}")
                        else:
                            st.write("לא נמצאו מקורות אספקה.")

                        signage_match = st.checkbox("האם השילוט בפועל תואם לתכנון?")
                        breaker_test = st.radio("האם האור כבה לאחר הפלת המאמת?", ("כן", "לא"))

                        remarks = []
                        if not signage_match:
                            remarks.append("השילוט בפועל לא תואם לתכנון – נדרש עדכון.")
                        if breaker_test == "לא":
                            remarks.append("המאמת לא כיבה את האור – נדרש תיקון.")

                        if st.button("📄 הפק דו\"ח מסירה"):
                            file = generate_report(room, room_type, planned, today, status, lux_result, dark_result, sources, participants, remarks)
                            with open(file, "rb") as f:
                                st.download_button("📥 הורד את הדו\"ח", data=f, file_name=file)
                    else:
                        st.warning("נדרש מד תאורה לביצוע הבדיקה.")
            else:
                st.error("נדרש אישור שכל המסמכים הוגשו. לא ניתן להמשיך.")
    else:
        st.error("מספר החדר אינו תקני. יש להזין קלט במבנה L1234 או P1234.")
