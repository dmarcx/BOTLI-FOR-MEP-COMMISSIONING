import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import fitz
import os

@st.cache_data
def load_data():
    ag = pd.read_csv("Lighting_AboveGround.csv")
    bg = pd.read_csv("Lighting_BelowGround.csv")
    return ag, bg

above_ground, below_ground = load_data()

# פונקציות עזר (נשמרו כמו קודם)
# ... (כל הפונקציות שמעל נשארות זהות, לא משוכפלות כאן לקיצור)

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
                st.info("✔️ כל המסמכים הוגשו.")
            else:
                st.error("✖️ נדרש אישור שכל המסמכים הוגשו. לא ניתן להמשיך.")
                st.stop()

            planned, today, status = get_schedule_date(room)
            st.info(f"📅 התאריך המתוכנן הוא {planned}, היום {today}, הבדיקה {status}.")

            if st.checkbox("האם ניתן להתקדם לביצוע הבדיקה בפועל?"):
                if st.checkbox("האם קיים מד תאורה זמין לביצוע הבדיקה?"):
                    # בדיקת גופי תאורה
                    st.subheader("💡 בדיקת גופי תאורה")
                    fixtures = get_lighting_fixtures(room)
                    for fix in fixtures:
                        st.info(fix)
                    match = st.radio("האם אלו גופי התאורה והכמות הקיימים בפועל?", ["כן", "לא"])
                    remarks = []
                    if match == "לא":
                        actual = st.text_area("אנא הזן את סוגי הגופים והכמויות כפי שנמצאו בפועל (שורה לכל פריט)")
                        if actual:
                            remarks += actual.splitlines()
                            st.success("הרשימה עודכנה בהצלחה.")

                    # משתתפים
                    st.subheader("👬 מי המשתתפים בבדיקה ומה תפקידם?")
                    participants_text = st.text_area("אנא הזן רשימת משתתפים בפורמט שם – תפקיד, שורה לכל משתתף")
                    participants = [line.strip() for line in participants_text.splitlines() if line.strip()]
                    if not participants:
                        st.warning("נדרשת רשימת משתתפים להמשך.")
                        st.stop()

                    # מדידת עוצמת הארה
                    measured = st.number_input("הזן את רמת ההארה שנמדדה (בלוקס):", min_value=0)
                    if measured:
                        lux_result = evaluate_lux(room_type, measured)
                        st.info(lux_result)

                        dark_result = ""
                        darker_area = st.radio("האם קיימים אזורים חשוכים יותר בחדר?", ("לא", "כן"))
                        if darker_area == "כן":
                            dark_measure = st.number_input("הזן את רמת ההארה באזור החשוך (בלוקס):", min_value=0)
                            if dark_measure:
                                dark_result = evaluate_lux(room_type, dark_measure)
                                st.info("באזור החשוך: " + dark_result)

                        # מקורות אספקה
                        sources = get_power_sources(room)
                        st.markdown("### ⚡ מקורות אספקה שנמצאו בתוכנית:")
                        if sources:
                            for s in sources:
                                st.write(f"🔌 {s}")
                        else:
                            st.write("לא נמצאו מקורות אספקה.")

                        # שילוט
                        signage_match = st.checkbox("האם השילוט בפועל תואם לתכנון?")
                        if not signage_match:
                            remarks.append("שילוט לא תואם – נדרש תיקון או עדכון תכנון.")

                        # מאמת
                        breaker_test = st.radio("האם האור כבה לאחר הפלת המאמת?", ("כן", "לא"))
                        if breaker_test == "לא":
                            remarks.append("נדרש לאמת את פעולת המאמת – האור לא כבה לאחר הפלתו.")

                        # הפקת הדו"ח
                        if st.button("📄 הפק דו"ח מסירה"):
                            file = generate_report(room, room_type, planned, today, status, lux_result, dark_result, sources, participants, remarks)
                            with open(file, "rb") as f:
                                st.download_button("📥 הורד את הדו"ח", data=f, file_name=file)
                    else:
                        st.warning("נדרש להזין ערך מדוד כדי להמשיך.")
                else:
                    st.warning("נדרש מד תאורה לביצוע הבדיקה.")
            else:
                st.warning("לא ניתן להמשיך ללא אישור התקדמות.")
    else:
        st.error("מספר החדר אינו תקני. יש להזין קלט במבנה L1234 או P1234.")
