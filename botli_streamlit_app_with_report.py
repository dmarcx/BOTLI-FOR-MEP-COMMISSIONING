import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import fitz
import os
import streamlit.components.v1 as components

# Load data files
@st.cache_data
def load_data():
    ag = pd.read_csv("Lighting_AboveGround.csv")
    bg = pd.read_csv("Lighting_BelowGround.csv")
    return ag, bg

above_ground, below_ground = load_data()

# דיבור במקום הקלדה
st.sidebar.header("🔊 זיהוי קולי במקום הקלדה")
with st.sidebar.expander("דבר והמערכת תזהה את הטקסט"):
    components.html("""
        <button onclick="startDictation()">התחל דיבור</button>
        <p id="output" style="direction: rtl; font-weight: bold;"></p>
        <script>
        function startDictation() {
            if (window.hasOwnProperty('webkitSpeechRecognition')) {
                var recognition = new webkitSpeechRecognition();
                recognition.continuous = false;
                recognition.interimResults = false;
                recognition.lang = "he-IL";
                recognition.start();
                recognition.onresult = function(e) {
                    document.getElementById('output').innerText = e.results[0][0].transcript;
                    recognition.stop();
                };
                recognition.onerror = function(e) {
                    recognition.stop();
                }
            } else {
                alert("הדפדפן שלך אינו תומך בזיהוי דיבור");
            }
        }
        </script>
    """, height=150)

# (שאר הקוד נשאר כמות שהוא...)
