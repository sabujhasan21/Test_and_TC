# testimonial_tc_streamlit.py
import os
import shutil
import pandas as pd
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont, TTFError
import streamlit as st

# ----------------------------
# Font Setup
# ----------------------------
FONT_PATH = os.path.join(os.path.dirname(__file__), "fonts", "SolaimanLipi.ttf")
BANG_FONT = "Times-Roman"
try:
    if os.path.exists(FONT_PATH):
        pdfmetrics.registerFont(TTFont("Bangla", FONT_PATH))
        BANG_FONT = "Bangla"
    else:
        st.warning("⚠ Bangla font not found — using English font.")
except TTFError:
    st.warning("⚠ Error loading Bangla font — using English font.")

# ----------------------------
# Student Database
# ----------------------------
class StudentDatabase:
    def __init__(self, storage_path="students_storage.xlsx"):
        self.df = pd.DataFrame(columns=["Serial","ID","Name","Father","Mother","Class","Session","DOB"])
        self.storage_path = storage_path
        self.filepath = None
        if os.path.exists(self.storage_path):
            try:
                self.load_excel(self.storage_path, copy_to_storage=False)
            except:
                pass

    def load_excel(self, path, copy_to_storage=True):
        df = pd.read_excel(path, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        expected = ["Serial","ID","Name","Father","Mother","Class","Session","DOB"]
        for col in expected:
            if col not in df.columns:
                df[col] = ""
        df["ID"] = df["ID"].astype(str)
        try:
            df["Serial"] = pd.to_numeric(df["Serial"], errors="coerce").fillna(0).astype(int)
        except:
            pass
        self.df = df[expected].copy()
        if copy_to_storage:
            shutil.copy(path, self.storage_path)
        self.filepath = self.storage_path if copy_to_storage else path

    def save_excel(self, path=None):
        if not path:
            path = self.filepath if self.filepath else self.storage_path
        self.df.to_excel(path, index=False, engine="openpyxl")
        self.filepath = path

    def get_next_serial(self):
        if self.df.empty:
            return 1
        try:
            ser = pd.to_numeric(self.df["Serial"], errors="coerce").astype(int)
            return int(ser.max())+1 if not ser.empty else 1
        except:
            return 1

    def get_student_by_id(self, student_id):
        if not student_id:
            return None
        matches = self.df[self.df["ID"].astype(str)==str(student_id)]
        if not matches.empty:
            row = matches.iloc[0]
            return {col: row[col] for col in self.df.columns}
        return None

    def upsert_student(self, data: dict):
        sid = str(data.get("ID","")).strip()
        if sid == "":
            raise ValueError("ID required")
        idx = self.df[self.df["ID"].astype(str)==sid].index
        if len(idx) > 0:
            i = idx[0]
            for k,v in data.items():
                if k in self.df.columns:
                    self.df.at[i,k] = v
        else:
            self.df = pd.concat([self.df, pd.DataFrame([data])], ignore_index=True)
        try:
            self.df["Serial"] = pd.to_numeric(self.df["Serial"], errors="coerce").fillna(0).astype(int)
        except:
            pass

# ----------------------------
# PDF Generator
# ----------------------------
def generate_testimonial_pdf(entry, gender, pdf_path):
    sn = entry["Serial"]
    date = entry["Date"]
    student_id = entry["ID"]
    student_class = entry["Class"]
    session = entry["Session"]
    name = entry["Name"]
    father = entry["Father"]
    mother = entry["Mother"]
    dob = entry["DOB"]

    if gender.lower()=="male":
        he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
    else:
        he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

    c = canvas.Canvas(pdf_path, pagesize=A4)
    W,H = A4
    left,right = 25*mm,25*mm

    # Heading
    heading_w, heading_h = 120*mm, 18*mm
    heading_x = (W - heading_w)/2
    heading_y = H - 60*mm
    c.roundRect(heading_x, heading_y, heading_w, heading_h, 6)
    c.setFont(BANG_FONT, 17)
    c.drawCentredString(W/2, heading_y + heading_h/2 - 6, "Testimonial Certificate")

    # Table
    table_x = left
    table_y_top = heading_y - 20*mm
    cell_w1, cell_w2, cell_h = 30*mm, 55*mm, 9*mm
    c.setFont(BANG_FONT, 11)
    keys = ["S/N","Date","ID No","Class","Session"]
    vals = [str(sn), date, student_id, student_class, session]
    for i,key in enumerate(keys):
        y = table_y_top - i*cell_h
        c.rect(table_x, y-cell_h, cell_w1, cell_h)
        c.rect(table_x+cell_w1, y-cell_h, cell_w2, cell_h)
        c.drawString(table_x+3, y-cell_h/2+2, key)
        c.drawString(table_x+cell_w1+4, y-cell_h/2+2, str(vals[i]))

    # Paragraph
    intro_y = table_y_top - len(keys)*cell_h - 10*mm
    c.setFont(BANG_FONT, 17)
    c.drawCentredString(W/2, intro_y, "This is to certify that")

    # Extra space before paragraph
    paragraph_y = intro_y - 15*mm

    paragraph = (
        f"{name} {son_daughter} of {father} and {mother} is a student of Class {student_class}. "
        f"Bearing ID/Roll: {student_id} in Daffodil University School & College. "
        f"As per our admission record {his_her} date of birth is {dob}. "
        f"To the best of my knowledge {he_she} was well mannered and possessed a good moral character. "
        f"{He_She} did not indulge {Him_Her}self in any activity subversive to the discipline. "
        f"I wish {Him_Her} every success in life."
    )

    style = ParagraphStyle(
        name="Justify",
        fontName=BANG_FONT,
        fontSize=11,
        leading=14,
        alignment=TA_JUSTIFY,
    )

    p = Paragraph(paragraph, style)
    frame = Frame(left, sig_y + 15*mm, W-left-right, paragraph_start_y - (sig_y + 15*mm), showBoundary=0)
    frame.addFromList([p], c)

    # Signature
    line_width = 60*mm
    c.line(left, sig_y, left+line_width, sig_y)
    c.setFont(BANG_FONT, 11)
    for i, line in enumerate(["SK Mahmudun Nabi","Principal (Acting)","Daffodil University School & College"]):
        c.drawString(left, sig_y-12-i*12, line)

    c.save()

# ----------------------------
# Transfer Certificate PDF
# ----------------------------
def generate_tc_pdf(entry, gender, pdf_path):
    sn = entry["Serial"]
    date = entry["Date"]
    student_id = entry["ID"]
    student_class = entry["Class"]
    session = entry["Session"]
    name = entry["Name"]
    father = entry["Father"]
    mother = entry["Mother"]
    dob = entry["DOB"]

    if gender.lower()=="male":
        he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
    else:
        he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

    c = canvas.Canvas(pdf_path, pagesize=A4)
    W,H = A4
    left,right = 25*mm,25*mm

    # Heading
    heading_w, heading_h = 120*mm, 18*mm
    heading_x = (W - heading_w)/2
    heading_y = H - 60*mm
    c.roundRect(heading_x, heading_y, heading_w, heading_h, 6)
    c.setFont(BANG_FONT, 17)
    c.drawCentredString(W/2, heading_y+heading_h/2 - 6, "Transfer Certificate")

    # Table
    table_x = left
    table_y_top = heading_y - 20*mm
    cell_w1, cell_w2, cell_h = 30*mm, 55*mm, 9*mm
    c.setFont(BANG_FONT, 11)
    keys = ["S/N","Date","ID No","Class","Session"]
    vals = [str(sn), date, student_id, student_class, session]
    for i,key in enumerate(keys):
        y = table_y_top - i*cell_h
        c.rect(table_x, y-cell_h, cell_w1, cell_h)
        c.rect(table_x+cell_w1, y-cell_h, cell_w2, cell_h)
        c.drawString(table_x+3, y-cell_h/2+2, key)
        c.drawString(table_x+cell_w1+4, y-cell_h/2+2, str(vals[i]))

    # Paragraph
    intro_y = table_y_top - len(keys)*cell_h - 10*mm
    c.setFont(BANG_FONT, 17)
    c.drawCentredString(W/2, intro_y, "This is to certify that")

    # Extra space
    paragraph_y = intro_y - 15*mm

    paragraph = (
        f"{name}, {son_daughter} of {father} and {mother}, was a student of Class {student_class} "
        f"(Bearing ID/Roll: {student_id}) at Daffodil University School & College. As per our "
        f"record, {his_her} date of birth is {dob}. During {his_her} stay, {he_she} maintained "
        f"good conduct and discipline. We wish {Him_Her} success in future life."
    )

    style = ParagraphStyle(
        name="JustifyTC",
        fontName=BANG_FONT,
        fontSize=11,
        leading=14,
        alignment=TA_JUSTIFY,
    )

    p = Paragraph(paragraph, style)
    sig_y = 110*mm
    frame_bottom = sig_y + 15*mm
    frame_top = paragraph_y - 5
    frame_height = max(40*mm, frame_top - frame_bottom)

    frame = Frame(left, frame_bottom, W-left-right, frame_height, showBoundary=0)
    frame.addFromList([p], c)

    # Signature
    c.line(left, sig_y, left+60*mm, sig_y)
    c.setFont(BANG_FONT, 11)
    for i, line in enumerate(["SK Mahmudun Nabi","Principal (Acting)","Daffodil University School & College"]):
        c.drawString(left, sig_y-12-i*12, line)

    c.save()

# ----------------------------
# Streamlit App
# ----------------------------
st.set_page_config(page_title="Testimonial & TC Generator", layout="wide")
st.title("Testimonial & Transfer Certificate Generator (Excel-based)")

# Session state
for key in ["form_serial","form_date","form_id","form_class","form_session","form_name","form_father","form_mother","form_dob","form_gender"]:
    if key not in st.session_state:
        st.session_state[key] = "" if "gender" not in key else "Male"

db = StudentDatabase()

# Load Excel
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx","xls"])
if uploaded_file:
    db.load_excel(uploaded_file)
    st.success(f"Excel Loaded. {len(db.df)} students in database.")
    st.session_state.form_serial = db.get_next_serial()

# Form Inputs
col1,col2 = st.columns(2)
with col1:
    st.session_state.form_serial = st.number_input("S/N", min_value=1, value=int(st.session_state.form_serial or 1))
    st.session_state.form_date = st.date_input("Date", value=datetime.today())
    st.session_state.form_id = st.text_input("Student ID", value=st.session_state.form_id)
    st.session_state.form_class = st.text_input("Class", value=st.session_state.form_class)
    st.session_state.form_session = st.text_input("Session", value=st.session_state.form_session)
with col2:
    st.session_state.form_name = st.text_input("Student Name", value=st.session_state.form_name)
    st.session_state.form_father = st.text_input("Father's Name", value=st.session_state.form_father)
    st.session_state.form_mother = st.text_input("Mother's Name", value=st.session_state.form_mother)
    st.session_state.form_dob = st.text_input("Date of Birth (DD/MM/YYYY)", value=st.session_state.form_dob)
    st.session_state.form_gender = st.selectbox("Gender", ["Male","Female"], index=0 if st.session_state.form_gender=="Male" else 1)

# Auto-fill by ID
if st.session_state.form_id:
    rec = db.get_student_by_id(st.session_state.form_id)
    if rec:
        for field in ["Serial","Name","Father","Mother","Class","Session","DOB"]:
            st.session_state["form_"+field.lower()] = rec.get(field, st.session_state.get("form_"+field.lower()))

# Generate Buttons
col_gen, col_tc = st.columns(2)
with col_gen:
    if st.button("Generate Testimonial PDF"):
        entry = {
            "Serial": int(st.session_state.form_serial),
            "ID": st.session_state.form_id,
            "Name": st.session_state.form_name,
            "Father": st.session_state.form_father,
            "Mother": st.session_state.form_mother,
            "Class": st.session_state.form_class,
            "Session": st.session_state.form_session,
            "DOB": st.session_state.form_dob,
            "Date": st.session_state.form_date.strftime("%d/%m/%Y")
        }
        db.upsert_student(entry)
        db.save_excel()
        pdf_path = f"testimonial_{entry['ID']}_{datetime.now().strftime('%H%M%S')}.pdf"
        generate_testimonial_pdf(entry, st.session_state.form_gender, pdf_path)
        st.success("Testimonial PDF Generated Successfully!")
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF", f, file_name=os.path.basename(pdf_path))

with col_tc:
    if st.button("Generate Transfer Certificate PDF"):
        entry = {
            "Serial": int(st.session_state.form_serial),
            "ID": st.session_state.form_id,
            "Name": st.session_state.form_name,
            "Father": st.session_state.form_father,
            "Mother": st.session_state.form_mother,
            "Class": st.session_state.form_class,
            "Session": st.session_state.form_session,
            "DOB": st.session_state.form_dob,
            "Date": st.session_state.form_date.strftime("%d/%m/%Y")
        }
        db.upsert_student(entry)
        db.save_excel()
        pdf_path = f"tc_{entry['ID']}_{datetime.now().strftime('%H%M%S')}.pdf"
        generate_tc_pdf(entry, st.session_state.form_gender, pdf_path)
        st.success("Transfer Certificate PDF Generated Successfully!")
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF", f, file_name=os.path.basename(pdf_path))

# Show Database
st.subheader("Student Database")
if not db.df.empty:
    edited_df = st.data_editor(db.df, num_rows="dynamic")
    if st.button("Save Edited Excel"):
        db.df = edited_df
        db.save_excel()
        st.success("Excel Saved Successfully!")

