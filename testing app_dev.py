# testimonial_tc_streamlit.py
import os
import shutil
import pandas as pd
from datetime import datetime
import streamlit as st
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ----------------------------
# Register Bangla Font
# ----------------------------
FONT_PATH = "fonts/SolaimanLipi.ttf"  # path to TTF font
pdfmetrics.registerFont(TTFont("Bangla", FONT_PATH))

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
        df["ID"] = df["ID"].apply(lambda x: str(int(float(x))) if str(x).replace(".","",1).isdigit() else str(x))
        self.df = df[expected].copy()
        try:
            self.df["Serial"] = pd.to_numeric(self.df["Serial"], errors="coerce").fillna(0).astype(int)
        except:
            pass
        self.filepath = self.storage_path if copy_to_storage else path
        if copy_to_storage:
            shutil.copy(path, self.storage_path)

    def save_excel(self, path=None):
        if not path:
            path = self.filepath if self.filepath else self.storage_path
        self.df.to_excel(path, index=False, engine="openpyxl")
        self.filepath = path

    def get_next_serial(self):
        if self.df.empty: return 1
        try:
            ser = pd.to_numeric(self.df["Serial"].dropna(), errors="coerce").astype(int)
            return int(ser.max())+1 if not ser.empty else 1
        except:
            return 1

    def get_student_by_id(self, student_id):
        if not student_id: return None
        matches = self.df[self.df["ID"].astype(str)==str(student_id)]
        if not matches.empty:
            row = matches.iloc[0]
            return {col: row[col] for col in self.df.columns}
        return None

    def upsert_student(self, data: dict):
        sid = str(data.get("ID","")).strip()
        if sid == "": raise ValueError("ID required")
        idx = self.df[self.df["ID"].astype(str)==sid].index
        if len(idx)>0:
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
# PDF Generation Helpers
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
    heading_w,heading_h = 120*mm,18*mm
    heading_x = (W-heading_w)/2
    heading_y = H-60*mm
    c.setLineWidth(1)
    c.roundRect(heading_x,heading_y,heading_w,heading_h,6,stroke=1,fill=0)
    c.setFont("Bangla",17)
    c.drawCentredString(W/2,heading_y+heading_h/2-6,"টেস্টিমোনিয়াল সার্টিফিকেট")

    # Left table
    table_x = left
    table_y_top = heading_y-20*mm
    cell_w1,cell_w2,cell_h = 30*mm,55*mm,9*mm
    c.setFont("Bangla",11)
    keys = ["S/N","তারিখ","আইডি নং","শ্রেণি","সেশন"]
    vals = [str(sn),date,student_id,student_class,session]
    for i,key in enumerate(keys):
        y = table_y_top-i*cell_h
        c.rect(table_x,y-cell_h,cell_w1,cell_h)
        c.rect(table_x+cell_w1,y-cell_h,cell_w2,cell_h)
        c.drawString(table_x+3,y-cell_h/2+2,key)
        c.drawString(table_x+cell_w1+4,y-cell_h/2+2,vals[i])

    # Intro paragraph
    intro_y = table_y_top-len(keys)*cell_h-10*mm
    c.setFont("Bangla",17)
    c.drawCentredString(W/2,intro_y,"This is to certify that")

    paragraph = (
        f"{name} {son_daughter} of {father} and {mother} হলেন {student_class} শ্রেণির ছাত্র। "
        f"আইডি নং: {student_id}, সেশন: {session}। "
        f"জন্ম তারিখ: {dob}। শিক্ষাজীবনের সময় আচরণ ও শৃঙ্খলা ভালো ছিল। "
        f"{He_She} এর ভবিষ্যতের জন্য শুভকামনা।"
    )

    sig_y = 110*mm
    style = ParagraphStyle(name="Justify", fontName="Bangla", fontSize=11, leading=14, alignment=TA_JUSTIFY)
    p = Paragraph(paragraph, style)
    frame_bottom = sig_y+15*mm
    frame_top = intro_y-10
    frame_height = max(30*mm, frame_top-frame_bottom)
    frame_y = frame_bottom
    frame = Frame(left,frame_y,W-left-right,frame_height,showBoundary=0)
    frame.addFromList([p],c)

    # Signature
    line_width = 60*mm
    c.line(left,sig_y,left+line_width,sig_y)
    c.setFont("Bangla",11)
    text_lines = ["SK Mahmudun Nabi","Principal (Acting)","Daffodil University School & College"]
    for i,line in enumerate(text_lines):
        c.drawString(left,sig_y-12-i*12,line)
    c.save()

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
    heading_w,heading_h = 120*mm,18*mm
    heading_x = (W-heading_w)/2
    heading_y = H-60*mm
    c.setLineWidth(1)
    c.roundRect(heading_x,heading_y,heading_w,heading_h,6,stroke=1,fill=0)
    c.setFont("Bangla",17)
    c.drawCentredString(W/2,heading_y+heading_h/2-6,"ট্রান্সফার সার্টিফিকেট")

    # Table
    table_x = left
    table_y_top = heading_y-20*mm
    cell_w1,cell_w2,cell_h = 30*mm,55*mm,9*mm
    c.setFont("Bangla",11)
    keys = ["S/N","তারিখ","আইডি নং","শ্রেণি","সেশন"]
    vals = [str(sn),date,student_id,student_class,session]
    for i,key in enumerate(keys):
        y = table_y_top-i*cell_h
        c.rect(table_x,y-cell_h,cell_w1,cell_h)
        c.rect(table_x+cell_w1,y-cell_h,cell_w2,cell_h)
        c.drawString(table_x+3,y-cell_h/2+2,key)
        c.drawString(table_x+cell_w1+4,y-cell_h/2+2,vals[i])

    intro_y = table_y_top-len(keys)*cell_h-10*mm
    c.setFont("Bangla",17)
    c.drawCentredString(W/2,intro_y,"This is to certify that")

    paragraph = (
        f"{name}, {son_daughter} of {father} and {mother}, "
        f"{student_class} শ্রেণির ছাত্র/ছাত্রী (ID: {student_id}) ছিলেন। "
        f"জন্ম তারিখ: {dob}। শিক্ষাজীবনের সময় আচরণ ও শৃঙ্খলা ভালো ছিল। "
        f"{He_She} এর ভবিষ্যতের জন্য শুভকামনা।"
    )

    sig_y = 110*mm
    style = ParagraphStyle(name="JustifyTC", fontName="Bangla", fontSize=11, leading=14, alignment=TA_JUSTIFY)
    p = Paragraph(paragraph, style)
    frame_bottom = sig_y+15*mm
    frame_top = intro_y-10
    frame_height = max(30*mm, frame_top-frame_bottom)
    frame_y = frame_bottom
    frame = Frame(left,frame_y,W-left-right,frame_height,showBoundary=0)
    frame.addFromList([p],c)

    line_width = 60*mm
    c.line(left,sig_y,left+line_width,sig_y)
    c.setFont("Bangla",11)
    text_lines = ["SK Mahmudun Nabi","Principal (Acting)","Daffodil University School & College"]
    for i,line in enumerate(text_lines):
        c.drawString(left,sig_y-12-i*12,line)
    c.save()

# ----------------------------
# Streamlit App
# ----------------------------
st.set_page_config(page_title="Testimonial & TC Generator", layout="wide")
st.title("Testimonial & Transfer Certificate Generator (Excel-based)")

# session_state initialization
for key in ["form_serial","form_date","form_id","form_class","form_session","form_name","form_father","form_mother","form_dob","form_gender"]:
    if key not in st.session_state:
        st.session_state[key] = "" if "gender" not in key else "Male"

db = StudentDatabase()

# ----------------------------
# Load Excel
# ----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx","xls"])
if uploaded_file:
    db.load_excel(uploaded_file)
    st.success(f"Excel Loaded. {len(db.df)} students in database.")
    st.session_state.form_serial = db.get_next_serial()

# ----------------------------
# Form Inputs
# ----------------------------
col1,col2 = st.columns([2,2])

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

# ----------------------------
# Auto-fill by ID
# ----------------------------
if st.session_state.form_id:
    rec = db.get_student_by_id(st.session_state.form_id)
    if rec:
        st.session_state.form_serial = rec.get("Serial", st.session_state.form_serial)
        st.session_state.form_name = rec.get("Name", st.session_state.form_name)
        st.session_state.form_father = rec.get("Father", st.session_state.form_father)
        st.session_state.form_mother = rec.get("Mother", st.session_state.form_mother)
        st.session_state.form_class = rec.get("Class", st.session_state.form_class)
        st.session_state.form_session = rec.get("Session", st.session_state.form_session)
        st.session_state.form_dob = rec.get("DOB", st.session_state.form_dob)

# ----------------------------
# Buttons
# ----------------------------
col_gen,col_preview = st.columns(2)

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
        pdf_path = f"testimonial_{entry['ID']}.pdf"
        generate_testimonial_pdf(entry, st.session_state.form_gender, pdf_path)
        st.success(f"Testimonial PDF Generated: {pdf_path}")
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF", f, file_name=os.path.basename(pdf_path))

with col_preview:
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
        pdf_path = f"transfer_certificate_{entry['ID']}.pdf"
        generate_tc_pdf(entry, st.session_state.form_gender, pdf_path)
        st.success(f"Transfer Certificate PDF Generated: {pdf_path}")
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF", f, file_name=os.path.basename(pdf_path))

# ----------------------------
# Show Excel Table
# ----------------------------
if not db.df.empty:
    st.subheader("Student Database")
    edited_df = st.data_editor(db.df, num_rows="dynamic")
    if st.button("Save Edited Excel"):
        db.df = edited_df
        db.save_excel()
        st.success("Excel Saved Successfully!")
