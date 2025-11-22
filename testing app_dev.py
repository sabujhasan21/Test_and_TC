# testimonial_app_streamlit.py
import streamlit as st
import pandas as pd
import os
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# ----------------------------
# Student Database Class
# ----------------------------
class StudentDatabase:
    def __init__(self, storage_path="students_storage.xlsx"):
        self.storage_path = storage_path
        self.df = pd.DataFrame(columns=["Serial", "ID", "Name", "Father", "Mother", "Class", "Session", "DOB"])
        if os.path.exists(storage_path):
            self.load_excel(storage_path, copy_to_storage=False)

    def load_excel(self, path, copy_to_storage=True):
        df = pd.read_excel(path, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        expected = ["Serial", "ID", "Name", "Father", "Mother", "Class", "Session", "DOB"]
        for col in expected:
            if col not in df.columns:
                df[col] = ""
        df = df[expected]
        df["ID"] = df["ID"].apply(lambda x: str(int(float(x))) if str(x).replace('.', '',1).isdigit() else str(x))
        try:
            df["Serial"] = pd.to_numeric(df["Serial"], errors="coerce").fillna(0).astype(int)
        except:
            pass
        self.df = df
        if copy_to_storage:
            df.to_excel(self.storage_path, index=False, engine="openpyxl")

    def save_excel(self, path=None):
        if path is None:
            path = self.storage_path
        self.df.to_excel(path, index=False, engine="openpyxl")

    def get_next_serial(self):
        if self.df.empty:
            return 1
        try:
            ser = pd.to_numeric(self.df["Serial"].dropna(), errors="coerce").astype(int)
            return int(ser.max()) + 1 if not ser.empty else 1
        except:
            return 1

    def get_student_by_id(self, student_id):
        if not student_id:
            return None
        matches = self.df[self.df["ID"].astype(str) == str(student_id)]
        if not matches.empty:
            row = matches.iloc[0]
            return dict(row)
        return None

    def upsert_student(self, data: dict):
        sid = str(data.get("ID", "")).strip()
        if sid == "":
            raise ValueError("ID required")
        idx = self.df[self.df["ID"].astype(str) == sid].index
        if len(idx) > 0:
            i = idx[0]
            for k, v in data.items():
                if k in self.df.columns:
                    self.df.at[i, k] = v
        else:
            self.df = pd.concat([self.df, pd.DataFrame([data])], ignore_index=True)
        try:
            self.df["Serial"] = pd.to_numeric(self.df["Serial"], errors="coerce").fillna(0).astype(int)
        except:
            pass

# ----------------------------
# PDF Generation
# ----------------------------
def create_testimonial_pdf(entry, gender, pdf_path):
    sn = entry["Serial"]
    date = entry["Date"]
    student_id = entry["ID"]
    student_class = entry["Class"]
    session = entry["Session"]
    name = entry["Name"]
    father = entry["Father"]
    mother = entry["Mother"]
    dob = entry["DOB"]

    if gender.lower() == "male":
        he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
    else:
        he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

    c = canvas.Canvas(pdf_path, pagesize=A4)
    W,H = A4
    left = 25*mm
    right = 25*mm

    # Heading
    heading_w = 120*mm
    heading_h = 18*mm
    heading_x = (W-heading_w)/2
    heading_y = H-60*mm
    c.setLineWidth(1)
    c.roundRect(heading_x, heading_y, heading_w, heading_h, 6, stroke=1, fill=0)
    c.setFont("Times-Bold", 17)
    c.drawCentredString(W/2, heading_y+heading_h/2-6, "Testimonial Certificate")

    # Left table
    table_x = left
    table_y_top = heading_y-20*mm
    cell_w1 = 30*mm
    cell_w2 = 55*mm
    cell_h = 9*mm
    keys = ["S/N", "Date", "ID No", "Class", "Session"]
    vals = [str(sn), date, student_id, student_class, session]
    c.setFont("Times-Roman", 11)
    for i,key in enumerate(keys):
        y = table_y_top-i*cell_h
        c.rect(table_x, y-cell_h, cell_w1, cell_h)
        c.rect(table_x+cell_w1, y-cell_h, cell_w2, cell_h)
        c.drawString(table_x+3, y-cell_h/2+2, key)
        c.drawString(table_x+cell_w1+4, y-cell_h/2+2, vals[i])

    # Intro text
    intro_y = table_y_top-len(keys)*cell_h-10*mm
    c.setFont("Times-Bold", 17)
    c.drawCentredString(W/2, intro_y, "This is to certify that")

    # Paragraph
    paragraph = (
        f"{name} {son_daughter} of {father} and {mother} is a student of Class {student_class}. "
        f"Bearing ID/Roll: {student_id}. Date of birth: {dob}. "
        f"To the best of our knowledge, {he_she} was well mannered and possessed good moral character. "
        f"{He_She} did not indulge in any activity against discipline. We wish {Him_Her} every success!"
    )
    style = ParagraphStyle(name="Justify", fontName="Times-Roman", fontSize=11, leading=14, alignment=TA_JUSTIFY)
    p = Paragraph(paragraph, style)
    sig_y = 110*mm
    frame_bottom = sig_y+15*mm
    frame_top = intro_y-10
    frame_height = max(30*mm, frame_top-frame_bottom)
    frame = Frame(left, frame_bottom, W-left-right, frame_height, showBoundary=0)
    frame.addFromList([p], c)

    # Signature
    line_width = 60*mm
    c.line(left, sig_y, left+line_width, sig_y)
    text_lines = ["SK Mahmudun Nabi","Principal (Acting)","Daffodil University School & College"]
    for i,line in enumerate(text_lines):
        c.drawString(left, sig_y-12-i*12, line)
    c.save()

# ----------------------------
# Streamlit App
# ----------------------------
st.set_page_config(page_title="Testimonial & TC Generator", layout="wide")
st.title("Testimonial & Transfer Certificate Generator (Excel-based)")

# Database
db = StudentDatabase()

# ----------------------------
# Upload Excel
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx","xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    storage_path = "students_storage.xlsx"
    df.to_excel(storage_path, index=False, engine="openpyxl")
    db.load_excel(storage_path, copy_to_storage=False)
    st.success(f"Excel Loaded. {len(db.df)} students in database.")
    if 'form_serial' not in st.session_state:
        st.session_state.form_serial = db.get_next_serial()

# ----------------------------
# Form Inputs
cols = st.columns(3)
with cols[0]:
    sn = st.number_input("S/N", min_value=1, value=st.session_state.get('form_serial',1))
with cols[1]:
    date = st.date_input("Date").strftime("%d/%m/%Y")
with cols[2]:
    student_id = st.text_input("ID No")

student_rec = db.get_student_by_id(student_id)
if student_rec:
    st.session_state.form_serial = int(student_rec.get("Serial", sn))
    name = st.text_input("Student Name", student_rec.get("Name",""))
    father = st.text_input("Father's Name", student_rec.get("Father",""))
    mother = st.text_input("Mother's Name", student_rec.get("Mother",""))
    student_class = st.text_input("Class", student_rec.get("Class",""))
    session_val = st.text_input("Session", student_rec.get("Session",""))
    dob = st.text_input("DOB (DD/MM/YYYY)", student_rec.get("DOB",""))
else:
    name = st.text_input("Student Name")
    father = st.text_input("Father's Name")
    mother = st.text_input("Mother's Name")
    student_class = st.text_input("Class")
    session_val = st.text_input("Session")
    dob = st.text_input("DOB (DD/MM/YYYY)")

gender = st.selectbox("Select Gender", ["Male","Female"])

# ----------------------------
# Generate PDF
if st.button("Generate Testimonial Certificate"):
    if not all([sn, student_id, name]):
        st.warning("S/N, ID, and Student Name required!")
    else:
        entry = {"Serial": sn,"ID":student_id,"Name":name,"Father":father,"Mother":mother,
                 "Class":student_class,"Session":session_val,"DOB":dob,"Date":date}
        db.upsert_student(entry)
        db.save_excel()
        pdf_path = f"testimonial_{student_id}.pdf"
        create_testimonial_pdf(entry, gender, pdf_path)
        st.success(f"PDF generated: {pdf_path}")
        st.download_button("Download PDF", data=open(pdf_path,"rb"), file_name=pdf_path, mime="application/pdf")
