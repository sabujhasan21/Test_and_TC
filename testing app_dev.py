# streamlit_testimonial_app.py
import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
import fitz  # PyMuPDF for PDF -> image preview

# ----------------- CONFIG -----------------
STORAGE_EXCEL = "students_storage.xlsx"
PDF_FOLDER = "generated_pdfs"
os.makedirs(PDF_FOLDER, exist_ok=True)

REQUIRED_COLUMNS = ["Serial", "ID", "Name", "Father", "Mother", "Class", "Session", "DOB"]

# ----------------- HELPERS -----------------
def ensure_df(df):
    # Ensure expected columns exist and proper formatting
    for c in REQUIRED_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    df = df[REQUIRED_COLUMNS].copy()
    # Serial numeric
    try:
        df["Serial"] = pd.to_numeric(df["Serial"].fillna(0), errors="coerce").fillna(0).astype(int)
    except Exception:
        pass
    # ID as integer-like string when possible
    def norm_id(x):
        sx = str(x)
        if sx.replace('.', '', 1).isdigit():
            try:
                return str(int(float(sx)))
            except Exception:
                return sx
        return sx
    df["ID"] = df["ID"].apply(norm_id)
    return df

def load_storage():
    if "df" in st.session_state:
        return st.session_state.df
    if os.path.exists(STORAGE_EXCEL):
        try:
            df = pd.read_excel(STORAGE_EXCEL, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            df = ensure_df(df)
            st.session_state.df = df
            return df
        except Exception:
            pass
    df = pd.DataFrame(columns=REQUIRED_COLUMNS)
    st.session_state.df = ensure_df(df)
    return st.session_state.df

def save_storage(path=None):
    if path is None:
        path = STORAGE_EXCEL
    df = st.session_state.df.copy()
    df.to_excel(path, index=False, engine="openpyxl")
    st.session_state.storage_path = path
    st.success(f"Saved Excel to: {os.path.basename(path)}")

def get_next_serial():
    df = load_storage()
    if df.empty:
        return 1
    try:
        ser = pd.to_numeric(df["Serial"].dropna(), errors="coerce").astype(int)
        return int(ser.max()) + 1 if not ser.empty else 1
    except Exception:
        return 1

def get_student_by_id(student_id):
    df = load_storage()
    if student_id is None or str(student_id).strip()=="":
        return None
    matches = df[df["ID"].astype(str) == str(student_id)]
    if not matches.empty:
        row = matches.iloc[0]
        return row.to_dict()
    return None

def upsert_student(entry):
    df = load_storage()
    sid = str(entry.get("ID", "")).strip()
    if sid == "":
        raise ValueError("ID required to upsert.")
    idx = df[df["ID"].astype(str) == sid].index
    if len(idx) > 0:
        i = idx[0]
        for k, v in entry.items():
            if k in df.columns:
                df.at[i, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    try:
        df["Serial"] = pd.to_numeric(df["Serial"], errors="coerce").fillna(0).astype(int)
    except Exception:
        pass
    st.session_state.df = ensure_df(df)

# PDF create (Testimonial / TC) using reportlab (same layout logic)
def _create_pdf_bytes(kind, entry, gender, date_str):
    """
    kind: 'testimonial' or 'tc'
    entry: dict with fields
    returns: bytes of PDF
    """
    sn = entry.get("Serial", "")
    student_id = entry.get("ID", "")
    student_class = entry.get("Class", "")
    session = entry.get("Session", "")
    name = entry.get("Name", "")
    father = entry.get("Father", "")
    mother = entry.get("Mother", "")
    dob = entry.get("DOB", "")

    if gender == "male":
        he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
    else:
        he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4
    left = 25 * mm
    right = 25 * mm

    # Heading
    heading_w = 120 * mm
    heading_h = 18 * mm
    heading_x = (W - heading_w)/2
    heading_y = H - 60*mm
    c.setLineWidth(1)
    c.roundRect(heading_x, heading_y, heading_w, heading_h, 6, stroke=1, fill=0)
    c.setFont("Times-Bold", 17)
    title = "Testimonial Certificate" if kind=="testimonial" else "Transfer Certificate"
    c.drawCentredString(W/2, heading_y + heading_h/2 - 6, title)

    # Left table
    table_x = left
    table_y_top = heading_y - 20*mm
    cell_w1 = 30*mm
    cell_w2 = 55*mm
    cell_h = 9*mm

    c.setFont("Times-Roman", 12)
    keys = ["S/N", "Date", "ID No", "Class", "Session"]
    vals = [str(sn), date_str, str(student_id), str(student_class), str(session)]
    for i, key in enumerate(keys):
        y = table_y_top - i*cell_h
        c.rect(table_x, y-cell_h, cell_w1, cell_h)
        c.rect(table_x+cell_w1, y-cell_h, cell_w2, cell_h)
        c.drawString(table_x+3, y-cell_h/2+2, key)
        c.drawString(table_x+cell_w1+4, y-cell_h/2+2, vals[i])

    # Intro
    intro_y = table_y_top - len(keys)*cell_h - 10*mm
    c.setFont("Times-Bold", 17)
    c.drawCentredString(W/2, intro_y, "This is to certify that")

    # Paragraph
    if kind == "testimonial":
        paragraph = (
            f"{name} {son_daughter} of {father} and {mother} is a student of Class: {student_class}. "
            f"Bearing ID/Roll: {student_id} in Daffodil University School & College. "
            f"As per our admission record {his_her} date of birth is {dob}. "
            f"To the best of my knowledge {he_she} was well mannered and possessed a good moral character. "
            f"{He_She} did not indulge {Him_Her}self in any activity subversive to the state and discipline during study. "
            f"I wish {Him_Her} every success in life!"
        )
    else:  # tc
        paragraph = (
            f"{name}, {son_daughter} of {father} and {mother}, "
            f"was a student of Class {student_class} (Bearing ID/Roll: {student_id}) at "
            f"Daffodil University School & College. As per our record, {his_her} date of birth "
            f"is {dob}. During {his_her} stay, {he_she} maintained good conduct and discipline. "
            f"We wish {Him_Her} every success in future life."
        )

    # signature baseline (same as before)
    sig_y = 110*mm

    # Paragraph style & frame calculation
    style = ParagraphStyle(
        name="Justify",
        fontName="Times-Roman",
        fontSize=12,
        leading=14,
        alignment=TA_JUSTIFY,
    )

    p = Paragraph(paragraph, style)

    frame_bottom = sig_y + 15 * mm
    frame_top = intro_y - 10
    frame_height = max(30 * mm, frame_top - frame_bottom)
    frame_y = frame_bottom

    frame = Frame(
        left,
        frame_y,
        W - left - right,
        frame_height,
        showBoundary=0
    )

    frame.addFromList([p], c)

    # Signature
    line_width = 60*mm
    c.line(left, sig_y, left+line_width, sig_y)
    c.setFont("Times-Roman", 12)
    text_lines = ["SK Mahmudun Nabi", "Principal (Acting)", "Daffodil University School & College"]
    for i, line in enumerate(text_lines):
        c.drawString(left, sig_y-12-i*12, line)

    c.save()
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

def save_pdf_file(filename, pdf_bytes):
    path = os.path.join(PDF_FOLDER, filename)
    with open(path, "wb") as f:
        f.write(pdf_bytes)
    return path

def preview_pdf_first_page(pdf_path):
    """Return PNG bytes of first page for display in Streamlit."""
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    except Exception:
        return None

# ----------------- STREAMLIT UI -----------------
st.set_page_config(page_title="Testimonial & TC Generator", layout="wide")
st.title("Testimonial & Transfer Certificate Generator — Streamlit (Full feature)")

# Left: Controls, Right: Table & Preview
col_left, col_right = st.columns([1, 1.2])

with col_left:
    st.header("Controls")

    # Load Excel
    uploaded = st.file_uploader("Load Excel (.xlsx/.xls)", type=["xlsx", "xls"], accept_multiple_files=False)
    if uploaded:
        tmp_path = os.path.join(".", "tmp_uploaded.xlsx")
        with open(tmp_path, "wb") as f:
            f.write(uploaded.read())
        # copy to storage
        try:
            shutil.copy(tmp_path, STORAGE_EXCEL)
            st.success("Excel uploaded and stored.")
            if "df" in st.session_state:
                del st.session_state["df"]
            load_storage()
            os.remove(tmp_path)
        except Exception as e:
            st.error(f"Could not store uploaded excel: {e}")

    if st.button("Save Excel"):
        # sync happens below from editor, but save current session df
        save_storage()

    if st.button("Save Excel As..."):
        # produce a downloadable excel
        df = load_storage()
        out = BytesIO()
        df.to_excel(out, index=False, engine="openpyxl")
        out.seek(0)
        st.download_button("Download Excel file", data=out, file_name=f"students_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    if st.button("New Serial"):
        st.session_state.form_serial = get_next_serial()
        st.info(f"Next serial set to {st.session_state.form_serial}")

    st.markdown("---")
    st.subheader("Form (fill and generate)")

    # Prepare session values
    df = load_storage()
    # form fields (use session_state to persist while editing)
    if "form_serial" not in st.session_state:
        st.session_state.form_serial = get_next_serial()
    if "form_date" not in st.session_state:
        st.session_state.form_date = datetime.today().strftime("%d/%m/%Y")
    if "form_id" not in st.session_state:
        st.session_state.form_id = ""
    if "form_class" not in st.session_state:
        st.session_state.form_class = ""
    if "form_session" not in st.session_state:
        st.session_state.form_session = ""
    if "form_name" not in st.session_state:
        st.session_state.form_name = ""
    if "form_father" not in st.session_state:
        st.session_state.form_father = ""
    if "form_mother" not in st.session_state:
        st.session_state.form_mother = ""
    if "form_dob" not in st.session_state:
        st.session_state.form_dob = ""
    if "form_gender" not in st.session_state:
        st.session_state.form_gender = "Male"

    # Inputs
    st.number_input("S/N", value=int(st.session_state.form_serial), key="form_serial", step=1)
    st.text_input("Date (DD/MM/YYYY)", value=st.session_state.form_date, key="form_date")
    st.text_input("ID No", value=st.session_state.form_id, key="form_id", on_change=None)
    # ID autofill button
    if st.button("Auto-fill from ID"):
        rec = get_student_by_id(st.session_state.form_id)
        if rec:
            st.session_state.form_serial = int(rec.get("Serial", st.session_state.form_serial) or st.session_state.form_serial)
            st.session_state.form_name = rec.get("Name", "")
            st.session_state.form_father = rec.get("Father", "")
            st.session_state.form_mother = rec.get("Mother", "")
            st.session_state.form_class = rec.get("Class", "")
            st.session_state.form_session = rec.get("Session", "")
            st.session_state.form_dob = rec.get("DOB", "")
            st.experimental_rerun()
        else:
            st.warning("No student with that ID found. Serial set to next.")
            st.session_state.form_serial = get_next_serial()

    st.text_input("Class", value=st.session_state.form_class, key="form_class")
    st.text_input("Session", value=st.session_state.form_session, key="form_session")
    st.text_input("Student Name", value=st.session_state.form_name, key="form_name")
    st.text_input("Father's Name", value=st.session_state.form_father, key="form_father")
    st.text_input("Mother's Name", value=st.session_state.form_mother, key="form_mother")
    st.text_input("Date of Birth (DD/MM/YYYY)", value=st.session_state.form_dob, key="form_dob")
    st.selectbox("Select Gender:", ["Male", "Female"], key="form_gender")

    st.markdown("")  # spacing
    colg1, colg2 = st.columns(2)
    if colg1.button("Generate Testimonial PDF"):
        # validation
        if not all([st.session_state.form_serial, st.session_state.form_date, st.session_state.form_id, st.session_state.form_name]):
            st.error("Please ensure at least S/N, Date, ID and Student Name are filled.")
        else:
            entry = {
                "Serial": int(st.session_state.form_serial) if str(st.session_state.form_serial).isdigit() else st.session_state.form_serial,
                "ID": str(st.session_state.form_id),
                "Name": st.session_state.form_name,
                "Father": st.session_state.form_father,
                "Mother": st.session_state.form_mother,
                "Class": st.session_state.form_class,
                "Session": st.session_state.form_session,
                "DOB": st.session_state.form_dob
            }
            try:
                upsert_student(entry)
                save_storage()  # persist
            except Exception as e:
                st.error(f"Could not upsert student data: {e}")
            date_str = st.session_state.form_date
            pdf_bytes = _create_pdf_bytes("testimonial", entry, st.session_state.form_gender.lower(), date_str)
            filename = f"testimonial_{entry['ID']}.pdf"
            path = save_pdf_file(filename, pdf_bytes)
            st.success(f"PDF saved: {filename}")
            st.session_state.last_pdf = path
            st.experimental_rerun()

    if colg2.button("Generate Transfer Certificate PDF"):
        if not all([st.session_state.form_serial, st.session_state.form_date, st.session_state.form_id, st.session_state.form_name]):
            st.error("Please ensure at least S/N, Date, ID and Student Name are filled.")
        else:
            entry = {
                "Serial": int(st.session_state.form_serial) if str(st.session_state.form_serial).isdigit() else st.session_state.form_serial,
                "ID": str(st.session_state.form_id),
                "Name": st.session_state.form_name,
                "Father": st.session_state.form_father,
                "Mother": st.session_state.form_mother,
                "Class": st.session_state.form_class,
                "Session": st.session_state.form_session,
                "DOB": st.session_state.form_dob
            }
            try:
                upsert_student(entry)
                save_storage()
            except Exception as e:
                st.error(f"Could not upsert data: {e}")
            date_str = st.session_state.form_date
            pdf_bytes = _create_pdf_bytes("tc", entry, st.session_state.form_gender.lower(), date_str)
            filename = f"transfer_certificate_{entry['ID']}.pdf"
            path = save_pdf_file(filename, pdf_bytes)
            st.success(f"TC PDF saved: {filename}")
            st.session_state.last_pdf = path
            st.experimental_rerun()

with col_right:
    st.header("Excel Data (editable)")
    df = load_storage()

    # Show editable table using st.data_editor (Streamlit >=1.18). Falls back to st.experimental_data_editor if older.
    try:
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    except Exception:
        edited = st.experimental_data_editor(df, num_rows="dynamic", use_container_width=True)

    # When user edits table, sync back
    if st.button("Apply Table Edits"):
        try:
            edited_df = pd.DataFrame(edited)
            edited_df = ensure_df(edited_df)
            st.session_state.df = edited_df
            st.success("Table changes applied to session.")
        except Exception as e:
            st.error(f"Could not apply edits: {e}")

    # Download last generated PDF & preview
    st.markdown("---")
    st.subheader("Last PDF Preview & Download")
    if "last_pdf" in st.session_state and st.session_state.last_pdf and os.path.exists(st.session_state.last_pdf):
        pdf_path = st.session_state.last_pdf
        st.write(f"**File:** {os.path.basename(pdf_path)}")
        img = preview_pdf_first_page(pdf_path)
        if img:
            st.image(img, caption="Preview (first page)", use_column_width=True)
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF", data=f, file_name=os.path.basename(pdf_path))
    else:
        st.info("No PDF generated yet. Generate one from left panel.")

    st.markdown("---")
    st.subheader("Other Utilities")
    # Clear stored PDF files
    if st.button("Clear generated PDFs"):
        try:
            for fn in os.listdir(PDF_FOLDER):
                os.remove(os.path.join(PDF_FOLDER, fn))
            st.success("Cleared generated PDFs.")
        except Exception as e:
            st.error(f"Could not clear PDFs: {e}")

    # Export current dataframe to Excel
    buf = BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    st.download_button("Download current Excel", data=buf, file_name="students_current.xlsx")

st.markdown("---")
st.write("Developed to mirror PyQt app features — If you want adjustments (logo, signature image, fonts, layout tweaks), tell me and I’ll add them.")
