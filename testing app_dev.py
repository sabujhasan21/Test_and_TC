# testimonial_app_justified.py
import sys
import os
import shutil
import pandas as pd
from datetime import datetime
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY


from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QGridLayout, QMessageBox, QFileDialog,
    QHBoxLayout, QScrollArea, QFrame, QComboBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QSizePolicy
)
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt

import fitz  # PyMuPDF for preview
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# ----------------------------
# Student Excel Database Class
# ----------------------------
class StudentDatabase:
    def __init__(self, storage_path="students_storage.xlsx"):
        self.df = pd.DataFrame(columns=["Serial", "ID", "Name", "Father", "Mother", "Class", "Session", "DOB"])
        self.filepath = None
        self.storage_path = storage_path
        if os.path.exists(self.storage_path):
            try:
                self.load_excel(self.storage_path, copy_to_storage=False)
            except Exception:
                pass

    def load_excel(self, path, copy_to_storage=True):
        df = pd.read_excel(path, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]

        expected = ["Serial", "ID", "Name", "Father", "Mother", "Class", "Session", "DOB"]
        for col in expected:
            if col not in df.columns:
                df[col] = ""

        # Force ID to integer-like string
        df["ID"] = df["ID"].apply(lambda x: str(int(float(x))) if str(x).replace('.', '', 1).isdigit() else str(x))

        # Keep only expected columns
        self.df = df[expected].copy()

        # Ensure Serial numeric
        try:
            self.df["Serial"] = pd.to_numeric(self.df["Serial"], errors="coerce").fillna(0).astype(int)
        except Exception:
            pass

        if copy_to_storage:
            try:
                shutil.copy(path, self.storage_path)
                self.filepath = self.storage_path
            except Exception:
                self.filepath = path
        else:
            self.filepath = path

    def save_excel(self, path=None):
        if path is None:
            path = self.filepath if self.filepath is not None else self.storage_path
        self.df.to_excel(path, index=False, engine="openpyxl")
        self.filepath = path

    def get_next_serial(self):
        if self.df.empty:
            return 1
        try:
            ser = pd.to_numeric(self.df["Serial"].dropna(), errors="coerce").astype(int)
            return int(ser.max()) + 1 if not ser.empty else 1
        except Exception:
            return 1

    def get_student_by_id(self, student_id):
        if not student_id:
            return None
        matches = self.df[self.df["ID"].astype(str) == str(student_id)]
        if not matches.empty:
            row = matches.iloc[0]
            return {
                "Serial": row["Serial"],
                "ID": row["ID"],
                "Name": row["Name"],
                "Father": row["Father"],
                "Mother": row["Mother"],
                "Class": row["Class"],
                "Session": row["Session"],
                "DOB": row["DOB"]
            }
        return None

    def upsert_student(self, data: dict):
        sid = str(data.get("ID", "")).strip()
        if sid == "":
            raise ValueError("ID required to upsert.")
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
# PDF Preview Widget
# ----------------------------
class PDFPreview(QWidget):
    def __init__(self, pdf_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDF Preview")
        self.setFixedSize(700, 900)
        layout = QVBoxLayout()
        try:
            doc = fitz.open(pdf_path)
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=150)
            tmp_image = os.path.join(os.path.dirname(pdf_path), "tmp_preview.png")
            pix.save(tmp_image)
            lbl = QLabel()
            lbl.setPixmap(QPixmap(tmp_image))
            lbl.setScaledContents(True)
            lbl.setFixedSize(650, 800)
            layout.addWidget(lbl)
        except Exception as e:
            layout.addWidget(QLabel(f"Preview not available: {e}"))

        btn_save = QPushButton("Save PDF As...")
        btn_save.setFixedHeight(40)
        btn_save.clicked.connect(lambda: self.save_pdf(pdf_path))
        layout.addWidget(btn_save)

        btn_close = QPushButton("Close")
        btn_close.setFixedHeight(40)
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

        self.setLayout(layout)

    def save_pdf(self, src_path):
        new_name, _ = QFileDialog.getSaveFileName(self, "Save PDF", "document.pdf", "PDF Files (*.pdf)")
        if new_name:
            shutil.copy(src_path, new_name)
            QMessageBox.information(self, "Saved", "PDF saved successfully!")

# ----------------------------
# Main Application Window
# ----------------------------
class TestimonialApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Official Testimonial Generator")
        self.setMinimumSize(900, 720)
        self.db = StudentDatabase()

        main_layout = QVBoxLayout()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QFrame()
        self.scroll_layout = QVBoxLayout(content)

        # Title
        title = QLabel("Testimonial & Transfer Certificate Generator (Excel-based)")
        title.setFont(QFont("Times New Roman", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        self.scroll_layout.addWidget(title)

        self.scroll_layout.addSpacing(8)

        # Top buttons
        top_btns = QHBoxLayout()
        self.load_btn = QPushButton("Load Excel (Upload)")
        self.load_btn.clicked.connect(self.load_excel)
        self.save_btn = QPushButton("Save Excel")
        self.save_btn.clicked.connect(self.save_excel)
        self.saveas_btn = QPushButton("Save Excel As...")
        self.saveas_btn.clicked.connect(self.saveas_excel)
        self.new_serial_btn = QPushButton("New Serial")
        self.new_serial_btn.clicked.connect(self.fill_new_serial)

        top_btns.addWidget(self.load_btn)
        top_btns.addWidget(self.save_btn)
        top_btns.addWidget(self.saveas_btn)
        top_btns.addStretch()
        top_btns.addWidget(self.new_serial_btn)
        self.scroll_layout.addLayout(top_btns)

        self.scroll_layout.addSpacing(10)

        # Form area
        form = QGridLayout()
        labels = [
            "S/N", "Date (DD/MM/YYYY)", "ID No",
            "Class", "Session", "Student Name",
            "Father's Name", "Mother's Name",
            "Date of Birth (DD/MM/YYYY)"
        ]
        self.inputs = []
        for i, text in enumerate(labels):
            lbl = QLabel(text)
            lbl.setFont(QFont("Times New Roman", 11))
            edit = QLineEdit()
            edit.setFont(QFont("Times New Roman", 11))
            edit.setFixedHeight(28)
            form.addWidget(lbl, i, 0)
            form.addWidget(edit, i, 1)
            self.inputs.append(edit)

        # Set today date automatically
        today_str = datetime.today().strftime("%d/%m/%Y")
        self.inputs[1].setText(today_str)

        # Gender dropdown
        gender_label = QLabel("Select Gender:")
        gender_label.setFont(QFont("Times New Roman", 12))
        self.gender_box = QComboBox()
        self.gender_box.addItems(["Male", "Female"])
        self.gender_box.setFixedWidth(160)
        form.addWidget(gender_label, 0, 2)
        form.addWidget(self.gender_box, 0, 3)

        self.scroll_layout.addLayout(form)

        # Wire up ID auto-fill (except DOB)
        self.inputs[2].textChanged.connect(self.on_id_changed)

        self.scroll_layout.addSpacing(8)

        # Buttons
        btns = QHBoxLayout()
        self.generate_btn = QPushButton("Generate Testimonial Certificate")
        self.generate_btn.clicked.connect(self.generate_pdf)
        self.preview_btn = QPushButton("Preview Last PDF")
        self.preview_btn.clicked.connect(self.preview_last_pdf)
        self.open_pdf_btn = QPushButton("Open Last PDF")
        self.open_pdf_btn.clicked.connect(self.open_last_pdf)

        # TC Button
        self.tc_generate_btn = QPushButton("Generate Transfer Certificate")
        self.tc_generate_btn.clicked.connect(self.generate_tc_pdf)

        btns.addStretch()
        btns.addWidget(self.generate_btn)
        btns.addWidget(self.tc_generate_btn)
        btns.addWidget(self.preview_btn)
        btns.addWidget(self.open_pdf_btn)
        btns.addStretch()

        self.scroll_layout.addLayout(btns)
        self.scroll_layout.addSpacing(12)

        # Table view
        table_label = QLabel("Excel Data (editable) — edit cells and click Save Excel to persist")
        table_label.setFont(QFont("Times New Roman", 11, QFont.Bold))
        self.scroll_layout.addWidget(table_label)

        self.table = QTableWidget()
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.itemChanged.connect(self.on_table_item_changed)

        self.scroll_layout.addWidget(self.table, stretch=1)

        footer = QLabel("Developed by Md Shahriar Hasan Sabuj – with the help of ChatGPT")
        footer.setFont(QFont("Times New Roman", 10, QFont.Bold))
        footer.setAlignment(Qt.AlignCenter)
        self.scroll_layout.addWidget(footer)

        scroll.setWidget(content)
        main_layout.addWidget(scroll)
        self.setLayout(main_layout)

        self.last_pdf = None

    # ---------------------
    # Excel functions
    # ---------------------
    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        try:
            self.db.load_excel(path, copy_to_storage=True)
            self.refresh_table()
            QMessageBox.information(self, "Loaded", f"Loaded Excel and stored as: {os.path.basename(self.db.filepath)}")
            self.inputs[0].setText(str(self.db.get_next_serial()))
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not load Excel:\n{e}")

    def save_excel(self):
        if not self.db.filepath:
            self.saveas_excel()
            return
        try:
            self.sync_table_to_df()
            self.db.save_excel()
            QMessageBox.information(self, "Saved", f"Saved to {os.path.basename(self.db.filepath)}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not save Excel:\n{e}")

    def saveas_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel As", "students.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            self.sync_table_to_df()
            self.db.save_excel(path)
            QMessageBox.information(self, "Saved", f"Saved to {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not save Excel:\n{e}")

    # ---------------------
    # Table helpers
    # ---------------------
    def refresh_table(self):
        df = self.db.df
        self.table.blockSignals(True)
        self.table.clear()
        self.table.setColumnCount(len(df.columns))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels(df.columns.tolist())
        for r in range(len(df)):
            for c in range(len(df.columns)):
                val = "" if pd.isna(df.iat[r, c]) else str(df.iat[r, c])
                item = QTableWidgetItem(val)
                self.table.setItem(r, c, item)
        self.table.blockSignals(False)

    def sync_table_to_df(self):
        cols = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
        rows = []
        for r in range(self.table.rowCount()):
            row = []
            for c in range(self.table.columnCount()):
                it = self.table.item(r, c)
                row.append(it.text() if it else "")
            rows.append(row)
        df = pd.DataFrame(rows, columns=cols)
        # Serial numeric
        if "Serial" in df.columns:
            try:
                df["Serial"] = pd.to_numeric(df["Serial"], errors="coerce").fillna(0).astype(int)
            except:
                pass
        # ID as integer string
        if "ID" in df.columns:
            df["ID"] = df["ID"].apply(lambda x: str(int(float(x))) if str(x).replace('.', '', 1).isdigit() else str(x))
        self.db.df = df

    def on_table_item_changed(self, item):
        r = item.row()
        c = item.column()
        colname = self.table.horizontalHeaderItem(c).text()
        try:
            self.db.df.at[r, colname] = item.text()
        except:
            pass

    # ---------------------
    # ID auto-fill (except DOB)
    # ---------------------
    def on_id_changed(self):
        sid = self.inputs[2].text().strip()
        if not sid:
            return
        rec = self.db.get_student_by_id(sid)
        if rec:
            self.inputs[0].setText(str(rec.get("Serial", "")))
            self.inputs[5].setText(str(rec.get("Name", "")))
            self.inputs[6].setText(str(rec.get("Father", "")))
            self.inputs[7].setText(str(rec.get("Mother", "")))
            self.inputs[3].setText(str(rec.get("Class", "")))
            self.inputs[4].setText(str(rec.get("Session", "")))
        else:
            self.inputs[0].setText(str(self.db.get_next_serial()))

    def fill_new_serial(self):
        self.inputs[0].setText(str(self.db.get_next_serial()))
        QMessageBox.information(self, "Serial", f"Next serial set to {self.inputs[0].text()}")

    # ---------------------
    # PDF Generation
    # ---------------------
    def generate_pdf(self):
        vals = [i.text().strip() for i in self.inputs]
        if any(v == "" for v in [vals[0], vals[1], vals[2], vals[5]]):
            QMessageBox.warning(self, "Error", "Please ensure at least S/N, Date, ID and Student Name are filled.")
            return

        sn, date, student_id, student_class, session, name, father, mother, dob = vals
        gender = self.gender_box.currentText().lower()
        if gender == "male":
            he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
        else:
            he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

        entry = {
            "Serial": int(sn) if str(sn).isdigit() else sn,
            "ID": student_id,
            "Name": name,
            "Father": father,
            "Mother": mother,
            "Class": student_class,
            "Session": session,
            "DOB": dob
        }
        try:
            self.db.upsert_student(entry)
            self.refresh_table()
            self.db.save_excel()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not upsert student data:\n{e}")
            return

        pdf_path = f"testimonial_{student_id}.pdf"
        try:
            self._create_pdf(pdf_path, entry, gender)
            self.last_pdf = pdf_path
            QMessageBox.information(self, "PDF Generated", f"PDF saved as {pdf_path}")
        except Exception as e:
            QMessageBox.warning(self, "PDF Error", f"Could not create PDF:\n{e}")

    def _create_pdf(self, pdf_path, entry, gender):
        sn = entry["Serial"]
        date = self.inputs[1].text().strip()
        student_id = entry["ID"]
        student_class = entry["Class"]
        session = entry["Session"]
        name = entry["Name"]
        father = entry["Father"]
        mother = entry["Mother"]
        dob = entry["DOB"]

        if gender == "male":
            he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
        else:
            he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

        c = canvas.Canvas(pdf_path, pagesize=A4)
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
        c.drawCentredString(W/2, heading_y + heading_h/2 - 6, "Testimonial Certificate")

        # Left table
        table_x = left
        table_y_top = heading_y - 20*mm
        cell_w1 = 30*mm
        cell_w2 = 55*mm
        cell_h = 9*mm

        c.setFont("Times-Roman", 11)
        keys = ["S/N", "Date", "ID No", "Class", "Session"]
        vals = [str(sn), date, student_id, student_class, session]
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

        # Paragraph (JUSTIFIED, auto-adjusted)
        paragraph = (
            f"{name} {son_daughter} of {father} and {mother} is a student of Class: {student_class}. "
            f"Bearing ID/Roll: {student_id} in Daffodil University School & College. "
            f"As per our admission record {his_her} date of birth is {dob}. "
            f"To the best of my knowledge {he_she} was well mannered and possessed a good moral character. "
            f"{He_She} did not indulge {Him_Her}self in any activity subversive to the state and discipline during study. "
            f"I wish {Him_Her} every success in life!"
        )

        # signature baseline (same as before)
        sig_y = 110*mm

        # Paragraph style & frame calculation
        style = ParagraphStyle(
            name="Justify",
            fontName="Times-Roman",
            fontSize=11,
            leading=14,
            alignment=TA_JUSTIFY,
        )

        p = Paragraph(paragraph, style)

        # Frame: place between intro_y and signature line, with small margins.
        # Calculate frame bottom (a little above signature) and height dynamically.
        frame_bottom = sig_y + 15 * mm
        frame_top = intro_y - 10  # intro_y is in points; minus a small offset
        frame_height = max(30 * mm, frame_top - frame_bottom)  # ensure some minimum height
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
        c.setFont("Times-Roman", 11)
        text_lines = ["SK Mahmudun Nabi", "Principal (Acting)", "Daffodil University School & College"]
        for i, line in enumerate(text_lines):
            c.drawString(left, sig_y-12-i*12, line)

        c.save()

    # ---------------------
    # Transfer Certificate generation (new)
    # ---------------------
    def generate_tc_pdf(self):
        vals = [i.text().strip() for i in self.inputs]
        if any(v == "" for v in [vals[0], vals[1], vals[2], vals[5]]):
            QMessageBox.warning(self, "Error", "Please ensure S/N, Date, ID and Student Name are filled.")
            return

        sn, date, student_id, student_class, session, name, father, mother, dob = vals
        gender = self.gender_box.currentText().lower()
        if gender == "male":
            he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
        else:
            he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

        entry = {
            "Serial": int(sn) if str(sn).isdigit() else sn,
            "ID": student_id,
            "Name": name,
            "Father": father,
            "Mother": mother,
            "Class": student_class,
            "Session": session,
            "DOB": dob
        }
        try:
            self.db.upsert_student(entry)
            self.refresh_table()
            self.db.save_excel()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not upsert data:\n{e}")
            return

        pdf_path = f"transfer_certificate_{student_id}.pdf"
        try:
            self._create_tc_pdf(pdf_path, entry, gender)
            self.last_pdf = pdf_path
            QMessageBox.information(self, "TC PDF Generated", f"TC PDF saved as {pdf_path}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not create TC:\n{e}")

    def _create_tc_pdf(self, pdf_path, entry, gender):
        sn = entry["Serial"]
        date = self.inputs[1].text().strip()
        student_id = entry["ID"]
        student_class = entry["Class"]
        session = entry["Session"]
        name = entry["Name"]
        father = entry["Father"]
        mother = entry["Mother"]
        dob = entry["DOB"]

        if gender == "male":
            he_she, He_She, his_her, Him_Her, son_daughter = "he","He","his","him","son"
        else:
            he_she, He_She, his_her, Him_Her, son_daughter = "she","She","her","her","daughter"

        c = canvas.Canvas(pdf_path, pagesize=A4)
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
        c.setFont("Times-Roman", 17)
        c.drawCentredString(W/2, heading_y + heading_h/2 - 6, "Transfer Certificate")

        # Left table
        table_x = left
        table_y_top = heading_y - 20*mm
        cell_w1 = 30*mm
        cell_w2 = 55*mm
        cell_h = 9*mm

        c.setFont("Times-Roman", 11)
        keys = ["S/N", "Date", "ID No", "Class", "Session"]
        vals = [str(sn), date, student_id, student_class, session]
        for i, key in enumerate(keys):
            y = table_y_top - i*cell_h
            c.rect(table_x, y-cell_h, cell_w1, cell_h)
            c.rect(table_x+cell_w1, y-cell_h, cell_w2, cell_h)
            c.drawString(table_x+3, y-cell_h/2+2, key)
            c.drawString(table_x+cell_w1+4, y-cell_h/2+2, vals[i])

        # Intro
        intro_y = table_y_top - len(keys)*cell_h - 10*mm
        c.setFont("Times-Roman", 17)
        c.drawCentredString(W/2, intro_y, "This is to certify that")

        # Transfer Certificate paragraph (JUSTIFIED, auto-adjusted)
        paragraph = (
            f"{name}, {son_daughter} of {father} and {mother}, "
            f"was a student of Class {student_class} (Bearing ID/Roll: {student_id}) at "
            f"Daffodil University School & College. As per our record, {his_her} date of birth "
            f"is {dob}. During {his_her} stay, {he_she} maintained good conduct and discipline. "
            f"We wish {Him_Her} every success in future life."
        )

        # signature baseline (same as before)
        sig_y = 110*mm

        style = ParagraphStyle(
            name="JustifyTC",
            fontName="Times-Roman",
            fontSize=11,
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
        c.setFont("Times-Roman", 11)
        text_lines = ["SK Mahmudun Nabi", "Principal (Acting)", "Daffodil University School & College"]
        for i, line in enumerate(text_lines):
            c.drawString(left, sig_y-12-i*12, line)

        c.save()

    def preview_last_pdf(self):
        if not self.last_pdf or not os.path.exists(self.last_pdf):
            QMessageBox.information(self, "No PDF", "No PDF generated yet.")
            return
        pv = PDFPreview(self.last_pdf, self)
        pv.show()

    def open_last_pdf(self):
        if not self.last_pdf or not os.path.exists(self.last_pdf):
            QMessageBox.information(self, "No PDF", "No PDF generated yet.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.last_pdf)
            elif sys.platform == "darwin":
                os.system(f"open '{self.last_pdf}'")
            else:
                os.system(f"xdg-open '{self.last_pdf}'")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open PDF:\n{e}")

# ----------------------------
# Run Application
# ----------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TestimonialApp()
    window.show()
    sys.exit(app.exec_())
