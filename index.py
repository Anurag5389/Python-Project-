import os
import re
import sqlite3
from tkinter import *
from tkinter import filedialog, simpledialog
import tkinter.ttk as ttk
import tkinter.messagebox as mb

DB_FILE = "student_reg.db"

# ------------------------- DB LAYER -------------------------
class DB:
    def __init__(self, path=DB_FILE):
        self.path = path
        self._ensure_db()

    def connect(self):
        return sqlite3.connect(self.path)

    def _ensure_column(self, table, coldef):
        # coldef like "photo_path TEXT"
        col = coldef.split()[0]
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute(f"PRAGMA table_info({table})")
            cols = [r[1].lower() for r in cur.fetchall()]
            if col.lower() not in cols:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {coldef}")
                conn.commit()

    def _ensure_db(self):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS student (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    roll_number TEXT NOT NULL UNIQUE,
                    firstname TEXT,
                    lastname TEXT,
                    gender TEXT,
                    age INTEGER,
                    address TEXT,
                    contact TEXT,
                    created_at TEXT DEFAULT (datetime('now'))
                )
            """)
            conn.commit()
        # optional column (safe to call repeatedly)
        self._ensure_column("student", "photo_path TEXT")

    def fetch_all(self, order_by="lastname"):
        with self.connect() as conn:
            cur = conn.cursor()
            try:
                cur.execute(f"""
                    SELECT id, roll_number, firstname, lastname, gender, age, address, contact, photo_path
                    FROM student ORDER BY {order_by} ASC
                """)
            except sqlite3.OperationalError:
                cur.execute("""
                    SELECT id, roll_number, firstname, lastname, gender, age, address, contact, photo_path
                    FROM student ORDER BY lastname ASC
                """)
            return cur.fetchall()

    def get_by_roll(self, roll_number: str):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, roll_number, firstname, lastname, gender, age, address, contact, photo_path
                FROM student
                WHERE roll_number=?
            """, (roll_number,))
            return cur.fetchone()

    def insert(self, row):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO student (roll_number, firstname, lastname, gender, age, address, contact, photo_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row.get("roll_number"),
                row.get("firstname"),
                row.get("lastname"),
                row.get("gender"),
                int(row.get("age")) if str(row.get("age")) else None,
                row.get("address"),
                row.get("contact"),
                row.get("photo_path"),
            ))
            conn.commit()

    def upsert_by_roll(self, row):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO student (roll_number, firstname, lastname, gender, age, address, contact, photo_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(roll_number) DO UPDATE SET
                    firstname=excluded.firstname,
                    lastname=excluded.lastname,
                    gender=excluded.gender,
                    age=excluded.age,
                    address=excluded.address,
                    contact=excluded.contact,
                    photo_path=excluded.photo_path
            """, (
                row.get("roll_number"),
                row.get("firstname"),
                row.get("lastname"),
                row.get("gender"),
                int(row.get("age")) if str(row.get("age")) else None,
                row.get("address"),
                row.get("contact"),
                row.get("photo_path"),
            ))
            conn.commit()

    def update(self, sid, row):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE student
                SET roll_number=?, firstname=?, lastname=?, gender=?, age=?, address=?, contact=?, photo_path=?
                WHERE id=?
            """, (
                row.get("roll_number"),
                row.get("firstname"),
                row.get("lastname"),
                row.get("gender"),
                int(row.get("age")) if str(row.get("age")) else None,
                row.get("address"),
                row.get("contact"),
                row.get("photo_path"),
                sid,
            ))
            conn.commit()

    def delete(self, sid):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM student WHERE id=?", (sid,))
            conn.commit()

# ------------------------- UI LAYER -------------------------
class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Student Registration Module")
        self.geometry("1020x680")
        self.configure(bg="#eaf2ff")
        self.db = DB()

        self.vars = {k: StringVar() for k in [
            "roll_number","firstname","lastname","gender","age","address","contact","photo_path"
        ]}
        self.vars["gender"].set("Male")

        self._style()
        self._build_ui()
        self.refresh_table()

    # ---------- Theme / Style ----------
    def _style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#eaf2ff")
        style.configure("TLabel", background="#eaf2ff", foreground="#0f1d40")
        style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), background="#d7e6ff", foreground="#0b1a39")
        style.configure("TButton", padding=6)
        style.map("TButton", background=[("!disabled", "#2f6fff")], foreground=[("!disabled", "white")])
        style.configure("Danger.TButton", background="#e74c3c", foreground="white")
        style.map("Danger.TButton", background=[("!disabled", "#e74c3c")])
        style.configure("Treeview", background="white", foreground="#0b1a39", rowheight=24, fieldbackground="white")
        style.configure("Treeview.Heading", background="#b8d0ff", foreground="#0b1a39", font=("Segoe UI", 10, "bold"))

    # ---------- UI construction ----------
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(side=TOP, fill=X)
        ttk.Label(top, text="Student Registration", style="Header.TLabel", anchor="center").pack(fill=X, ipady=8)

        form = ttk.Frame(self, padding=(10, 8))
        form.pack(side=TOP, fill=X)
        self._add_form_row(form, 0, "Roll Number*", "roll_number")
        self._add_form_row(form, 1, "First Name*", "firstname")
        self._add_form_row(form, 2, "Last Name*", "lastname")

        ttk.Label(form, text="Gender").grid(row=3, column=0, sticky=E, padx=6, pady=3)
        gframe = ttk.Frame(form)
        gframe.grid(row=3, column=1, sticky=W, padx=6, pady=3)
        ttk.Radiobutton(gframe, text="Male", value="Male", variable=self.vars["gender"]).pack(side=LEFT, padx=(0,10))
        ttk.Radiobutton(gframe, text="Female", value="Female", variable=self.vars["gender"]).pack(side=LEFT)

        self._add_form_row(form, 4, "Age", "age")
        self._add_form_row(form, 5, "Address", "address")
        self._add_form_row(form, 6, "Contact", "contact")
        self._add_form_row(form, 7, "Photo (optional)", "photo_path")

        browse_frame = ttk.Frame(form)
        browse_frame.grid(row=7, column=2, sticky=W, padx=4)
        ttk.Button(browse_frame, text="Browse...", command=self._browse_photo).pack()

        btns = ttk.Frame(form)
        btns.grid(row=0, column=3, rowspan=5, padx=14, pady=4, sticky=N)
        ttk.Button(btns, text="Save New", width=18, command=self.save_new).pack(pady=4)
        ttk.Button(btns, text="Update Selected", width=18, command=self.update_selected).pack(pady=4)
        ttk.Button(btns, text="Delete Selected", width=18, style="Danger.TButton", command=self.delete_selected).pack(pady=4)
        ttk.Button(btns, text="Clear Form", width=18, command=self.clear_form).pack(pady=4)

        io = ttk.Frame(self)
        io.pack(side=TOP, fill=X, padx=10, pady=6)
        ttk.Button(io, text="Import from Excel (.xlsx)", command=self.import_excel).pack(side=LEFT)
        ttk.Button(io, text="Export to Excel (.xlsx)", command=self.export_excel).pack(side=LEFT, padx=8)
        ttk.Button(io, text="Export Table PDF", command=self.export_pdf).pack(side=LEFT)
        ttk.Button(io, text="All ID Cards (PDF)", command=self.print_all_id_cards).pack(side=LEFT, padx=8)
        ttk.Button(io, text="Single ID Card (PDF)", command=self.print_single_id_card).pack(side=LEFT)
        ttk.Button(io, text="Download Excel Template", command=self.save_template).pack(side=RIGHT)

        table_frame = ttk.Frame(self)
        table_frame.pack(fill=BOTH, expand=1, padx=10, pady=10)
        cols = ("id","roll_number","firstname","lastname","gender","age","address","contact","photo_path")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c)
            width = 80
            if c in ("address",): width = 220
            if c in ("firstname","lastname","roll_number"): width = 140
            if c in ("photo_path",): width = 220
            self.tree.column(c, width=width, anchor=W)
        vsb = ttk.Scrollbar(table_frame, orient=VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=1)
        vsb.pack(side=LEFT, fill=Y)
        hsb.pack(side=BOTTOM, fill=X)
        self.tree.bind("<ButtonRelease-1>", self.on_select)

        ttk.Label(self, text=f"DB: {os.path.abspath(self.db.path)}", anchor=W).pack(side=BOTTOM, fill=X)

    def _add_form_row(self, parent, r, label, key):
        ttk.Label(parent, text=label, width=16).grid(row=r, column=0, sticky=E, padx=6, pady=3)
        ttk.Entry(parent, textvariable=self.vars[key], width=30).grid(row=r, column=1, sticky=W, padx=6, pady=3)

    # ---------- Helpers ----------
    def _browse_photo(self):
        path = filedialog.askopenfilename(
            title="Select Photo",
            filetypes=[("Images","*.png;*.jpg;*.jpeg;*.gif;*.webp;*.bmp"), ("All","*.*")]
        )
        if path:
            self.vars["photo_path"].set(path)

    # ---------- CRUD actions ----------
    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for row in self.db.fetch_all(order_by="roll_number"):
            self.tree.insert("", "end", values=row)

    def form_payload(self):
        return {k: v.get().strip() for k, v in self.vars.items()}

    # --------- Validations ----------
    def validate(self, data, allow_blank_roll=False):
        roll = data.get("roll_number", "")
        if not roll and not allow_blank_roll:
            mb.showwarning("Validation", "Roll Number is required.")
            return False
        if roll and not re.fullmatch(r"[A-Za-z0-9_\-\/]{1,20}", roll):
            mb.showwarning("Validation", "Roll Number can only have letters, digits, -, _, / (max 20 chars).")
            return False
        fn = data.get("firstname", "")
        ln = data.get("lastname", "")
        if not fn:
            mb.showwarning("Validation", "First Name is required.")
            return False
        if not ln:
            mb.showwarning("Validation", "Last Name is required.")
            return False
        if not re.fullmatch(r"[A-Za-z][A-Za-z ]{0,49}", fn):
            mb.showwarning("Validation", "First Name must be letters/spaces (max 50).")
            return False
        if not re.fullmatch(r"[A-Za-z][A-Za-z ]{0,49}", ln):
            mb.showwarning("Validation", "Last Name must be letters/spaces (max 50).")
            return False
        gender = data.get("gender", "")
        if gender not in ("Male", "Female"):
            mb.showwarning("Validation", "Please select Gender (Male/Female).")
            return False
        if data.get("age"):
            if not data["age"].isdigit():
                mb.showwarning("Validation", "Age must be a number.")
                return False
            if not (1 <= int(data["age"]) <= 120):
                mb.showwarning("Validation", "Age must be between 1 and 120.")
                return False
        if len(data.get("address","")) > 200:
            mb.showwarning("Validation", "Address too long (max 200 chars).")
            return False
        contact = data.get("contact", "")
        if contact:
            if not re.fullmatch(r"\d{10}", contact):
                mb.showwarning("Validation", "Contact must be exactly 10 digits.")
                return False
        return True

    def save_new(self):
        data = self.form_payload()
        if not self.validate(data):
            return
        try:
            self.db.insert(data)
            self.refresh_table()
            self.clear_form()
            mb.showinfo("Saved", "Student registered successfully.")
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed: student.roll_number" in str(e):
                mb.showerror("Duplicate", "This Roll Number already exists.")
            else:
                mb.showerror("DB Error", str(e))

    def on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0])['values']
        keys = ["id","roll_number","firstname","lastname","gender","age","address","contact","photo_path"]
        data = dict(zip(keys, vals))
        self.vars["roll_number"].set(data.get("roll_number",""))
        self.vars["firstname"].set(data.get("firstname",""))
        self.vars["lastname"].set(data.get("lastname",""))
        self.vars["gender"].set(data.get("gender","") or "Male")
        self.vars["age"].set(str(data.get("age","") if data.get("age") is not None else ""))
        self.vars["address"].set(data.get("address",""))
        self.vars["contact"].set(data.get("contact",""))
        self.vars["photo_path"].set(data.get("photo_path",""))

    def get_selected_id(self):
        sel = self.tree.selection()
        if not sel:
            mb.showwarning("Select", "Please select a row in the table.")
            return None
        return self.tree.item(sel[0])['values'][0]

    def update_selected(self):
        sid = self.get_selected_id()
        if not sid:
            return
        data = self.form_payload()
        if not self.validate(data):
            return
        try:
            self.db.update(sid, data)
            self.refresh_table()
            mb.showinfo("Updated", "Record updated.")
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed: student.roll_number" in str(e):
                mb.showerror("Duplicate", "Another record already has this Roll Number.")
            else:
                mb.showerror("DB Error", str(e))

    def delete_selected(self):
        sid = self.get_selected_id()
        if not sid:
            return
        if not mb.askyesno("Confirm", "Delete selected record?"):
            return
        self.db.delete(sid)
        self.refresh_table()
        self.clear_form()

    def clear_form(self):
        for k, v in self.vars.items():
            v.set("")
        self.vars["gender"].set("Male")

    # ---------- Import/Export ----------
    def import_excel(self):
        path = filedialog.askopenfilename(title="Select Excel file (.xlsx)", filetypes=[("Excel .xlsx", "*.xlsx")])
        if not path:
            return
        try:
            import pandas as pd
        except Exception:
            mb.showerror("Missing package", "Please install pandas and openpyxl:\n\npip install pandas openpyxl")
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            mb.showerror("Read error", f"Unable to read Excel: {e}")
            return
        df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
        required = ["roll_number","firstname","lastname","gender","age","address","contact"]
        has_photo = "photo_path" in df.columns
        missing = [c for c in required if c not in df.columns]
        if missing:
            mb.showerror("Columns missing", f"Excel must contain columns: {', '.join(required)}\nMissing: {', '.join(missing)}")
            return
        count = 0
        for _, r in df.iterrows():
            row = {k: ("" if pd.isna(r.get(k)) else str(r.get(k))).strip() for k in required}
            row["photo_path"] = ("" if not has_photo or pd.isna(r.get("photo_path")) else str(r.get("photo_path"))).strip()
            g = row.get("gender","").lower()
            row["gender"] = "Male" if g.startswith("m") else ("Female" if g.startswith("f") else "")
            row["age"] = int(row["age"]) if row["age"].isdigit() else None
            if not row["roll_number"] or not row["firstname"] or not row["lastname"] or row["gender"] not in ("Male","Female"):
                continue
            try:
                self.db.upsert_by_roll(row)
                count += 1
            except Exception as e:
                print("Failed row:", row, e)
        self.refresh_table()
        mb.showinfo("Import complete", f"Imported/updated {count} rows.")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel .xlsx","*.xlsx")], title="Save Excel")
        if not path:
            return
        try:
            import pandas as pd
        except Exception:
            mb.showerror("Missing package", "Please install pandas and openpyxl:\n\npip install pandas openpyxl")
            return
        rows = self.db.fetch_all(order_by="roll_number")
        cols = ["id","roll_number","firstname","lastname","gender","age","address","contact","photo_path"]
        df = pd.DataFrame(rows, columns=cols)
        try:
            df.to_excel(path, index=False)
            mb.showinfo("Exported", f"Saved to {path}")
        except Exception as e:
            mb.showerror("Write error", str(e))

    def save_template(self):
        path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel .xlsx","*.xlsx")],
                title="Save Template"
            )
        if not path:
            return
        try:
            import pandas as pd
        except Exception:
            mb.showerror("Missing package",
                         "Please install pandas and openpyxl:\n\npip install pandas openpyxl")
            return
        cols = ["roll_number","firstname","lastname","gender","age","address","contact","photo_path"]
        try:
            pd.DataFrame(columns=cols).to_excel(path, index=False)
            mb.showinfo("Template saved", f"Empty template saved to {path}")
        except Exception as e:
            mb.showerror("Write error", str(e))


    def export_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF","*.pdf")], title="Save PDF")
        if not path:
            return
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas
            from reportlab.lib import colors
        except Exception:
            mb.showerror("Missing package", "Please install reportlab:\n\npip install reportlab")
            return
        rows = self.db.fetch_all(order_by="roll_number")
        cols = ["Roll No","First Name","Last Name","Gender","Age","Address","Contact","Photo"]
        c = canvas.Canvas(path, pagesize=A4)
        width, height = A4
        left = 40
        top = height - 40
        line_h = 18
        title = "Student Registration Data"
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, top, title)
        c.setFont("Helvetica", 9)
        y = top - 30
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 9)
        x_positions = [left, 100, 200, 300, 370, 420, 500, 560]
        for x, h in zip(x_positions, cols):
            c.drawString(x, y, h)
        y -= 10
        c.line(left, y, width - left, y)
        y -= 12
        c.setFont("Helvetica", 9)
        for r in rows:
            _, roll, fn, ln, g, age, addr, contact, photo_path = r
            values = [str(roll or ''), str(fn or ''), str(ln or ''), str(g or ''), str(age or ''), str(addr or ''), str(contact or ''), "✓" if photo_path else ""]
            for x, val in zip(x_positions, values):
                c.drawString(x, y, val[:40])
            y -= line_h
            if y < 60:
                c.showPage()
                c.setFont("Helvetica-Bold", 14)
                c.drawString(left, height - 40, title)
                c.setFont("Helvetica-Bold", 9)
                y = height - 70
                for x, h in zip(x_positions, cols):
                    c.drawString(x, y, h)
                y -= 10
                c.line(left, y, width - left, y)
                y -= 12
                c.setFont("Helvetica", 9)
        c.save()
        mb.showinfo("Exported", f"Saved to {path}")

    # ---------- ID CARD RENDERING CORE ----------
    def _render_id_cards(self, rows, path):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.pdfgen import canvas
            from reportlab.lib import colors
            from reportlab.graphics import renderPDF
            from reportlab.graphics.barcode import qr
        except Exception:
            mb.showerror("Missing package", "Please install reportlab:\n\npip install reportlab")
            return

        norm_rows = [list(r) for r in rows]

        page_w, page_h = landscape(A4)
        c = canvas.Canvas(path, pagesize=(page_w, page_h))

        margin = 24
        cols = 3
        rows_per_page = 2
        gap_x = 18
        gap_y = 18

        card_w = (page_w - 2*margin - (cols-1)*gap_x) / cols
        card_h = (page_h - 2*margin - (rows_per_page-1)*gap_y) / rows_per_page
        radius = 8

        def draw_round_rect(x, y, w, h, r, stroke=1):
            c.roundRect(x, y, w, h, r, stroke=stroke, fill=0)

        def draw_card(x, y, data):
            row = list(data) if isinstance(data, (tuple, list)) else [data]
            row += [""] * max(0, 9 - len(row))
            _id, roll, fn, ln, gender, age, addr, contact, photo_path = row[:9]

            title_bg = colors.Color(0.16, 0.43, 1.0)   # blue
            text_dark = colors.HexColor("#0b1a39")

            draw_round_rect(x, y, card_w, card_h, radius)
            c.setFillColor(title_bg)
            c.roundRect(x, y+card_h-28, card_w, 28, radius, stroke=0, fill=1)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(x+12, y+card_h-20, "STUDENT ID CARD")

            # photo (optional)
            photo_size = 60
            px = x + card_w - photo_size - 12
            py = y + card_h - 12 - photo_size
            if photo_path and os.path.exists(str(photo_path)):
                try:
                    c.drawImage(photo_path, px, py, width=photo_size, height=photo_size,
                                preserveAspectRatio=True, mask='auto')
                except Exception:
                    pass

            # text block
            c.setFillColor(text_dark)
            c.setFont("Helvetica-Bold", 11)
            c.drawString(x+12, y+card_h-50, f"Name: {fn} {ln}")
            c.setFont("Helvetica", 10)
            c.drawString(x+12, y+card_h-68, f"Roll: {roll}")
            c.drawString(x+12, y+card_h-86, f"Gender: {gender}    Age: {'' if age in (None,'') else age}")
            c.drawString(x+12, y+card_h-104, f"Contact: {contact}")
            addr_str = (str(addr or ""))[:48]
            c.drawString(x+12, y+card_h-122, f"Address: {addr_str}")

            # QR (roll number as payload)
            try:
                qr_code = qr.QrCodeWidget(str(roll))
                b = qr_code.getBounds()
                qr_w = 100
                qr_h = 100
                w = b[2]-b[0]; h = b[3]-b[1]
                from reportlab.graphics.shapes import Drawing, Rect
                d = Drawing(qr_w, qr_h, transform=[qr_w/w,0,0,qr_h/h,0,0])
                d.add(Rect(0, 0, qr_w, qr_h, fillColor=colors.white, strokeColor=None))
                d.add(qr_code)
                from reportlab.graphics import renderPDF
                renderPDF.draw(d, c, x + card_w - qr_w - 12, y + 12)
            except Exception:
                pass

          

        col = row = 0
        for rec in norm_rows:
            cx = margin + col * (card_w + gap_x)
            cy = margin + (rows_per_page-1-row) * (card_h + gap_y)
            draw_card(cx, cy, rec)
            col += 1
            if col >= cols:
                col = 0
                row += 1
            if row >= rows_per_page:
                c.showPage()
                row = 0
                col = 0

        c.save()

    # ---------- Two new buttons ----------
    def print_all_id_cards(self):
        rows = self.db.fetch_all(order_by="roll_number")
        if not rows:
            mb.showwarning("No Data", "No students to print.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF","*.pdf")],
            title="Save ID Cards PDF"
        )
        if not path:
            return
        self._render_id_cards(rows, path)
        mb.showinfo("ID Cards", f"Saved to {path}")

    def print_single_id_card(self):
        # prefer selected row
        selected = self.tree.selection()
        if selected:
            rec = self.tree.item(selected[0])["values"]
        else:
            roll = simpledialog.askstring("Single ID Card", "Enter Roll Number:")
            if not roll:
                return
            rec = self.db.get_by_roll(roll.strip())
            if not rec:
                mb.showerror("Not Found", f"No student with roll number '{roll}'.")
                return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF","*.pdf")],
            title="Save Single ID Card PDF"
        )
        if not path:
            return
        self._render_id_cards([rec], path)
        mb.showinfo("ID Card", f"Saved to {path}")

   

if __name__ == "__main__":
    app = App()
    app.mainloop()
