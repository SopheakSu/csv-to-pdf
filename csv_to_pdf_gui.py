#!/usr/bin/env python3
"""
CSV to PDF Report Generator — offline desktop app.
No internet, no Telegram, no hosting required to run this.

Just fill in the fields and click Generate.
"""

import os
import re
import csv
import threading
import subprocess
import sys
from datetime import datetime
from collections import defaultdict

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

DATE_RE = re.compile(r'^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}$')


# ------------------------------------------------------------
# --- CSV parsing / PDF generation (same logic as before) ----
# ------------------------------------------------------------

def load_rows(csv_path):
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        return [r for r in reader if any(c.strip() for c in r)]


def parse_dispatch_fields(row):
    lo, hi = 15, 37
    segment = row[lo:hi] if len(row) >= hi else row[lo:]
    dates = [v.strip() for v in segment if DATE_RE.match(v.strip())]
    texts = [v.strip() for v in segment if v.strip() and re.search(r'[A-Za-z]', v)]
    command = dates[0] if dates else ''
    measure = dates[-1] if dates else ''
    action, operator_name, detail = '', '', ''
    if texts:
        detail = texts[-1]
        if len(texts) >= 2:
            first = texts[0]
            if ',' in first or not re.search(r'\b(Team|Respond|Center|ESS)\b', first, re.I):
                operator_name = first
            else:
                action = first
    remarks = row[37].strip() if len(row) > 37 else ''
    return command, action, operator_name, measure, detail, remarks


def build_records(rows, log=None):
    def sort_key(row):
        try:
            return datetime.strptime(row[0].strip(), '%Y/%m/%d %H:%M:%S')
        except Exception:
            return datetime.min

    rows = sorted(rows, key=sort_key)
    state = {}
    records = []
    for row in rows:
        row = row + [''] * max(0, 45 - len(row))
        receive_date, area = row[0].strip(), row[2].strip()
        cust_id, cust_name = row[3].strip(), row[4].strip()
        signal_code, signal_name = row[8].strip(), row[9].strip()
        key = (cust_id, area)
        if signal_name in ('Arm', 'Disarm'):
            status = signal_name
            state[key] = signal_name
            operator = row[12].strip()
            command = action = measure = detail = remarks = ''
        else:
            status = state.get(key, '')
            command, action, operator, measure, detail, remarks = parse_dispatch_fields(row)
        records.append({
            'receive_date': receive_date, 'area': area, 'cust_id': cust_id,
            'cust_name': cust_name, 'signal_code': signal_code, 'signal_name': signal_name,
            'status': status, 'operator': operator, 'command': command,
            'action': action, 'measure': measure, 'detail': detail, 'remarks': remarks,
        })
    return records


COLUMNS = [
    ('receive_date', 'Receive date', 0.72 * inch), ('area', 'Area', 0.42 * inch),
    ('cust_id', 'CustomerID', 0.95 * inch), ('cust_name', 'Customer name', 1.35 * inch),
    ('signal_code', 'Signal code', 0.72 * inch), ('signal_name', 'Signal name', 1.05 * inch),
    ('status', 'Status', 0.55 * inch), ('operator', 'Operator', 0.85 * inch),
    ('command', 'Command', 0.72 * inch), ('action', 'Action', 0.95 * inch),
    ('measure', 'Measure', 0.72 * inch), ('detail', 'Detail', 0.95 * inch),
]


def extract_id_number(cust_id):
    """Extract the meaningful numeric part of an ID like '01-002019-____' -> '2019'."""
    parts = cust_id.split('-')
    if len(parts) >= 2 and parts[1].strip():
        num = parts[1].strip().lstrip('0')
        return num if num else '0'
    digits = re.sub(r'\D', '', cust_id)
    return digits or 'na'


def sanitize_filename(name):
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    return name.strip()[:150] or 'unnamed'


def make_pdf(records, out_path, group_label):
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle('cell', parent=styles['Normal'], fontSize=6.5, leading=8)
    header_style = ParagraphStyle('hdr', parent=styles['Normal'], fontSize=7, leading=8,
                                   textColor=colors.white, fontName='Helvetica-Bold')
    title_style = ParagraphStyle('title', parent=styles['Normal'], fontSize=9, leading=12)

    doc = SimpleDocTemplate(out_path, pagesize=landscape(letter),
                             leftMargin=0.35 * inch, rightMargin=0.35 * inch,
                             topMargin=0.4 * inch, bottomMargin=0.4 * inch)

    dates = [r['receive_date'] for r in records if r['receive_date']]
    period_start = min(dates) if dates else ''
    period_end = max(dates) if dates else ''
    generated = datetime.now().strftime('%y/%m/%d %H:%M')

    story = [
        Paragraph(f"Period: {period_start} ~ {period_end}&nbsp;&nbsp;&nbsp; "
                  f"List by time&nbsp;&nbsp;&nbsp; Generated {generated}", title_style),
        Paragraph(f"<b>{group_label}</b>", title_style),
        Spacer(1, 6),
    ]

    header_row = [Paragraph(h, header_style) for _, h, _ in COLUMNS]
    table_data = [header_row]
    for r in records:
        table_data.append([Paragraph(str(r.get(key, '')), cell_style) for key, _, _ in COLUMNS])

    table = Table(table_data, colWidths=[w for _, _, w in COLUMNS], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f2f2f2')]),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(table)

    def add_page_number(canvas, doc_):
        canvas.saveState()
        canvas.setFont('Helvetica', 7)
        canvas.drawRightString(landscape(letter)[0] - 0.35 * inch, 0.25 * inch, f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)


FIELD_MAP = {'Customer Name': 'cust_name', 'Customer ID': 'cust_id', 'Area': 'area'}


# ------------------------------------------------------------
# --- GUI --------------------------------------------------
# ------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV to PDF Report Generator")
        self.geometry("640x480")
        self.resizable(False, False)

        pad = {'padx': 10, 'pady': 6}

        tk.Label(self, text="1. Choose your CSV file", font=('', 10, 'bold')).pack(anchor='w', **pad)
        row1 = tk.Frame(self); row1.pack(fill='x', **pad)
        self.csv_path_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.csv_path_var, width=60).pack(side='left', padx=(0, 6))
        tk.Button(row1, text="Browse...", command=self.browse_csv).pack(side='left')

        tk.Label(self, text="2. Filter by", font=('', 10, 'bold')).pack(anchor='w', **pad)
        row2 = tk.Frame(self); row2.pack(fill='x', **pad)
        self.field_var = tk.StringVar(value='Customer Name')
        ttk.Combobox(row2, textvariable=self.field_var, values=list(FIELD_MAP.keys()),
                     state='readonly', width=20).pack(side='left', padx=(0, 10))
        tk.Label(row2, text="Value (leave blank for ALL):").pack(side='left')
        self.value_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.value_var, width=28).pack(side='left', padx=(6, 0))

        tk.Label(self, text="3. Choose output folder", font=('', 10, 'bold')).pack(anchor='w', **pad)
        row3 = tk.Frame(self); row3.pack(fill='x', **pad)
        self.out_dir_var = tk.StringVar()
        tk.Entry(row3, textvariable=self.out_dir_var, width=60).pack(side='left', padx=(0, 6))
        tk.Button(row3, text="Browse...", command=self.browse_out_dir).pack(side='left')

        tk.Button(self, text="Generate PDFs", font=('', 11, 'bold'), bg='#2c7', fg='white',
                  command=self.start_generate).pack(pady=14)

        self.progress = ttk.Progressbar(self, mode='indeterminate', length=580)
        self.progress.pack(pady=(0, 8))

        tk.Label(self, text="Log:").pack(anchor='w', padx=10)
        self.log_box = tk.Text(self, height=12, width=78)
        self.log_box.pack(padx=10, pady=(0, 10))

    def log(self, msg):
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.update_idletasks()

    def browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            self.csv_path_var.set(path)
            if not self.out_dir_var.get():
                self.out_dir_var.set(os.path.join(os.path.dirname(path), "pdf_reports"))

    def browse_out_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.out_dir_var.set(path)

    def start_generate(self):
        csv_path = self.csv_path_var.get().strip()
        out_dir = self.out_dir_var.get().strip()
        if not csv_path or not os.path.isfile(csv_path):
            messagebox.showerror("Error", "Please choose a valid CSV file.")
            return
        if not out_dir:
            messagebox.showerror("Error", "Please choose an output folder.")
            return
        os.makedirs(out_dir, exist_ok=True)

        self.progress.start(10)
        self.log_box.delete('1.0', 'end')
        t = threading.Thread(target=self.run_generate, args=(csv_path, out_dir), daemon=True)
        t.start()

    def run_generate(self, csv_path, out_dir):
        try:
            field = FIELD_MAP[self.field_var.get()]
            value = self.value_var.get().strip()

            self.log(f"Reading {csv_path} ...")
            rows = load_rows(csv_path)
            self.log(f"  {len(rows)} rows loaded")

            self.log("Resolving fields (status tracking, dispatch parsing) ...")
            records = build_records(rows)

            groups = defaultdict(list)
            for r in records:
                groups[r[field]].append(r)
            self.log(f"  {len(groups)} unique values found for this field")

            targets = list(groups.keys())
            if value:
                targets = [g for g in targets if value.lower() in g.lower()]
                if not targets:
                    self.log("No matching records for that filter value.")
                    self.finish()
                    return

            self.log(f"Generating {len(targets)} PDF(s) ...")
            written = 0
            for key in sorted(targets):
                group_records = groups[key]
                cust_id = group_records[0]['cust_id']
                id_prefix = extract_id_number(cust_id)
                fname = f"{id_prefix} - {sanitize_filename(key)}.pdf"
                out_path = os.path.join(out_dir, fname)
                make_pdf(group_records, out_path, key)
                written += 1
                if written % 25 == 0:
                    self.log(f"  {written}/{len(targets)} done")

            self.log(f"\nDone! {written} PDF(s) saved to:\n{out_dir}")
            self.finish(open_folder=out_dir)
        except Exception as e:
            self.log(f"ERROR: {e}")
            self.finish()

    def finish(self, open_folder=None):
        self.progress.stop()
        if open_folder:
            try:
                if sys.platform == 'win32':
                    os.startfile(open_folder)
                elif sys.platform == 'darwin':
                    subprocess.run(['open', open_folder])
                else:
                    subprocess.run(['xdg-open', open_folder])
            except Exception:
                pass
            messagebox.showinfo("Done", "PDF generation complete! The output folder has been opened.")


if __name__ == '__main__':
    App().mainloop()
