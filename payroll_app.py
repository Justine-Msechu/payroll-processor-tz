import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, sys, subprocess, hashlib, json, datetime, tempfile, threading, webbrowser

# ══════════════════════════════════════════════
#  APP IDENTITY & GITHUB UPDATE SETTINGS
# ══════════════════════════════════════════════
VERSION       = "1.0.0"
APP_TITLE     = "Payroll Processor — Tanzania"
USERS_FILE    = "payroll_users.json"

# ── To enable auto-update, set these to your GitHub details ──
# 1. Create a GitHub repo (e.g. "payroll-processor-tz")
# 2. Upload payroll_app.py and a file called version.txt (containing just: 1.0.1)
# 3. Change the two lines below to your username and repo name
GITHUB_USER   = "Justine-Msechu"  # Specifical when i want to Update the app while its on use i use this gitname
GITHUB_REPO   = "payroll-processor-tz"    
GITHUB_BRANCH = "main"
_base         = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}"
VERSION_URL   = f"{_base}/version.txt"
APP_URL       = f"{_base}/payroll_app.py"
UPDATE_READY  = GITHUB_USER != "YOUR_GITHUB_USERNAME"

HEADERS = [
    "Name", "Salary", "NSSF 10%", "PAY", "P.A.Y.E",
    "ALLOWANCE", "GROSS PAY", "LOAN Deduction",
    "LOAN BOARD (15%)", "NET PAY", "AMOUNT TO BE PAID"
]

# ══════════════════════════════════════════════
#  THEMES
# ══════════════════════════════════════════════
THEMES = {
    "Dark":dict(BG="#000000", PANEL="#263547", ACCENT="#2ecc71", ACCENT2="#27ae60",
        TEXT="#ecf0f1", SUBTEXT="#95a5a6", ENTRY_BG="#2d4059", ENTRY_FG="#ecf0f1",
        RED="#e74c3c", GOLD="#f39c12", HEADER_BG="#27ae60",
        BTN_SAVE="#2980b9", BTN_PRINT="#e67e22", BTN_OPEN="#8e44ad",
        TREE_BG="#2d4059", TREE_FG="#ecf0f1", TREE_HEAD="#1A5276", SEP="#34495e",),
    "Dark Blue": dict(
        BG="#1e2a3a", PANEL="#263547", ACCENT="#2ecc71", ACCENT2="#27ae60",
        TEXT="#ecf0f1", SUBTEXT="#95a5a6", ENTRY_BG="#2d4059", ENTRY_FG="#ecf0f1",
        RED="#e74c3c", GOLD="#f39c12", HEADER_BG="#27ae60",
        BTN_SAVE="#2980b9", BTN_PRINT="#e67e22", BTN_OPEN="#8e44ad",
        TREE_BG="#2d4059", TREE_FG="#ecf0f1", TREE_HEAD="#1A5276", SEP="#34495e",
    ),
    "Light": dict(
        BG="#f0f4f8", PANEL="#ffffff", ACCENT="#2563eb", ACCENT2="#1d4ed8",
        TEXT="#1e293b", SUBTEXT="#64748b", ENTRY_BG="#e2e8f0", ENTRY_FG="#1e293b",
        RED="#dc2626", GOLD="#d97706", HEADER_BG="#2563eb",
        BTN_SAVE="#0369a1", BTN_PRINT="#b45309", BTN_OPEN="#7c3aed",
        TREE_BG="#f8fafc", TREE_FG="#1e293b", TREE_HEAD="#1e40af", SEP="#cbd5e1",
    ),
    "Green": dict(
        BG="#0f2418", PANEL="#1a3a2a", ACCENT="#4ade80", ACCENT2="#16a34a",
        TEXT="#dcfce7", SUBTEXT="#86efac", ENTRY_BG="#14532d", ENTRY_FG="#dcfce7",
        RED="#f87171", GOLD="#fbbf24", HEADER_BG="#16a34a",
        BTN_SAVE="#0e7490", BTN_PRINT="#b45309", BTN_OPEN="#7c3aed",
        TREE_BG="#14532d", TREE_FG="#dcfce7", TREE_HEAD="#166534", SEP="#166534",
    ),
    "Purple": dict(
        BG="#1e1b2e", PANEL="#2d2b45", ACCENT="#a78bfa", ACCENT2="#7c3aed",
        TEXT="#ede9fe", SUBTEXT="#c4b5fd", ENTRY_BG="#3b3563", ENTRY_FG="#ede9fe",
        RED="#f87171", GOLD="#fbbf24", HEADER_BG="#7c3aed",
        BTN_SAVE="#2563eb", BTN_PRINT="#d97706", BTN_OPEN="#16a34a",
        TREE_BG="#3b3563", TREE_FG="#ede9fe", TREE_HEAD="#4c1d95", SEP="#4c1d95",
    ),
    "Orange": dict(
        BG="#1c1007", PANEL="#2d1f0e", ACCENT="#fb923c", ACCENT2="#ea580c",
        TEXT="#fff7ed", SUBTEXT="#fed7aa", ENTRY_BG="#431407", ENTRY_FG="#fff7ed",
        RED="#f87171", GOLD="#fbbf24", HEADER_BG="#ea580c",
        BTN_SAVE="#0369a1", BTN_PRINT="#7c3aed", BTN_OPEN="#16a34a",
        TREE_BG="#431407", TREE_FG="#fff7ed", TREE_HEAD="#7c2d12", SEP="#7c2d12",
    ),
}

# ══════════════════════════════════════════════
#  CALCULATIONS
# ══════════════════════════════════════════════
def calc_paye_tz(pay):
    if pay <= 270_000:   return 0.0
    if pay <= 520_000:   return round((pay - 270_000) * 0.08, 2)
    if pay <= 760_000:   return round(20_000 + (pay - 520_000) * 0.20, 2)
    if pay <= 1_000_000: return round(68_000 + (pay - 760_000) * 0.25, 2)
    return round(128_000 + (pay - 1_000_000) * 0.30, 2)

def calculate(salary, allowance, has_loan, loan_amount, has_loan_board):
    nssf        = round(salary * 0.10, 2)
    pay         = round(salary - nssf, 2)
    paye        = calc_paye_tz(pay)
    gross_pay   = round(salary + allowance, 2)
    loan_ded    = round(loan_amount, 2) if has_loan else 0.0
    loan_board  = round(salary * 0.15, 2) if has_loan_board else 0.0
    net_pay     = round(gross_pay - nssf - paye, 2)
    amount_paid = round(net_pay - loan_ded - loan_board, 2)
    return nssf, pay, paye, gross_pay, loan_ded, loan_board, net_pay, amount_paid

# ══════════════════════════════════════════════
#  AUTH
# ══════════════════════════════════════════════
def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def load_users():
    if not os.path.exists(USERS_FILE):
        d = {"admin": {"hash": hash_pw("admin123"), "role": "admin"}}
        save_users(d)
        return d
    with open(USERS_FILE) as f:
        return json.load(f)

def save_users(u):
    with open(USERS_FILE, "w") as f:
        json.dump(u, f, indent=2)

def verify_login(username, password):
    u = load_users().get(username.strip().lower())
    return u["role"] if u and u["hash"] == hash_pw(password) else None

# ══════════════════════════════════════════════
#  EXCEL EXPORT
# ══════════════════════════════════════════════
def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def save_to_excel(records, filepath, created_by, month_label):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Payroll"
    now_str    = datetime.datetime.now().strftime("%d %B %Y  %H:%M")
    info_fill  = PatternFill("solid", fgColor="D6EAF8")

    # Info header rows 1-4
    ws.merge_cells("A1:K1")
    c = ws["A1"]
    c.value     = "PAYROLL REGISTER  —  " + month_label
    c.font      = Font(bold=True, size=13, color="1A5276")
    c.fill      = info_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for row_i, (label, value) in enumerate([
        ("Created by",           created_by),
        ("Date & Time",          now_str),
        ("Number of Employees",  str(len(records))),
    ], start=2):
        ws.cell(row=row_i, column=1, value=label).font = Font(bold=True, size=9, color="1A5276")
        ws.cell(row=row_i, column=2, value=value).font = Font(size=9, color="2C3E50")
        ws.cell(row=row_i, column=1).fill = info_fill
        ws.cell(row=row_i, column=2).fill = info_fill

    ws.row_dimensions[5].height = 6  # spacer

    # Column headers — row 6
    header_row = 6
    hfill = PatternFill("solid", fgColor="1A5276")
    for col, h in enumerate(HEADERS, 1):
        c = ws.cell(row=header_row, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10)
        c.fill      = hfill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = thin_border()
    ws.row_dimensions[header_row].height = 32

    # Data rows
    next_row = header_row + 1
    for r in records:
        row_data = [r["name"], r["salary"], r["nssf"], r["pay"], r["paye"],
                    r["allowance"], r["gross_pay"], r["loan_ded"],
                    r["loan_board"], r["net_pay"], r["amount_paid"]]
        rfill = PatternFill("solid", fgColor="EBF5FB" if next_row % 2 == 0 else "FDFEFE")
        for col, val in enumerate(row_data, 1):
            c = ws.cell(row=next_row, column=col, value=val)
            c.fill      = rfill
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = thin_border()
            if col > 1:
                c.number_format = "#,##0.00"
        next_row += 1

    # Spacer
    ws.row_dimensions[next_row].height = 10
    next_row += 1

    # Summary box
    total_salary = sum(r["salary"]      for r in records)
    total_nssf   = sum(r["nssf"]        for r in records)
    total_paye   = sum(r["paye"]        for r in records)
    total_gross  = sum(r["gross_pay"]   for r in records)
    total_loan   = sum(r["loan_ded"]    for r in records)
    total_lb     = sum(r["loan_board"]  for r in records)
    total_net    = sum(r["net_pay"]     for r in records)
    total_amount = sum(r["amount_paid"] for r in records)

    sum_title_fill = PatternFill("solid", fgColor="1A5276")
    sum_key_fill   = PatternFill("solid", fgColor="D6EAF8")
    highlight_fill = PatternFill("solid", fgColor="145A32")

    ws.merge_cells(f"A{next_row}:K{next_row}")
    sc = ws.cell(row=next_row, column=1, value="PAYROLL SUMMARY")
    sc.font      = Font(bold=True, color="FFFFFF", size=11)
    sc.fill      = sum_title_fill
    sc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[next_row].height = 22
    next_row += 1

    for label, value, highlight in [
        ("Total Number of Employees",       len(records),   False),
        ("Total Gross Salary",               total_salary,   False),
        ("Total NSSF Deductions",            total_nssf,     False),
        ("Total PAYE Tax",                   total_paye,     False),
        ("Total Gross Pay (Salary+Allow.)",  total_gross,    False),
        ("Total Loan Deductions",            total_loan,     False),
        ("Total Loan Board Deductions",      total_lb,       False),
        ("Total Net Pay",                    total_net,      False),
        ("TOTAL AMOUNT TO BE PAID OUT",      total_amount,   True),
    ]:
        lc = ws.cell(row=next_row, column=1, value=label)
        vc = ws.cell(row=next_row, column=2,
                     value=value if isinstance(value, int) else round(value, 2))
        fill = highlight_fill if highlight else sum_key_fill
        for cell in (lc, vc):
            cell.fill   = fill
            cell.border = thin_border()
        if highlight:
            lc.font = Font(bold=True, color="FFFFFF", size=11)
            vc.font = Font(bold=True, color="FFFFFF", size=11)
            ws.row_dimensions[next_row].height = 24
        else:
            lc.font = Font(bold=True, size=9, color="1A5276")
            vc.font = Font(size=9, color="1A5276")
        lc.alignment = Alignment(horizontal="left",  vertical="center", indent=1)
        vc.alignment = Alignment(horizontal="right", vertical="center")
        if not isinstance(value, int):
            vc.number_format = "#,##0.00"
        next_row += 1

    # Auto-fit columns
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(w + 4, 18)

    ws.freeze_panes = f"A{header_row + 1}"
    wb.save(filepath)

# ══════════════════════════════════════════════
#  PRINT — HTML opened in browser
# ══════════════════════════════════════════════
def make_print_html(records, month_label, created_by):
    now_str = datetime.datetime.now().strftime("%d %B %Y  %H:%M")
    rows_html = ""
    for i, r in enumerate(records):
        bg = "#eaf4fb" if i % 2 == 0 else "#ffffff"
        rows_html += (
            "<tr style=\"background:" + bg + "\">"
            "<td style=\"text-align:left;font-weight:500\">" + r["name"] + "</td>"
            "<td>" + f"{r['salary']:,.2f}" + "</td>"
            "<td>" + f"{r['nssf']:,.2f}" + "</td>"
            "<td>" + f"{r['pay']:,.2f}" + "</td>"
            "<td>" + f"{r['paye']:,.2f}" + "</td>"
            "<td>" + f"{r['allowance']:,.2f}" + "</td>"
            "<td>" + f"{r['gross_pay']:,.2f}" + "</td>"
            "<td>" + f"{r['loan_ded']:,.2f}" + "</td>"
            "<td>" + f"{r['loan_board']:,.2f}" + "</td>"
            "<td>" + f"{r['net_pay']:,.2f}" + "</td>"
            "<td style=\"color:#145a32;font-weight:bold\">" + f"{r['amount_paid']:,.2f}" + "</td>"
            "</tr>"
        )

    total_salary = sum(r["salary"]      for r in records)
    total_nssf   = sum(r["nssf"]        for r in records)
    total_paye   = sum(r["paye"]        for r in records)
    total_net    = sum(r["net_pay"]     for r in records)
    total_amount = sum(r["amount_paid"] for r in records)

    header_cells = "".join("<th>" + h + "</th>" for h in HEADERS)

    html = (
        "<!DOCTYPE html><html><head><meta charset=\"UTF-8\">"
        "<title>Payroll " + month_label + "</title>"
        "<style>"
        "@page{size:A4 landscape;margin:1.2cm}"
        "*{box-sizing:border-box;font-family:'Segoe UI',Arial,sans-serif;margin:0;padding:0}"
        "body{padding:14px;color:#1e293b}"
        ".top{border-bottom:3px solid #1A5276;padding-bottom:8px;margin-bottom:10px;"
        "display:flex;justify-content:space-between;align-items:flex-end}"
        ".top h1{font-size:18px;color:#1A5276;letter-spacing:.5px}"
        ".meta{font-size:10px;color:#64748b;text-align:right;line-height:1.7}"
        "table{width:100%;border-collapse:collapse;font-size:9px}"
        "th{background:#1A5276;color:white;padding:7px 4px;text-align:center;"
        "font-weight:bold;border:1px solid #b0c4d8}"
        "td{padding:5px 4px;text-align:right;border:1px solid #dde3ea;white-space:nowrap}"
        ".summary{margin-top:16px;display:flex;gap:10px;flex-wrap:wrap}"
        ".scard{background:#f0f9ff;border:2px solid #1A5276;border-radius:6px;"
        "padding:10px 16px;min-width:150px;text-align:center}"
        ".scard.hl{background:#145a32;border-color:#145a32;color:white}"
        ".scard .lbl{font-size:9px;color:#64748b;margin-bottom:4px}"
        ".scard.hl .lbl{color:#a7f3d0}"
        ".scard .val{font-size:13px;font-weight:bold;color:#1A5276}"
        ".scard.hl .val{color:white;font-size:15px}"
        ".sigs{margin-top:20px;display:flex;justify-content:space-around}"
        ".sig{text-align:center;width:180px}"
        ".sig-line{border-top:1px solid #555;margin-bottom:3px}"
        ".footer{margin-top:10px;font-size:8px;color:#94a3b8;text-align:center}"
        "@media print{body{padding:0}}"
        "</style></head><body>"
        "<div class=\"top\">"
        "<div><h1>PAYROLL REGISTER" + ("  \u2014  " + month_label if month_label else "") + "</h1></div>"
        "<div class=\"meta\">"
        "<div><b>Created by:</b> " + created_by + "</div>"
        "<div><b>Date &amp; Time:</b> " + now_str + "</div>"
        "<div><b>Employees:</b> " + str(len(records)) + "</div>"
        "</div></div>"
        "<table><thead><tr>" + header_cells + "</tr></thead>"
        "<tbody>" + rows_html + "</tbody></table>"
        "<div class=\"summary\">"
        "<div class=\"scard\"><div class=\"lbl\">Employees</div>"
        "<div class=\"val\">" + str(len(records)) + "</div></div>"
        "<div class=\"scard\"><div class=\"lbl\">Total Gross Salary</div>"
        "<div class=\"val\">TZS " + f"{total_salary:,.2f}" + "</div></div>"
        "<div class=\"scard\"><div class=\"lbl\">Total NSSF</div>"
        "<div class=\"val\">TZS " + f"{total_nssf:,.2f}" + "</div></div>"
        "<div class=\"scard\"><div class=\"lbl\">Total PAYE Tax</div>"
        "<div class=\"val\">TZS " + f"{total_paye:,.2f}" + "</div></div>"
        "<div class=\"scard\"><div class=\"lbl\">Total Net Pay</div>"
        "<div class=\"val\">TZS " + f"{total_net:,.2f}" + "</div></div>"
        "<div class=\"scard hl\"><div class=\"lbl\">TOTAL AMOUNT TO BE PAID OUT</div>"
        "<div class=\"val\">TZS " + f"{total_amount:,.2f}" + "</div></div>"
        "</div>"
        "<div class=\"sigs\">"
        "<div class=\"sig\"><div class=\"sig-line\"></div>Prepared by</div>"
        "<div class=\"sig\"><div class=\"sig-line\"></div>Reviewed by</div>"
        "<div class=\"sig\"><div class=\"sig-line\"></div>Approved by</div>"
        "</div>"
        "<div class=\"footer\">All amounts in TZS &nbsp;|&nbsp; "
        "NSSF = 10% of Salary &nbsp;|&nbsp; PAYE = TRA 2024/25 &nbsp;|&nbsp; "
        "Loan Board = 15% of Salary</div>"
        "</body></html>"
    )
    return html

def open_print_in_browser(html_content):
    """Write HTML to a temp file and open it in the default browser."""
    tmp = tempfile.NamedTemporaryFile(
        delete=False, suffix=".html", mode="w", encoding="utf-8")
    tmp.write(html_content)
    tmp.close()
    url = "file:///" + tmp.name.replace("\\", "/")
    webbrowser.open(url)

# ══════════════════════════════════════════════
#  AUTO-UPDATE FROM GITHUB
# ══════════════════════════════════════════════
def _ver(v):
    try:
        return tuple(int(x) for x in v.strip().split("."))
    except Exception:
        return (0, 0, 0)

def check_for_update_bg(callback):
    """Background thread — calls callback(latest_version_str) if newer."""
    def _run():
        try:
            import urllib.request
            with urllib.request.urlopen(VERSION_URL, timeout=6) as r:
                latest = r.read().decode().strip()
            if _ver(latest) > _ver(VERSION):
                callback(latest)
        except Exception:
            pass
    threading.Thread(target=_run, daemon=True).start()

def check_for_update_bg_with_result(result_holder):
    """Background thread — puts result into result_holder[0]."""
    def _run():
        try:
            import urllib.request
            with urllib.request.urlopen(VERSION_URL, timeout=6) as r:
                latest = r.read().decode().strip()
            result_holder[0] = latest if _ver(latest) > _ver(VERSION) else "up_to_date"
        except Exception:
            result_holder[0] = "error"
    threading.Thread(target=_run, daemon=True).start()

def download_and_install(latest_version, parent):
    """Download new payroll_app.py from GitHub and restart."""
    import urllib.request, shutil

    # Progress dialog
    dlg = tk.Toplevel(parent)
    dlg.title("Downloading update...")
    dlg.configure(bg="#1e2a3a")
    dlg.resizable(False, False)
    dlg.grab_set()
    sw = parent.winfo_screenwidth()
    sh = parent.winfo_screenheight()
    dlg.geometry("380x150+" + str((sw - 380) // 2) + "+" + str((sh - 150) // 2))
    tk.Label(dlg, text="Downloading version " + latest_version + "...",
             font=("Segoe UI", 12, "bold"), bg="#1e2a3a", fg="#2ecc71").pack(pady=(24, 6))
    tk.Label(dlg, text="Please wait, do not close the app",
             font=("Segoe UI", 9), bg="#1e2a3a", fg="#95a5a6").pack()
    pb = ttk.Progressbar(dlg, mode="indeterminate", length=320)
    pb.pack(pady=14)
    pb.start(10)
    dlg.update()

    try:
        this_file = os.path.abspath(__file__)
        backup    = this_file + ".bak"
        tmp_path  = this_file + ".new"
        shutil.copy2(this_file, backup)
        urllib.request.urlretrieve(APP_URL, tmp_path)

        # Sanity check
        with open(tmp_path, "r", encoding="utf-8") as f:
            content = f.read()
        if "PayrollApp" not in content or "calc_paye_tz" not in content:
            os.remove(tmp_path)
            dlg.destroy()
            messagebox.showerror("Update Failed",
                "The downloaded file looks incorrect.\nPlease try again later.",
                parent=parent)
            return

        dlg.destroy()
        shutil.move(tmp_path, this_file)
        messagebox.showinfo("Updated!",
            "The app has been updated to version " + latest_version + ".\n\n"
            "The app will now restart.",
            parent=parent)
        os.execv(sys.executable, [sys.executable, this_file])

    except Exception as e:
        try:
            dlg.destroy()
        except Exception:
            pass
        messagebox.showerror("Update Failed",
            "Could not download the update.\n\n"
            "Error: " + str(e) + "\n\n"
            "Check your internet connection and try again.",
            parent=parent)

# ══════════════════════════════════════════════
#  SCROLLABLE FRAME
# ══════════════════════════════════════════════
class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win  = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            self._win, width=e.width))
        for w in (self.canvas, self.inner):
            w.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(
                int(-1 * (e.delta / 120)), "units"))

    def set_bg(self, bg):
        self.canvas.configure(bg=bg)
        self.inner.configure(bg=bg)

# ══════════════════════════════════════════════
#  TOOLTIP
# ══════════════════════════════════════════════
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text   = text
        self.tip    = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _=None):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry("+" + str(x) + "+" + str(y))
        tk.Label(self.tip, text=self.text, font=("Segoe UI", 8),
                 bg="#fffde7", fg="#333", relief="solid", bd=1,
                 padx=6, pady=3).pack()

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None

# ══════════════════════════════════════════════
#  LOGIN WINDOW
# ══════════════════════════════════════════════
class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(False, False)
        self.configure(bg="#1e2a3a")
        self.logged_in_user = None
        self.logged_in_role = None
        self._build()
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(str(w) + "x" + str(h) + "+" + str((sw - w) // 2) + "+" + str((sh - h) // 2))

    def _build(self):
        T = THEMES["Dark Blue"]
        top = tk.Frame(self, bg=T["HEADER_BG"], pady=22)
        top.pack(fill="x")
        tk.Label(top, text="💼", font=("Segoe UI", 40),
                 bg=T["HEADER_BG"], fg="white").pack()
        tk.Label(top, text="PAYROLL PROCESSOR", font=("Segoe UI", 17, "bold"),
                 bg=T["HEADER_BG"], fg="white").pack()
        tk.Label(top, text="Tanzania  ·  Please log in to continue",
                 font=("Segoe UI", 9), bg=T["HEADER_BG"], fg="#d5f5e3").pack(pady=(2, 0))

        card = tk.Frame(self, bg=T["PANEL"], padx=44, pady=30)
        card.pack(padx=50, pady=28)
        self._u_var = tk.StringVar()
        self._p_var = tk.StringVar()
        for row, (icon, lbl, var, show) in enumerate([
            ("👤", "Username", self._u_var, ""),
            ("🔑", "Password", self._p_var, "●"),
        ], start=1):
            tk.Label(card, text=icon + "  " + lbl, font=("Segoe UI", 10),
                     bg=T["PANEL"], fg=T["TEXT"], anchor="w"
                     ).grid(row=row * 2 - 1, column=0, sticky="w", pady=(8, 1))
            e = tk.Entry(card, textvariable=var, show=show, font=("Segoe UI", 12),
                         bg=T["ENTRY_BG"], fg=T["ENTRY_FG"],
                         insertbackground="white", relief="flat", bd=6, width=24)
            e.grid(row=row * 2, column=0, ipady=6, sticky="ew")
            if row == 1:
                e.focus_set()

        self.err_var = tk.StringVar()
        tk.Label(card, textvariable=self.err_var, font=("Segoe UI", 9),
                 bg=T["PANEL"], fg=T["RED"], wraplength=280
                 ).grid(row=5, column=0, pady=(6, 0))
        tk.Button(card, text="🔓   LOG IN", font=("Segoe UI", 12, "bold"),
                  bg=T["ACCENT2"], fg="white", relief="flat", cursor="hand2",
                  command=self._login, activebackground=T["ACCENT"]
                  ).grid(row=6, column=0, pady=(16, 0), ipadx=12, ipady=10, sticky="ew")
        tk.Label(card, text="First-time login:  username = admin  |  password = admin123",
                 font=("Segoe UI", 7), bg=T["PANEL"], fg=T["SUBTEXT"]
                 ).grid(row=7, column=0, pady=(12, 0))
        self.bind("<Return>", lambda e: self._login())

    def _login(self):
        u = self._u_var.get().strip()
        p = self._p_var.get()
        if not u or not p:
            self.err_var.set("⚠  Please fill in both fields.")
            return
        role = verify_login(u, p)
        if role:
            self.logged_in_user = u.lower()
            self.logged_in_role = role
            self.destroy()
        else:
            self.err_var.set("❌  Wrong username or password. Try again.")
            self._p_var.set("")

# ══════════════════════════════════════════════
#  USER MANAGER DIALOG
# ══════════════════════════════════════════════
class UserManagerDialog(tk.Toplevel):
    def __init__(self, parent, T):
        super().__init__(parent)
        self.T = T
        self.title("Manage Users")
        self.configure(bg=T["BG"])
        self.resizable(False, False)
        self._build()
        self.grab_set()
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        px = parent.winfo_x() + parent.winfo_width() // 2
        py = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(str(w) + "x" + str(h) + "+" + str(px - w // 2) + "+" + str(py - h // 2))

    def _build(self):
        T = self.T
        tk.Label(self, text="👥  Manage Users", font=("Segoe UI", 13, "bold"),
                 bg=T["BG"], fg=T["ACCENT"]).pack(pady=(16, 4))
        tk.Frame(self, bg=T["ACCENT"], height=2).pack(fill="x", padx=20)

        lf = tk.Frame(self, bg=T["PANEL"], padx=14, pady=10)
        lf.pack(fill="both", expand=True, padx=20, pady=10)
        tk.Label(lf, text="Existing Users:", font=("Segoe UI", 9, "bold"),
                 bg=T["PANEL"], fg=T["TEXT"]).pack(anchor="w")
        self.lb = tk.Listbox(lf, font=("Segoe UI", 9), bg=T["ENTRY_BG"],
                              fg=T["ENTRY_FG"], relief="flat", height=5,
                              selectbackground=T["ACCENT2"])
        self.lb.pack(fill="both", expand=True, pady=4)
        tk.Button(lf, text="🗑  Remove Selected User", font=("Segoe UI", 8),
                  bg=T["RED"], fg="white", relief="flat", cursor="hand2",
                  command=self._remove).pack(anchor="e", ipady=3, ipadx=8, pady=(2, 0))
        self._refresh()

        af = tk.Frame(self, bg=T["PANEL"], padx=14, pady=10)
        af.pack(fill="x", padx=20, pady=(0, 10))
        tk.Label(af, text="Add a New User:", font=("Segoe UI", 9, "bold"),
                 bg=T["PANEL"], fg=T["TEXT"]).pack(anchor="w", pady=(0, 6))
        self._nu = tk.StringVar()
        self._np = tk.StringVar()
        self._nr = tk.StringVar(value="accountant")
        for lbl, var, show in [("Username", self._nu, ""), ("Password", self._np, "●")]:
            row = tk.Frame(af, bg=T["PANEL"])
            row.pack(fill="x", pady=2)
            tk.Label(row, text=lbl, width=12, anchor="w", font=("Segoe UI", 9),
                     bg=T["PANEL"], fg=T["TEXT"]).pack(side="left")
            tk.Entry(row, textvariable=var, show=show, font=("Segoe UI", 9),
                     bg=T["ENTRY_BG"], fg=T["ENTRY_FG"], insertbackground="white",
                     relief="flat", bd=3).pack(side="right", expand=True, fill="x", ipady=4)
        rr = tk.Frame(af, bg=T["PANEL"])
        rr.pack(fill="x", pady=2)
        tk.Label(rr, text="Role", width=12, anchor="w", font=("Segoe UI", 9),
                 bg=T["PANEL"], fg=T["TEXT"]).pack(side="left")
        om = tk.OptionMenu(rr, self._nr, "accountant", "admin")
        om.configure(bg=T["ENTRY_BG"], fg=T["TEXT"], relief="flat",
                     font=("Segoe UI", 9), highlightthickness=0)
        om["menu"].configure(bg=T["ENTRY_BG"], fg=T["TEXT"])
        om.pack(side="left", padx=2)
        tk.Button(af, text="➕  Add User", font=("Segoe UI", 10, "bold"),
                  bg=T["ACCENT2"], fg="white", relief="flat", cursor="hand2",
                  command=self._add).pack(fill="x", pady=(8, 0), ipady=6)
        self.msg = tk.StringVar()
        tk.Label(self, textvariable=self.msg, font=("Segoe UI", 8),
                 bg=T["BG"], fg=T["GOLD"]).pack(pady=(0, 10))

    def _refresh(self):
        self.lb.delete(0, "end")
        for u, d in load_users().items():
            self.lb.insert("end", "  " + u + "   (" + d["role"] + ")")

    def _add(self):
        u = self._nu.get().strip().lower()
        p = self._np.get()
        if not u or not p:
            self.msg.set("⚠  Please enter both username and password.")
            return
        users = load_users()
        if u in users:
            self.msg.set("⚠  User '" + u + "' already exists.")
            return
        users[u] = {"hash": hash_pw(p), "role": self._nr.get()}
        save_users(users)
        self.msg.set("✅  User '" + u + "' added.")
        self._nu.set("")
        self._np.set("")
        self._refresh()

    def _remove(self):
        sel = self.lb.curselection()
        if not sel:
            return
        uname = self.lb.get(sel[0]).strip().split()[0]
        if uname == "admin" and len(load_users()) == 1:
            messagebox.showwarning("Cannot Remove",
                "Cannot remove the only admin account.", parent=self)
            return
        if messagebox.askyesno("Confirm", "Remove user '" + uname + "'?", parent=self):
            users = load_users()
            users.pop(uname, None)
            save_users(users)
            self.msg.set("User '" + uname + "' removed.")
            self._refresh()

# ══════════════════════════════════════════════
#  CHANGE PASSWORD DIALOG
# ══════════════════════════════════════════════
class ChangePasswordDialog(tk.Toplevel):
    def __init__(self, parent, T, current_user):
        super().__init__(parent)
        self.T = T
        self.current_user = current_user
        self.title("Change My Password")
        self.configure(bg=T["BG"])
        self.resizable(False, False)
        self._build()
        self.grab_set()
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        px = parent.winfo_x() + parent.winfo_width() // 2
        py = parent.winfo_y() + parent.winfo_height() // 2
        self.geometry(str(w) + "x" + str(h) + "+" + str(px - w // 2) + "+" + str(py - h // 2))

    def _build(self):
        T = self.T
        card = tk.Frame(self, bg=T["PANEL"], padx=32, pady=26)
        card.pack(padx=26, pady=26)
        tk.Label(card, text="🔒  Change My Password", font=("Segoe UI", 12, "bold"),
                 bg=T["PANEL"], fg=T["ACCENT"]
                 ).grid(row=0, column=0, columnspan=2, pady=(0, 16))
        self._vars = {}
        for r, (lbl, key) in enumerate([
            ("Current Password", "old"),
            ("New Password",     "new1"),
            ("Confirm New",      "new2"),
        ], start=1):
            tk.Label(card, text=lbl, font=("Segoe UI", 9), bg=T["PANEL"],
                     fg=T["TEXT"], anchor="w", width=18
                     ).grid(row=r, column=0, sticky="w", pady=5)
            var = tk.StringVar()
            tk.Entry(card, textvariable=var, show="●", font=("Segoe UI", 10),
                     bg=T["ENTRY_BG"], fg=T["ENTRY_FG"],
                     insertbackground="white", relief="flat", bd=4, width=22
                     ).grid(row=r, column=1, padx=(8, 0), ipady=6)
            self._vars[key] = var
        self.msg = tk.StringVar()
        tk.Label(card, textvariable=self.msg, font=("Segoe UI", 8),
                 bg=T["PANEL"], fg=T["RED"], wraplength=260
                 ).grid(row=4, column=0, columnspan=2, pady=(4, 0))
        tk.Button(card, text="✅  Save New Password", font=("Segoe UI", 10, "bold"),
                  bg=T["ACCENT2"], fg="white", relief="flat", cursor="hand2",
                  command=self._change
                  ).grid(row=5, column=0, columnspan=2,
                         pady=(14, 0), ipadx=10, ipady=8, sticky="ew")

    def _change(self):
        old = self._vars["old"].get()
        n1  = self._vars["new1"].get()
        n2  = self._vars["new2"].get()
        users = load_users()
        u = users.get(self.current_user)
        if not u or u["hash"] != hash_pw(old):
            self.msg.set("⚠  Current password is incorrect.")
            return
        if len(n1) < 6:
            self.msg.set("⚠  New password must be at least 6 characters.")
            return
        if n1 != n2:
            self.msg.set("⚠  Passwords do not match.")
            return
        users[self.current_user]["hash"] = hash_pw(n1)
        save_users(users)
        messagebox.showinfo("Done ✅", "Your password has been updated!", parent=self)
        self.destroy()

# ══════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════
class PayrollApp(tk.Tk):
    def __init__(self, username, role):
        super().__init__()
        self.username         = username
        self.role             = role
        self.records          = []
        self._last_saved_path = None
        self._update_result   = None
        self.theme_name       = tk.StringVar(value="Dark Blue")
        self.T                = THEMES["Dark Blue"]
        self.title(APP_TITLE + "  —  " + username)
        try:
            self.state("zoomed")
        except Exception:
            self.attributes("-zoomed", True)
        self.minsize(820, 560)
        self._build_ui()
        self._apply_theme()
        # Auto-check for updates 3 s after startup (silent, background)
        if UPDATE_READY:
            self.after(3000, self._auto_update_check)

    # ── Theme ─────────────────────────────────────────────────
    def _apply_theme(self, *_):
        self.T = THEMES[self.theme_name.get()]
        T = self.T
        self.configure(bg=T["BG"])
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("P.Treeview", background=T["TREE_BG"], foreground=T["TREE_FG"],
                    fieldbackground=T["TREE_BG"], rowheight=27, font=("Segoe UI", 9))
        s.configure("P.Treeview.Heading", background=T["TREE_HEAD"],
                    foreground="white", font=("Segoe UI", 8, "bold"), relief="flat")
        s.map("P.Treeview", background=[("selected", T["ACCENT2"])])
        if hasattr(self, "tree"):
            self.tree.configure(style="P.Treeview")
        self._restyle(self)
        for attr, ck in [
            ("btn_add",    "BTN_SAVE"),
            ("btn_save",   "BTN_SAVE"),
            ("btn_print",  "BTN_PRINT"),
            ("btn_open",   "BTN_OPEN"),
            ("btn_reset",  "SUBTEXT"),
            ("btn_remove", "RED"),
        ]:
            if hasattr(self, attr):
                getattr(self, attr).configure(bg=T[ck])
        if hasattr(self, "_sf"):
            self._sf.set_bg(T["PANEL"])
        self._refresh_totals_bar()

    def _restyle(self, root):
        T = self.T
        ap  = {THEMES[n]["PANEL"]     for n in THEMES}
        ah  = {THEMES[n]["HEADER_BG"] for n in THEMES}
        aa  = {THEMES[n]["ACCENT"]    for n in THEMES}
        as_ = {THEMES[n]["SUBTEXT"]   for n in THEMES}
        ag  = {THEMES[n]["GOLD"]      for n in THEMES}
        ar  = {THEMES[n]["RED"]       for n in THEMES}
        aa2 = {THEMES[n]["ACCENT2"]   for n in THEMES}

        def walk(w):
            cls = type(w).__name__
            try:
                if cls in ("Frame", "Canvas"):
                    bg = w.cget("bg")
                    w.configure(bg=T["HEADER_BG"] if bg in ah
                                else T["PANEL"] if bg in ap
                                else T["BG"])
                elif cls == "Label":
                    bg = w.cget("bg")
                    fg = w.cget("fg")
                    nbg = (T["HEADER_BG"] if bg in ah
                           else T["PANEL"] if bg in ap
                           else T["BG"])
                    if   fg in ("white", "#ffffff"): nfg = "white"
                    elif fg in aa:   nfg = T["ACCENT"]
                    elif fg in as_:  nfg = T["SUBTEXT"]
                    elif fg in ag:   nfg = T["GOLD"]
                    elif fg in ar:   nfg = T["RED"]
                    elif fg in aa2:  nfg = T["ACCENT2"]
                    else:            nfg = T["TEXT"]
                    w.configure(bg=nbg, fg=nfg)
                elif cls == "Entry":
                    w.configure(bg=T["ENTRY_BG"], fg=T["ENTRY_FG"],
                                insertbackground=T["TEXT"])
                elif cls == "Checkbutton":
                    bg = w.cget("bg")
                    nbg = T["PANEL"] if bg in ap else T["BG"]
                    w.configure(bg=nbg, fg=T["TEXT"],
                                selectcolor=T["ENTRY_BG"],
                                activebackground=nbg,
                                activeforeground=T["TEXT"])
                elif cls == "OptionMenu":
                    w.configure(bg=T["ENTRY_BG"], fg=T["TEXT"],
                                activebackground=T["ACCENT2"],
                                highlightthickness=0)
                    w["menu"].configure(bg=T["ENTRY_BG"], fg=T["TEXT"],
                                        activebackground=T["ACCENT2"])
            except Exception:
                pass
            for child in w.winfo_children():
                walk(child)

        walk(root)

    # ── Build UI ──────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()

        # KEY: footer and totals MUST be packed BEFORE the expanding body frame
        self._build_footer()
        self._build_totals_bar()

        body = tk.Frame(self, bg=self.T["BG"])
        body.pack(fill="both", expand=True, padx=6, pady=(4, 0))
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # Responsive left panel: 22 % of screen width, clamped 220-300 px
        sw = self.winfo_screenwidth()
        panel_w = max(220, min(300, int(sw * 0.22)))
        left_wrap = tk.Frame(body, bg=self.T["PANEL"], width=panel_w)
        left_wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        left_wrap.pack_propagate(False)
        self._sf = ScrollableFrame(left_wrap, bg=self.T["PANEL"])
        self._sf.pack(fill="both", expand=True)

        right_wrap = tk.Frame(body, bg=self.T["PANEL"])
        right_wrap.grid(row=0, column=1, sticky="nsew")

        self._build_form(self._sf.inner)
        self._build_table(right_wrap)

    # ── Header ────────────────────────────────────────────────
    def _build_header(self):
        T = self.T
        hdr = tk.Frame(self, bg=T["HEADER_BG"])
        hdr.pack(fill="x")

        left = tk.Frame(hdr, bg=T["HEADER_BG"])
        left.pack(side="left", padx=12, pady=8)
        tk.Label(left, text="💼  PAYROLL PROCESSOR",
                 font=("Segoe UI", 15, "bold"), bg=T["HEADER_BG"], fg="white").pack(anchor="w")
        tk.Label(left,
                 text="Tanzania  ·  TRA PAYE 2024/25  ·  v" + VERSION + "  ·  " + self.username + " (" + self.role + ")",
                 font=("Segoe UI", 8), bg=T["HEADER_BG"], fg="white").pack(anchor="w")

        right = tk.Frame(hdr, bg=T["HEADER_BG"])
        right.pack(side="right", padx=12)

        tk.Label(right, text="🎨", font=("Segoe UI", 9),
                 bg=T["HEADER_BG"], fg="white").pack(side="left")
        om = tk.OptionMenu(right, self.theme_name, *THEMES.keys(),
                           command=self._apply_theme)
        om.configure(bg=T["ENTRY_BG"], fg=T["TEXT"], font=("Segoe UI", 8),
                     relief="flat", highlightthickness=0, padx=6)
        om["menu"].configure(bg=T["ENTRY_BG"], fg=T["TEXT"], font=("Segoe UI", 8))
        om.pack(side="left", padx=(4, 10))

        if self.role == "admin":
            b = tk.Button(right, text="👥 Users", font=("Segoe UI", 8, "bold"),
                          bg=T["BTN_SAVE"], fg="white", relief="flat", cursor="hand2",
                          command=lambda: UserManagerDialog(self, self.T))
            b.pack(side="left", padx=2, ipady=4, ipadx=6)
            ToolTip(b, "Add or remove users")

        b2 = tk.Button(right, text="🔒 Password", font=("Segoe UI", 8, "bold"),
                       bg=T["ENTRY_BG"], fg=T["TEXT"], relief="flat", cursor="hand2",
                       command=lambda: ChangePasswordDialog(self, self.T, self.username))
        b2.pack(side="left", padx=2, ipady=4, ipadx=6)
        ToolTip(b2, "Change your login password")

        if UPDATE_READY:
            b3 = tk.Button(right, text="⬆ Update", font=("Segoe UI", 8, "bold"),
                           bg="#16a085", fg="white", relief="flat", cursor="hand2",
                           command=self._manual_update_check)
            b3.pack(side="left", padx=2, ipady=4, ipadx=6)
            ToolTip(b3, "Check GitHub for a newer version")

        b4 = tk.Button(right, text="⏻ Logout", font=("Segoe UI", 8, "bold"),
                       bg=T["RED"], fg="white", relief="flat", cursor="hand2",
                       command=self._logout)
        b4.pack(side="left", padx=(2, 0), ipady=4, ipadx=6)
        ToolTip(b4, "Log out")

    # ── Form ──────────────────────────────────────────────────
    def _build_form(self, p):
        T = self.T
        tk.Label(p, text="➕  ADD EMPLOYEE", font=("Segoe UI", 10, "bold"),
                 bg=T["PANEL"], fg=T["ACCENT"]).pack(pady=(10, 2), padx=10, anchor="w")
        tk.Frame(p, bg=T["ACCENT"], height=2).pack(fill="x", padx=10)

        self.entries = {}
        self._big_field(p, "👤  Employee Name",     "Name",      "e.g. John Banda")
        self._big_field(p, "💰  Basic Salary (TZS)", "Salary",    "e.g. 500000")
        self._big_field(p, "➕  Allowance (TZS)",   "ALLOWANCE", "Leave as 0 if none")

        note = tk.Frame(p, bg=T["ENTRY_BG"])
        note.pack(fill="x", padx=10, pady=(4, 0))
        tk.Label(note, text="ℹ  P.A.Y.E is auto-calculated\n   using TRA 2024/25 brackets",
                 font=("Segoe UI", 7, "italic"), bg=T["ENTRY_BG"], fg=T["SUBTEXT"],
                 justify="left", padx=6, pady=4).pack(anchor="w")

        # Loan deduction
        tk.Frame(p, bg=T["SEP"], height=1).pack(fill="x", padx=10, pady=8)
        self.has_loan = tk.BooleanVar()
        lq = tk.Frame(p, bg=T["PANEL"])
        lq.pack(fill="x", padx=10, pady=2)
        tk.Label(lq, text="Does this employee have a loan?",
                 font=("Segoe UI", 9), bg=T["PANEL"], fg=T["TEXT"]).pack(side="left")
        tk.Checkbutton(lq, text="Yes", variable=self.has_loan,
                       font=("Segoe UI", 9, "bold"), bg=T["PANEL"], fg=T["GOLD"],
                       selectcolor=T["ENTRY_BG"], activebackground=T["PANEL"],
                       activeforeground=T["GOLD"],
                       command=self._toggle_loan).pack(side="right")
        self.loan_amount_frame = tk.Frame(p, bg=T["PANEL"])
        self.loan_amount_frame.pack(fill="x", padx=10, pady=(0, 4))
        tk.Label(self.loan_amount_frame, text="   Monthly loan deduction (TZS):",
                 font=("Segoe UI", 8), bg=T["PANEL"], fg=T["GOLD"]).pack(anchor="w")
        self.loan_var   = tk.StringVar()
        self.loan_entry = tk.Entry(self.loan_amount_frame, textvariable=self.loan_var,
                                   font=("Segoe UI", 11), bg=T["ENTRY_BG"],
                                   fg=T["ENTRY_FG"], insertbackground=T["TEXT"],
                                   relief="flat", bd=4, state="disabled")
        self.loan_entry.pack(fill="x", ipady=5, pady=(2, 0))

        # Loan board
        tk.Frame(p, bg=T["SEP"], height=1).pack(fill="x", padx=10, pady=6)
        self.has_loan_board = tk.BooleanVar()
        lbq = tk.Frame(p, bg=T["PANEL"])
        lbq.pack(fill="x", padx=10, pady=2)
        tk.Label(lbq, text="Deduct Loan Board? (15% of salary)",
                 font=("Segoe UI", 9), bg=T["PANEL"], fg=T["TEXT"]).pack(side="left")
        tk.Checkbutton(lbq, text="Yes", variable=self.has_loan_board,
                       font=("Segoe UI", 9, "bold"), bg=T["PANEL"], fg=T["GOLD"],
                       selectcolor=T["ENTRY_BG"], activebackground=T["PANEL"],
                       activeforeground=T["GOLD"],
                       command=self.preview_calc).pack(side="right")
        self.lb_info = tk.Label(p, text="   15% will be deducted from salary",
                                 font=("Segoe UI", 7, "italic"),
                                 bg=T["PANEL"], fg=T["SUBTEXT"])
        self.lb_info.pack(anchor="w", padx=14)

        # Calculated preview
        tk.Frame(p, bg=T["ACCENT"], height=1).pack(fill="x", padx=10, pady=(10, 3))
        tk.Label(p, text="📊  CALCULATED VALUES",
                 font=("Segoe UI", 8, "bold"), bg=T["PANEL"], fg=T["SUBTEXT"]
                 ).pack(anchor="w", padx=10)
        self.calc_vars = {}
        for af, ck, tip in [
            ("NSSF 10%",          "GOLD",    "10% of basic salary"),
            ("PAY",               "GOLD",    "Salary minus NSSF"),
            ("P.A.Y.E",           "RED",     "Income tax (TRA brackets)"),
            ("GROSS PAY",         "GOLD",    "Salary + Allowance"),
            ("LOAN BOARD",        "SUBTEXT", "15% if ticked above"),
            ("NET PAY",           "ACCENT",  "After NSSF and PAYE deducted"),
            ("AMOUNT TO BE PAID", "ACCENT",  "What the employee receives in hand"),
        ]:
            f = tk.Frame(p, bg=T["PANEL"])
            f.pack(fill="x", padx=10, pady=1)
            lbl = tk.Label(f, text=af, width=20, anchor="w",
                           font=("Segoe UI", 8), bg=T["PANEL"], fg=T["SUBTEXT"])
            lbl.pack(side="left")
            ToolTip(lbl, tip)
            var = tk.StringVar(value="—")
            self.calc_vars[af] = var
            tk.Label(f, textvariable=var, anchor="e",
                     font=("Segoe UI", 9, "bold"), bg=T["PANEL"], fg=T[ck]).pack(side="right")

        for key in ("Salary", "ALLOWANCE"):
            self.entries[key].trace_add("write", lambda *_: self.preview_calc())
        self.loan_var.trace_add("write", lambda *_: self.preview_calc())

        tk.Frame(p, bg=T["ACCENT"], height=2).pack(fill="x", padx=10, pady=(10, 4))
        self.btn_add = tk.Button(p, text="➕   ADD THIS EMPLOYEE",
                                  font=("Segoe UI", 11, "bold"),
                                  bg=T["BTN_SAVE"], fg="white", relief="flat",
                                  cursor="hand2", command=self.add_employee)
        self.btn_add.pack(fill="x", padx=10, ipady=10, pady=3)
        ToolTip(self.btn_add, "Click to add this employee to the payroll list")

        tk.Button(p, text="🗑  Clear / Start new entry",
                  font=("Segoe UI", 8), bg=T["PANEL"], fg=T["SUBTEXT"],
                  relief="flat", cursor="hand2",
                  command=self.clear_form).pack(pady=(0, 12))

    def _big_field(self, parent, label, key, hint=""):
        T = self.T
        f = tk.Frame(parent, bg=T["PANEL"])
        f.pack(fill="x", padx=10, pady=4)
        tk.Label(f, text=label, anchor="w", font=("Segoe UI", 9, "bold"),
                 bg=T["PANEL"], fg=T["TEXT"]).pack(anchor="w")
        var = tk.StringVar()
        tk.Entry(f, textvariable=var, font=("Segoe UI", 11),
                 bg=T["ENTRY_BG"], fg=T["ENTRY_FG"],
                 insertbackground=T["TEXT"], relief="flat", bd=4).pack(fill="x", ipady=6)
        if hint:
            tk.Label(f, text="   " + hint, font=("Segoe UI", 7),
                     bg=T["PANEL"], fg=T["SUBTEXT"]).pack(anchor="w")
        self.entries[key] = var

    # ── Table ─────────────────────────────────────────────────
    def _build_table(self, p):
        T = self.T
        p.rowconfigure(1, weight=1)
        p.columnconfigure(0, weight=1)

        hrow = tk.Frame(p, bg=T["PANEL"])
        hrow.grid(row=0, column=0, sticky="ew")
        tk.Label(hrow, text="📋  PAYROLL REGISTER",
                 font=("Segoe UI", 10, "bold"), bg=T["PANEL"], fg=T["ACCENT"]
                 ).pack(side="left", padx=10, pady=(10, 2))
        self.count_var = tk.StringVar(value="No employees yet")
        tk.Label(hrow, textvariable=self.count_var,
                 font=("Segoe UI", 8), bg=T["PANEL"], fg=T["SUBTEXT"]
                 ).pack(side="right", padx=10, pady=(10, 2))
        tk.Frame(p, bg=T["ACCENT"], height=2).grid(
            row=0, column=0, sticky="ew", padx=8, pady=(30, 0))

        tf = tk.Frame(p, bg=T["PANEL"])
        tf.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)

        s = ttk.Style()
        s.theme_use("clam")
        s.configure("P.Treeview", background=T["TREE_BG"], foreground=T["TREE_FG"],
                    fieldbackground=T["TREE_BG"], rowheight=27, font=("Segoe UI", 9))
        s.configure("P.Treeview.Heading", background=T["TREE_HEAD"],
                    foreground="white", font=("Segoe UI", 8, "bold"), relief="flat")
        s.map("P.Treeview", background=[("selected", T["ACCENT2"])])

        self.tree = ttk.Treeview(tf, columns=HEADERS, show="headings",
                                  style="P.Treeview", height=15)
        col_w = {
            "Name": 130, "Salary": 90, "NSSF 10%": 80, "PAY": 80,
            "P.A.Y.E": 80, "ALLOWANCE": 90, "GROSS PAY": 90,
            "LOAN Deduction": 95, "LOAN BOARD (15%)": 100,
            "NET PAY": 85, "AMOUNT TO BE PAID": 115,
        }
        for c in HEADERS:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=col_w.get(c, 85),
                             anchor="center", minwidth=60, stretch=True)
        self.tree.column("Name", anchor="w", stretch=True)

        vsb = ttk.Scrollbar(tf, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        rm = tk.Frame(p, bg=T["PANEL"])
        rm.grid(row=2, column=0, sticky="e", padx=6, pady=(0, 4))
        self.btn_remove = tk.Button(rm, text="🗑  Remove Selected Row",
                                     font=("Segoe UI", 8),
                                     bg=T["RED"], fg="white", relief="flat",
                                     cursor="hand2", command=self.remove_selected)
        self.btn_remove.pack(ipady=4, ipadx=8)
        ToolTip(self.btn_remove, "Click on a row first, then click here to remove it")

    # ── Totals bar ────────────────────────────────────────────
    def _build_totals_bar(self):
        T = self.T
        self.totals_bar = tk.Frame(self, bg=T["PANEL"], pady=5)
        self.totals_bar.pack(fill="x", padx=6, pady=(1, 0), side="bottom")

        tk.Label(self.totals_bar, text="📊 TOTALS:",
                 font=("Segoe UI", 8, "bold"), bg=T["PANEL"], fg=T["SUBTEXT"]
                 ).pack(side="left", padx=(10, 4))

        self._tot_vars = {}
        for label, key, ck in [
            ("Staff",              "count",        "SUBTEXT"),
            ("Total Salary",       "total_salary", "GOLD"),
            ("NSSF",               "total_nssf",   "SUBTEXT"),
            ("PAYE Tax",           "total_paye",   "RED"),
            ("Net Pay",            "total_net",    "GOLD"),
            ("AMOUNT TO PAY OUT",  "total_amount", "ACCENT"),
        ]:
            box = tk.Frame(self.totals_bar, bg=T["ENTRY_BG"], padx=8, pady=3)
            box.pack(side="left", padx=3)
            tk.Label(box, text=label, font=("Segoe UI", 6),
                     bg=T["ENTRY_BG"], fg=T["SUBTEXT"]).pack()
            var = tk.StringVar(value="—")
            self._tot_vars[key] = var
            tk.Label(box, textvariable=var, font=("Segoe UI", 9, "bold"),
                     bg=T["ENTRY_BG"], fg=T[ck]).pack()

    def _refresh_totals_bar(self):
        if not hasattr(self, "_tot_vars"):
            return
        if not self.records:
            for var in self._tot_vars.values():
                var.set("—")
            return
        self._tot_vars["count"].set(str(len(self.records)))
        self._tot_vars["total_salary"].set("TZS " + f"{sum(r['salary'] for r in self.records):,.0f}")
        self._tot_vars["total_nssf"].set("TZS " + f"{sum(r['nssf'] for r in self.records):,.0f}")
        self._tot_vars["total_paye"].set("TZS " + f"{sum(r['paye'] for r in self.records):,.0f}")
        self._tot_vars["total_net"].set("TZS " + f"{sum(r['net_pay'] for r in self.records):,.0f}")
        self._tot_vars["total_amount"].set("TZS " + f"{sum(r['amount_paid'] for r in self.records):,.0f}")
        # Re-apply theme colours to boxes on theme switch
        T = self.T
        for child in self.totals_bar.winfo_children():
            if isinstance(child, tk.Frame):
                child.configure(bg=T["ENTRY_BG"])
                for lbl in child.winfo_children():
                    if isinstance(lbl, tk.Label):
                        lbl.configure(bg=T["ENTRY_BG"])

    # ── Footer — always visible at the bottom ─────────────────
    def _build_footer(self):
        T  = self.T
        sw = self.winfo_screenwidth()
        btn_fs = 10 if sw >= 1366 else 9

        foot = tk.Frame(self, bg=T["PANEL"], pady=6)
        foot.pack(fill="x", side="bottom")

        # Month row
        mr = tk.Frame(foot, bg=T["PANEL"])
        mr.pack(fill="x", padx=14, pady=(0, 5))
        tk.Label(mr, text="📅  Payroll Month:",
                 font=("Segoe UI", 9), bg=T["PANEL"], fg=T["TEXT"]).pack(side="left")
        self.month_var = tk.StringVar(
            value=datetime.datetime.now().strftime("%B %Y"))
        tk.Entry(mr, textvariable=self.month_var, font=("Segoe UI", 9),
                 bg=T["ENTRY_BG"], fg=T["ENTRY_FG"], insertbackground=T["TEXT"],
                 relief="flat", bd=3, width=14).pack(side="left", padx=8, ipady=3)
        tk.Label(mr, text="← change before saving or printing",
                 font=("Segoe UI", 7, "italic"),
                 bg=T["PANEL"], fg=T["SUBTEXT"]).pack(side="left")

        # Buttons row
        bf = tk.Frame(foot, bg=T["PANEL"])
        bf.pack(padx=14, fill="x")

        self.btn_save = tk.Button(
            bf, text="💾  Save to Excel",
            font=("Segoe UI", btn_fs, "bold"),
            bg=T["BTN_SAVE"], fg="white", relief="flat", cursor="hand2",
            command=self.save_excel)
        self.btn_save.pack(side="left", expand=True, fill="x", padx=3, ipady=9)
        ToolTip(self.btn_save, "Save payroll to Excel — you choose the folder")

        self.btn_print = tk.Button(
            bf, text="🖨  Print Payroll",
            font=("Segoe UI", btn_fs, "bold"),
            bg=T["BTN_PRINT"], fg="white", relief="flat", cursor="hand2",
            command=self.print_payroll)
        self.btn_print.pack(side="left", expand=True, fill="x", padx=3, ipady=9)
        ToolTip(self.btn_print, "Open the payroll in your browser then print with Ctrl+P")

        self.btn_open = tk.Button(
            bf, text="📂  Open Last Excel",
            font=("Segoe UI", btn_fs, "bold"),
            bg=T["BTN_OPEN"], fg="white", relief="flat", cursor="hand2",
            command=self.open_file)
        self.btn_open.pack(side="left", expand=True, fill="x", padx=3, ipady=9)
        ToolTip(self.btn_open, "Open the Excel file you last saved")

        self.btn_reset = tk.Button(
            bf, text="🔄  New Session",
            font=("Segoe UI", btn_fs, "bold"),
            bg=T["SUBTEXT"], fg="white", relief="flat", cursor="hand2",
            command=self.new_session)
        self.btn_reset.pack(side="left", expand=True, fill="x", padx=3, ipady=9)
        ToolTip(self.btn_reset, "Clear the list and start a new payroll")

    # ── Logic ─────────────────────────────────────────────────
    def _toggle_loan(self):
        if self.has_loan.get():
            self.loan_entry.configure(state="normal")
            self.loan_entry.focus_set()
        else:
            self.loan_var.set("")
            self.loan_entry.configure(state="disabled")
        self.preview_calc()

    def _num(self, key):
        try:
            return float(self.entries[key].get().replace(",", "").strip() or 0)
        except Exception:
            return 0.0

    def _loan_amt(self):
        try:
            return float(self.loan_var.get().replace(",", "") or 0)
        except Exception:
            return 0.0

    def preview_calc(self, *_):
        s = self._num("Salary")
        a = self._num("ALLOWANCE")
        nssf, pay, paye, gross, ld, lb, net, amt = calculate(
            s, a, self.has_loan.get(), self._loan_amt(), self.has_loan_board.get())
        fmt = lambda v: "TZS  " + f"{v:,.2f}" if s else "—"
        self.calc_vars["NSSF 10%"].set(fmt(nssf))
        self.calc_vars["PAY"].set(fmt(pay))
        self.calc_vars["P.A.Y.E"].set(fmt(paye))
        self.calc_vars["GROSS PAY"].set(fmt(gross))
        self.calc_vars["LOAN BOARD"].set(
            fmt(lb) if self.has_loan_board.get() else "Not selected")
        self.calc_vars["NET PAY"].set(fmt(net))
        self.calc_vars["AMOUNT TO BE PAID"].set(fmt(amt))
        self.lb_info.configure(
            text="   Deduction will be  TZS " + f"{lb:,.2f}"
            if self.has_loan_board.get() and s
            else "   15% will be deducted from salary")

    def add_employee(self):
        name   = self.entries["Name"].get().strip()
        salary = self._num("Salary")
        if not name:
            messagebox.showwarning("⚠  Missing Name",
                "Please type the employee's full name.")
            return
        if salary <= 0:
            messagebox.showwarning("⚠  Missing Salary",
                "Please enter the employee's salary.")
            return
        if self.has_loan.get() and self._loan_amt() <= 0:
            messagebox.showwarning("⚠  Missing Loan Amount",
                "You ticked 'Has loan?' — please enter how much to deduct each month.")
            return

        allowance = self._num("ALLOWANCE")
        nssf, pay, paye, gross, ld, lb, net, amt = calculate(
            salary, allowance, self.has_loan.get(),
            self._loan_amt(), self.has_loan_board.get())
        rec = dict(name=name, salary=salary, nssf=nssf, pay=pay, paye=paye,
                   allowance=allowance, gross_pay=gross, loan_ded=ld,
                   loan_board=lb, net_pay=net, amount_paid=amt)
        self.records.append(rec)
        self.tree.insert("", "end", values=(
            name,
            f"{salary:,.2f}", f"{nssf:,.2f}", f"{pay:,.2f}",
            f"{paye:,.2f}", f"{allowance:,.2f}", f"{gross:,.2f}",
            f"{ld:,.2f}", f"{lb:,.2f}", f"{net:,.2f}", f"{amt:,.2f}",
        ))
        n = len(self.records)
        self.count_var.set(str(n) + " employee" + ("s" if n != 1 else "") + " added")
        self._refresh_totals_bar()
        self.clear_form()
        messagebox.showinfo("✅  Added",
            name + " has been added.\n\n"
            "Amount to be paid:  TZS " + f"{amt:,.2f}" + "\n\n"
            "Total employees now: " + str(n))

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Nothing selected",
                "Please click on an employee row in the table first.")
            return
        if messagebox.askyesno("Confirm", "Remove the selected employee?"):
            for item in sel:
                idx = self.tree.index(item)
                self.tree.delete(item)
                self.records.pop(idx)
            n = len(self.records)
            self.count_var.set(
                str(n) + " employee" + ("s" if n != 1 else "") + " added"
                if n else "No employees yet")
            self._refresh_totals_bar()

    def clear_form(self):
        for v in self.entries.values():
            v.set("")
        self.has_loan.set(False)
        self.loan_var.set("")
        self.loan_entry.configure(state="disabled")
        self.has_loan_board.set(False)
        for v in self.calc_vars.values():
            v.set("—")
        self.lb_info.configure(text="   15% will be deducted from salary")

    def save_excel(self):
        if not self.records:
            messagebox.showwarning("⚠  Nothing to Save",
                "Please add at least one employee first.")
            return
        month = self.month_var.get().replace(" ", "_")
        fp = filedialog.asksaveasfilename(
            title="Choose where to save the Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            initialfile="Payroll_" + month + ".xlsx")
        if not fp:
            return
        try:
            save_to_excel(self.records, fp, self.username, self.month_var.get())
            self._last_saved_path = fp
            messagebox.showinfo("✅  Saved",
                "Excel file saved to:\n\n" + fp + "\n\n"
                + str(len(self.records)) + " employee(s) included.\n\n"
                "Click  📂 Open Last Excel  to view it.")
        except Exception as e:
            messagebox.showerror("❌  Error saving file", str(e))

    def print_payroll(self):
        if not self.records:
            messagebox.showwarning("⚠  Nothing to Print",
                "Please add at least one employee first.")
            return
        html = make_print_html(self.records, self.month_var.get(), self.username)
        open_print_in_browser(html)
        messagebox.showinfo("🖨  How to Print",
            "The payroll has opened in your web browser.\n\n"
            "To print it, do this:\n\n"
            "  1.  Look at the browser window that just opened\n"
            "  2.  Hold the  Ctrl  key and press  P\n"
            "       (that is the keyboard shortcut for Print)\n"
            "  3.  Your printer list will appear\n"
            "  4.  Choose your printer and click  Print\n\n"
            "Tip: choose  Save as PDF  to save a PDF file instead.")

    def open_file(self):
        if not self._last_saved_path:
            messagebox.showinfo("No file yet",
                "You have not saved an Excel file yet.\n\n"
                "Click  💾 Save to Excel  first.")
            return
        if not os.path.exists(self._last_saved_path):
            messagebox.showinfo("File not found",
                "The file could not be found:\n" + self._last_saved_path)
            return
        try:
            os.startfile(self._last_saved_path)
        except AttributeError:
            subprocess.call(["xdg-open", self._last_saved_path])

    def new_session(self):
        if self.records and not messagebox.askyesno(
                "🔄  New Session",
                "This will clear all employees from the list.\n\n"
                "⚠  Make sure you have already saved the Excel file!\n\n"
                "Continue?"):
            return
        self.records.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.count_var.set("No employees yet")
        self._refresh_totals_bar()
        self.clear_form()
        self._last_saved_path = None

    # ── Update methods ────────────────────────────────────────
    def _auto_update_check(self):
        """Silent check on startup — only interrupts if update found."""
        def on_found(latest):
            self.after(0, lambda: self._prompt_update(latest))
        check_for_update_bg(on_found)

    def _manual_update_check(self):
        """User clicked the Update button — show result either way."""
        self._update_result = None
        result_holder = [None]
        check_for_update_bg_with_result(result_holder)

        def poll(elapsed=0):
            if elapsed >= 8000:
                messagebox.showinfo("✅  Up to Date",
                    "You are already on the latest version (v" + VERSION + ").\n\n"
                    "No update available right now.")
                return
            val = result_holder[0]
            if val is None:
                self.after(300, lambda: poll(elapsed + 300))
            elif val == "up_to_date":
                messagebox.showinfo("✅  Up to Date",
                    "You are already on the latest version (v" + VERSION + ").")
            elif val == "error":
                messagebox.showwarning("⚠  Could Not Check",
                    "Could not connect to GitHub to check for updates.\n\n"
                    "Please check your internet connection.")
            else:
                self._prompt_update(val)

        self.after(300, lambda: poll(300))

    def _prompt_update(self, latest):
        if messagebox.askyesno("🔄  Update Available",
                "A new version is available!\n\n"
                "   Your version :  v" + VERSION + "\n"
                "   New version  :  v" + latest + "\n\n"
                "Download and install now?\n"
                "(The app will restart automatically.)"):
            download_and_install(latest, self)

    def _logout(self):
        if messagebox.askyesno("⏻  Logout",
                "Are you sure you want to log out?\n\n"
                "⚠  Save your work first!"):
            self.destroy()
            start_app()

# ══════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════
def start_app():
    login = LoginWindow()
    login.mainloop()
    if login.logged_in_user:
        app = PayrollApp(login.logged_in_user, login.logged_in_role)
        app.mainloop()

if __name__ == "__main__":
    start_app()
