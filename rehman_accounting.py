import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
import csv
import os
import shutil
import logging
import tempfile
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

# ================= DATABASE =================
class Database:
    def __init__(self, path: str):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.cur = self.conn.cursor()
        self._init_schema()

    def _init_schema(self):
        self.cur.execute("""
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            type TEXT NOT NULL CHECK (type IN ('Asset','Liability','Income','Expense'))
        )
        """)

        self.cur.execute("""
        CREATE TABLE IF NOT EXISTS vouchers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            description TEXT
        )
        """)

        self.cur.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voucher_id INTEGER NOT NULL,
            account_id INTEGER NOT NULL,
            debit REAL NOT NULL DEFAULT 0 CHECK (debit >= 0),
            credit REAL NOT NULL DEFAULT 0 CHECK (credit >= 0),
            CHECK (
                (debit > 0 AND credit = 0) OR
                (credit > 0 AND debit = 0)
            ),
            FOREIGN KEY (voucher_id) REFERENCES vouchers(id) ON DELETE CASCADE,
            FOREIGN KEY (account_id) REFERENCES accounts(id)
        )
        """)

        self.cur.execute("CREATE INDEX IF NOT EXISTS idx_tx_voucher ON transactions(voucher_id)")
        self.cur.execute("CREATE INDEX IF NOT EXISTS idx_tx_account ON transactions(account_id)")
        self.conn.commit()

    def execute(self, query, params=(), commit=True):
        self.cur.execute(query, params)
        if commit:
            self.conn.commit()
        return self.cur

    def fetch_one(self, query, params=()):
        return self.cur.execute(query, params).fetchone()

    def fetch_all(self, query, params=()):
        return self.cur.execute(query, params).fetchall()

    def commit(self):
        self.conn.commit()

    def close(self):
        self.conn.close()

# ================= PATHS =================
APP_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.path.expanduser("~"), "RIAS_Data")
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "rehman_industries.db")
SEED_DIR = getattr(sys, "_MEIPASS", APP_DIR)
SEED_DB_PATH = os.path.join(SEED_DIR, "rehman_industries.db")

if not os.path.exists(DB_PATH) and os.path.exists(SEED_DB_PATH):
    shutil.copy2(SEED_DB_PATH, DB_PATH)

db = Database(DB_PATH)

# ================= RELIABILITY =================
logging.basicConfig(
    filename=os.path.join(DATA_DIR, "rias.log"),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

def backup_database(tag="manual"):
    if not os.path.exists(db.path):
        return
    backup_dir = os.path.join(DATA_DIR, "db_backups")
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"rehman_industries_{tag}_{timestamp}.db"
    backup_path = os.path.join(backup_dir, backup_name)
    shutil.copy2(db.path, backup_path)

def check_db_integrity():
    try:
        result = db.fetch_one("PRAGMA integrity_check")
        if not result or result[0].lower() != "ok":
            messagebox.showwarning(
                "Database Check",
                f"Integrity check reported issues:\n{result[0] if result else 'Unknown error'}",
            )
        else:
            set_status("Database integrity OK")
    except Exception:
        logging.exception("Integrity check failed")
        messagebox.showwarning("Database Check", "Integrity check failed. See rias.log.")

# ================= MAIN WINDOW =================
root = tk.Tk()
root.title("Rehman Industries Accounting System (RIAS)")
root.geometry("1400x850")

# ================= STYLING =================
style = ttk.Style()
style.theme_use("clam")

PRIMARY = "#0b1220"
ACCENT = "#2563eb"
ACCENT_SOFT = "#e7f0ff"
SURFACE = "#f6f7fb"
SURFACE_ALT = "#ffffff"
MUTED = "#6b7280"
LINE = "#e5e7eb"

style.configure("TFrame", background=SURFACE)
style.configure("TLabel", background=SURFACE, foreground=PRIMARY, font=("Segoe UI", 10))
style.configure("TButton", font=("Segoe UI", 10), padding=(10, 6))
style.map("TButton", background=[("active", "#e7edf7")])

style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"))
style.configure("SubHeader.TLabel", font=("Segoe UI", 10, "bold"), foreground=MUTED)
style.configure("Muted.TLabel", font=("Segoe UI", 9), foreground=MUTED)

style.configure("Card.TFrame", background=SURFACE_ALT, borderwidth=1, relief="solid")
style.configure("Card.TLabel", background=SURFACE_ALT, font=("Segoe UI", 12, "bold"))

style.configure("Sidebar.TFrame", background=PRIMARY)
style.configure("Sidebar.TLabel", background=PRIMARY, foreground="white", font=("Segoe UI", 12, "bold"))
style.configure("Nav.TButton", background=PRIMARY, foreground="white", anchor="w", padding=(12, 10))
style.map(
    "Nav.TButton",
    background=[("active", "#111827"), ("selected", "#1f2937")],
    foreground=[("active", "white"), ("selected", "white")],
)

style.configure("Primary.TButton", background=ACCENT, foreground="white", padding=(12, 8))
style.map("Primary.TButton", background=[("active", "#1d4ed8")], foreground=[("active", "white")])

style.configure("Treeview", font=("Segoe UI", 10), rowheight=26, background=SURFACE_ALT, fieldbackground=SURFACE_ALT)
style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

# ================= STATUS BAR =================
status_var = tk.StringVar(value="Ready")
status_bar = ttk.Label(root, textvariable=status_var, relief="sunken", anchor="w", padding=(10, 6))
status_bar.pack(side="bottom", fill="x")

def set_status(msg, duration=5000):
    status_var.set(msg)
    root.after(duration, lambda: status_var.set("Ready"))

# ================= MAIN LAYOUT =================
main = ttk.Frame(root)
main.pack(fill="both", expand=True)

sidebar = ttk.Frame(main, width=260, padding=20, style="Sidebar.TFrame")
sidebar.pack(side="left", fill="y")

content = ttk.Frame(main, padding=0)
content.pack(side="right", fill="both", expand=True)

content_header = ttk.Frame(content, padding=(24, 18), style="TFrame")
content_header.pack(fill="x")
ttk.Label(content_header, text="Rehman Industries Accounting", style="Header.TLabel").pack(anchor="w")
ttk.Label(content_header, text=datetime.now().strftime("%A, %B %d, %Y"), style="Muted.TLabel").pack(anchor="w")

content_area = ttk.Frame(content, padding=24)
content_area.pack(fill="both", expand=True)

# ================= NAVIGATION =================
frames = {}
nav_buttons = {}
current_frame_name = None

def show_frame(name, frame):
    global current_frame_name
    for f in frames.values():
        f.pack_forget()
    frame.pack(fill="both", expand=True)
    set_active_nav(name)
    current_frame_name = name

def nav_button(text, frame):
    btn = ttk.Button(sidebar, text=text, style="Nav.TButton", command=lambda: show_frame(text, frame))
    btn.pack(fill="x", pady=6)
    nav_buttons[text] = btn

def set_active_nav(name):
    for key, btn in nav_buttons.items():
        btn.state(["!selected"])
        if key == name:
            btn.state(["selected"])

# ================= UTILITY FUNCTIONS =================
def execute_query(query, params=(), commit=True):
    return db.execute(query, params, commit=commit)

def fetch_one(query, params=()):
    return db.fetch_one(query, params)

def fetch_all(query, params=()):
    return db.fetch_all(query, params)

# ================= REUSABLE TREEVIEW =================
def create_treeview(parent, columns, headings, height=15, numeric_cols=None, pack=True):
    tree = ttk.Treeview(parent, columns=columns, show="headings", height=height)
    numeric_cols = set(numeric_cols or [])
    tree.tag_configure("odd", background="#ffffff")
    tree.tag_configure("even", background="#f1f5f9")
    for col, hd in zip(columns, headings):
        is_numeric = col in numeric_cols
        anchor = "e" if is_numeric else "w"
        tree.heading(col, text=hd, anchor=anchor)
        tree.column(col, anchor=anchor)
    if pack:
        tree.pack(fill="both", expand=True, pady=12)
    return tree

def fmt(amount):
    return f"{amount:,.2f}"

def insert_tree_row(tree, values, row_index):
    tag = "even" if row_index % 2 == 0 else "odd"
    tree.insert("", "end", values=values, tags=(tag,))

def is_valid_date(value):
    try:
        datetime.strptime(value, "%Y-%m-%d")
        return True
    except ValueError:
        return False

# ================= CHART OF ACCOUNTS =================
frame_accounts = ttk.Frame(content_area)
frames["Accounts"] = frame_accounts
ttk.Label(frame_accounts, text="Chart of Accounts", style="Header.TLabel").pack(anchor="w")
ttk.Label(frame_accounts, text="Manage your accounts list", style="SubHeader.TLabel").pack(anchor="w", pady=(2, 12))

form_acc = ttk.Frame(frame_accounts)
form_acc.pack(anchor="w", pady=10)

ttk.Label(form_acc, text="Account Name").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
ttk.Label(form_acc, text="Account Type").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)

acc_name = ttk.Entry(form_acc, width=40)
acc_type = ttk.Combobox(form_acc, values=["Asset", "Liability", "Income", "Expense"], state="readonly", width=37)
acc_name.grid(row=0, column=1, pady=4)
acc_type.grid(row=1, column=1, pady=4)

ttk.Label(form_acc, text="Search").grid(row=0, column=2, sticky="w", padx=(20, 6), pady=4)
search_acc = ttk.Entry(form_acc, width=26)
search_acc.grid(row=0, column=3, pady=4, sticky="w")
ttk.Label(form_acc, text="Type Filter").grid(row=1, column=2, sticky="w", padx=(20, 6), pady=4)
type_filter = ttk.Combobox(form_acc, values=["All", "Asset", "Liability", "Income", "Expense"], state="readonly", width=23)
type_filter.grid(row=1, column=3, pady=4, sticky="w")
type_filter.set("All")

accounts_tree = create_treeview(frame_accounts, ["Name", "Type"], ["Account Name", "Type"])
selected_account_name = None

def refresh_accounts(filter_text=""):
    accounts_tree.delete(*accounts_tree.get_children())
    filter_type = type_filter.get()
    if filter_type and filter_type != "All":
        rows = fetch_all(
            "SELECT name, type FROM accounts WHERE name LIKE ? AND type=? ORDER BY name",
            (f"%{filter_text}%", filter_type),
        )
    else:
        rows = fetch_all(
            "SELECT name, type FROM accounts WHERE name LIKE ? ORDER BY name",
            (f"%{filter_text}%",),
        )
    for i, r in enumerate(rows):
        insert_tree_row(accounts_tree, r, i)
    refresh_dropdowns()

def add_account():
    name = acc_name.get().strip()
    acc_type_value = acc_type.get().strip()
    if not name or not acc_type_value:
        messagebox.showerror("Error", "All fields required")
        return
    try:
        existing = fetch_one("SELECT 1 FROM accounts WHERE LOWER(name)=LOWER(?)", (name,))
        if existing:
            messagebox.showerror("Error", "Account already exists")
            return
        execute_query("INSERT INTO accounts (name, type) VALUES (?, ?)", (name, acc_type_value))
        acc_name.delete(0, tk.END)
        acc_type.set("")
        refresh_accounts()
        acc_name.focus_set()
        set_status("Account added")
    except Exception:
        logging.exception("Add account failed")
        messagebox.showerror("Error", "Account already exists")

def clear_account_form():
    global selected_account_name
    selected_account_name = None
    acc_name.delete(0, tk.END)
    acc_type.set("")
    acc_name.focus_set()
    set_status("Form cleared")

def update_account():
    global selected_account_name
    if not selected_account_name:
        messagebox.showerror("Error", "Select an account to update")
        return
    name = acc_name.get().strip()
    acc_type_value = acc_type.get().strip()
    if not name or not acc_type_value:
        messagebox.showerror("Error", "All fields required")
        return
    try:
        existing = fetch_one(
            "SELECT 1 FROM accounts WHERE LOWER(name)=LOWER(?) AND name<>?",
            (name, selected_account_name),
        )
        if existing:
            messagebox.showerror("Error", "Account name already exists")
            return
        execute_query(
            "UPDATE accounts SET name=?, type=? WHERE name=?",
            (name, acc_type_value, selected_account_name),
        )
        selected_account_name = name
        refresh_accounts()
        acc_name.focus_set()
        set_status("Account updated")
    except sqlite3.IntegrityError:
        messagebox.showerror("Error", "Account name already exists")
    except Exception:
        logging.exception("Update account failed")
        messagebox.showerror("Error", "Failed to update account")

def delete_account():
    global selected_account_name
    if not selected_account_name:
        messagebox.showerror("Error", "Select an account to delete")
        return
    if not messagebox.askyesno("Confirm", "Delete selected account?"):
        return
    try:
        execute_query("DELETE FROM accounts WHERE name=?", (selected_account_name,))
        selected_account_name = None
        refresh_accounts()
        clear_account_form()
        acc_name.focus_set()
        set_status("Account deleted")
    except sqlite3.IntegrityError:
        messagebox.showerror("Error", "Account is used in transactions")
    except Exception:
        logging.exception("Delete account failed")
        messagebox.showerror("Error", "Failed to delete account")

def on_account_select(event):
    global selected_account_name
    selection = accounts_tree.selection()
    if not selection:
        return
    values = accounts_tree.item(selection[0], "values")
    if not values:
        return
    selected_account_name = values[0]
    acc_name.delete(0, tk.END)
    acc_name.insert(0, values[0])
    acc_type.set(values[1])

accounts_tree.bind("<<TreeviewSelect>>", on_account_select)
accounts_tree.bind("<Delete>", lambda e: delete_account())

def on_account_search(event=None):
    refresh_accounts(search_acc.get().strip())

search_acc.bind("<KeyRelease>", on_account_search)
type_filter.bind("<<ComboboxSelected>>", lambda e: refresh_accounts(search_acc.get().strip()))
search_acc.bind("<Escape>", lambda e: (search_acc.delete(0, tk.END), refresh_accounts()))

actions_acc = ttk.Frame(frame_accounts)
actions_acc.pack(anchor="w", pady=6)
ttk.Button(actions_acc, text="Add Account", command=add_account, style="Primary.TButton").grid(row=0, column=0, padx=(0, 8))
ttk.Button(actions_acc, text="Update Selected", command=update_account).grid(row=0, column=1, padx=(0, 8))
ttk.Button(actions_acc, text="Delete Selected", command=delete_account).grid(row=0, column=2, padx=(0, 8))
ttk.Button(actions_acc, text="Clear", command=clear_account_form).grid(row=0, column=3)
ttk.Button(actions_acc, text="Clear Search", command=lambda: (search_acc.delete(0, tk.END), refresh_accounts())).grid(row=0, column=4, padx=(8, 0))

# Make primary action visible near inputs
actions_acc_top = ttk.Frame(form_acc)
actions_acc_top.grid(row=2, column=1, sticky="w", pady=(6, 0))
ttk.Button(actions_acc_top, text="Add Account", command=add_account, style="Primary.TButton").grid(row=0, column=0, padx=(0, 8))
ttk.Button(actions_acc_top, text="Delete Selected", command=delete_account).grid(row=0, column=1, padx=(0, 8))

# ================= VOUCHER ENTRY =================
frame_voucher = ttk.Frame(content_area)
frames["Voucher Entry"] = frame_voucher
ttk.Label(frame_voucher, text="Voucher Entry", style="Header.TLabel").pack(anchor="w")
ttk.Label(frame_voucher, text="Record balanced debit and credit lines", style="SubHeader.TLabel").pack(anchor="w", pady=(2, 12))
voucher_mode_var = tk.StringVar(value="New Voucher")
ttk.Label(frame_voucher, textvariable=voucher_mode_var, style="Muted.TLabel").pack(anchor="w")

form_voucher = ttk.Frame(frame_voucher)
form_voucher.pack(anchor="w", pady=10)

ttk.Label(form_voucher, text="Date").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
ttk.Label(form_voucher, text="Description").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
ttk.Label(form_voucher, text="Account").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=4)
ttk.Label(form_voucher, text="Debit").grid(row=2, column=1, sticky="w", padx=(0, 10), pady=4)
ttk.Label(form_voucher, text="Credit").grid(row=2, column=2, sticky="w", padx=(0, 10), pady=4)

date_entry = ttk.Entry(form_voucher, width=20)
date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
desc_entry = ttk.Entry(form_voucher, width=60)
account_cb = ttk.Combobox(form_voucher, state="readonly", width=30)
debit_entry = ttk.Entry(form_voucher, width=15)
credit_entry = ttk.Entry(form_voucher, width=15)

date_entry.grid(row=0, column=1, pady=4, sticky="w")
desc_entry.grid(row=1, column=1, columnspan=3, pady=4, sticky="w")
account_cb.grid(row=2, column=0, pady=4, sticky="w")
debit_entry.grid(row=2, column=1, pady=4, sticky="w")
credit_entry.grid(row=2, column=2, pady=4, sticky="w")

ttk.Button(form_voucher, text="Today", command=lambda: date_entry.delete(0, tk.END) or date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))).grid(row=0, column=2, padx=(6, 0), sticky="w")

voucher_hint = ttk.Label(frame_voucher, text="", style="Muted.TLabel")
voucher_hint.pack(anchor="w")

voucher_tree = create_treeview(
    frame_voucher,
    ["Account", "Debit", "Credit"],
    ["Account", "Debit", "Credit"],
    height=12,
    numeric_cols=["Debit", "Credit"],
)
entries = []
current_voucher_id = None
voucher_tree.bind("<Delete>", lambda e: remove_selected_line())

totals_frame = ttk.Frame(frame_voucher)
totals_frame.pack(anchor="w", pady=(6, 0))
total_debit_var = tk.StringVar(value="0.00")
total_credit_var = tk.StringVar(value="0.00")
balance_diff_var = tk.StringVar(value="0.00")
ttk.Label(totals_frame, text="Total Debit:", style="Muted.TLabel").grid(row=0, column=0, padx=(0, 6))
ttk.Label(totals_frame, textvariable=total_debit_var, style="SubHeader.TLabel").grid(row=0, column=1, padx=(0, 12))
ttk.Label(totals_frame, text="Total Credit:", style="Muted.TLabel").grid(row=0, column=2, padx=(0, 6))
ttk.Label(totals_frame, textvariable=total_credit_var, style="SubHeader.TLabel").grid(row=0, column=3, padx=(0, 12))
ttk.Label(totals_frame, text="Difference:", style="Muted.TLabel").grid(row=0, column=4, padx=(0, 6))
ttk.Label(totals_frame, textvariable=balance_diff_var, style="SubHeader.TLabel").grid(row=0, column=5)

def update_voucher_totals():
    total_debit = sum(e[2] for e in entries)
    total_credit = sum(e[3] for e in entries)
    total_debit_var.set(fmt(total_debit))
    total_credit_var.set(fmt(total_credit))
    balance_diff_var.set(fmt(abs(total_debit - total_credit)))
    if entries and round(total_debit, 2) != round(total_credit, 2):
        voucher_hint.config(text="Totals must balance to save.")
    elif entries:
        voucher_hint.config(text="")

def set_voucher_mode(editing):
    voucher_mode_var.set("Edit Voucher" if editing else "New Voucher")
    label = "Update Voucher" if editing else "Save Voucher"
    save_btn.config(text=label)
    top_save_btn.config(text=label)

def clear_voucher_form():
    global current_voucher_id
    current_voucher_id = None
    entries.clear()
    voucher_tree.delete(*voucher_tree.get_children())
    date_entry.delete(0, tk.END)
    date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
    desc_entry.delete(0, tk.END)
    account_cb.set("")
    debit_entry.delete(0, tk.END)
    credit_entry.delete(0, tk.END)
    voucher_hint.config(text="")
    set_voucher_mode(False)
    update_voucher_totals()
    set_status("New voucher ready")

def rebuild_voucher_tree():
    voucher_tree.delete(*voucher_tree.get_children())
    for i, e in enumerate(entries):
        insert_tree_row(voucher_tree, (e[1], fmt(e[2]), fmt(e[3])), i)

def remove_selected_line():
    selection = voucher_tree.selection()
    if not selection:
        voucher_hint.config(text="Select a line to remove.")
        return
    idx = voucher_tree.index(selection[0])
    if idx < 0 or idx >= len(entries):
        return
    entries.pop(idx)
    rebuild_voucher_tree()
    update_voucher_totals()
    voucher_hint.config(text="Line removed.")

def add_line():
    if not is_valid_date(date_entry.get().strip()):
        voucher_hint.config(text="Date must be in YYYY-MM-DD format.")
        return
    try:
        d = float(debit_entry.get() or 0)
        c = float(credit_entry.get() or 0)
    except ValueError:
        voucher_hint.config(text="Enter valid numbers for Debit/Credit.")
        return
    if d < 0 or c < 0:
        voucher_hint.config(text="Debit/Credit must be positive values.")
        return
    if (d > 0 and c > 0) or (d == 0 and c == 0):
        voucher_hint.config(text="Enter either debit or credit, not both or none.")
        return
    row = fetch_one("SELECT id FROM accounts WHERE name=?", (account_cb.get(),))
    if not row:
        voucher_hint.config(text="Select a valid account.")
        return
    acc_id = row[0]
    entries.append((acc_id, account_cb.get(), d, c))
    insert_tree_row(voucher_tree, (account_cb.get(), fmt(d), fmt(c)), len(entries) - 1)
    debit_entry.delete(0, tk.END)
    credit_entry.delete(0, tk.END)
    voucher_hint.config(text="Line added.")
    update_voucher_totals()
    account_cb.focus_set()

def save_voucher():
    if not entries:
        voucher_hint.config(text="Add at least one line before saving.")
        return
    total_debit = round(sum(e[2] for e in entries), 2)
    total_credit = round(sum(e[3] for e in entries), 2)
    if total_debit != total_credit:
        if not messagebox.askyesno(
            "Unbalanced Voucher",
            f"Total Debit ({fmt(total_debit)}) does not equal Total Credit ({fmt(total_credit)}).\n"
            "Do you want to save anyway?",
        ):
            voucher_hint.config(text="Total Debit must equal Total Credit.")
            return
    if not is_valid_date(date_entry.get().strip()):
        voucher_hint.config(text="Date must be in YYYY-MM-DD format.")
        return
    try:
        backup_database("pre_save")
        with db.conn:
            if current_voucher_id is None:
                db.cur.execute(
                    "INSERT INTO vouchers (date, description) VALUES (?, ?)",
                    (date_entry.get(), desc_entry.get()),
                )
                vid = db.cur.lastrowid
            else:
                vid = current_voucher_id
                db.cur.execute(
                    "UPDATE vouchers SET date=?, description=? WHERE id=?",
                    (date_entry.get(), desc_entry.get(), vid),
                )
                db.cur.execute("DELETE FROM transactions WHERE voucher_id=?", (vid,))
            for e in entries:
                db.cur.execute(
                    "INSERT INTO transactions (voucher_id, account_id, debit, credit) VALUES (?, ?, ?, ?)",
                    (vid, e[0], e[2], e[3]),
                )
    except sqlite3.Error:
        logging.exception("Save voucher failed")
        voucher_hint.config(text="Failed to save voucher. Please try again.")
        return
    set_status("Voucher saved")
    voucher_hint.config(text="Voucher saved successfully.")
    refresh_dashboard()
    refresh_voucher_history()
    clear_voucher_form()

ttk.Button(form_voucher, text="Add Line", command=add_line).grid(row=2, column=3, padx=(6, 0))
top_save_btn = ttk.Button(form_voucher, text="Save Voucher", command=save_voucher, style="Primary.TButton")
top_save_btn.grid(row=0, column=3, padx=(6, 0))
voucher_actions = ttk.Frame(frame_voucher)
voucher_actions.pack(anchor="w", pady=10)
ttk.Button(voucher_actions, text="Remove Selected Line", command=remove_selected_line).grid(row=0, column=0, padx=(0, 8))
save_btn = ttk.Button(voucher_actions, text="Save Voucher", command=save_voucher, style="Primary.TButton")
save_btn.grid(row=0, column=1, padx=(0, 8))
ttk.Button(voucher_actions, text="New Voucher", command=clear_voucher_form).grid(row=0, column=2, padx=(0, 8))

# ================= REPORTS =================
frame_reports = ttk.Frame(content_area)
frames["Reports"] = frame_reports
ttk.Label(frame_reports, text="Financial Reports", style="Header.TLabel").pack(anchor="w")
ttk.Label(frame_reports, text="Quick insights and statements", style="SubHeader.TLabel").pack(anchor="w", pady=(2, 12))
current_report_title = tk.StringVar(value="Report")

# Dashboard cards
card_frame = ttk.Frame(frame_reports)
card_frame.pack(fill="x", pady=10)

cards = {}
for i, text in enumerate(["Total Assets", "Total Liabilities", "Net Profit"]):
    f = ttk.Frame(card_frame, style="Card.TFrame", padding=15)
    f.grid(row=0, column=i, padx=10, sticky="nsew")
    ttk.Label(f, text=text, style="Card.TLabel").pack()
    lbl = ttk.Label(f, text="0.00", style="Card.TLabel")
    lbl.pack()
    cards[text] = lbl

ttk.Separator(frame_reports, orient="horizontal").pack(fill="x", pady=8)

# Date filter
filter_frame = ttk.Frame(frame_reports)
filter_frame.pack(fill="x", pady=8)
ttk.Label(filter_frame, text="From:").pack(side="left")
date_from = ttk.Entry(filter_frame, width=12)
date_from.pack(side="left", padx=5)
date_from.insert(0, "2000-01-01")
ttk.Label(filter_frame, text="To:").pack(side="left")
date_to = ttk.Entry(filter_frame, width=12)
date_to.pack(side="left", padx=5)
date_to.insert(0, datetime.now().strftime("%Y-%m-%d"))
ttk.Button(filter_frame, text="Today", command=lambda: date_to.delete(0, tk.END) or date_to.insert(0, datetime.now().strftime("%Y-%m-%d"))).pack(side="left", padx=6)
ttk.Label(filter_frame, text="Account:").pack(side="left", padx=(12, 4))
report_account_cb = ttk.Combobox(filter_frame, state="readonly", width=28)
report_account_cb.pack(side="left")
ttk.Label(filter_frame, text="Cash Accts:").pack(side="left", padx=(12, 4))
cash_accounts_entry = ttk.Entry(filter_frame, width=28)
cash_accounts_entry.pack(side="left")
ttk.Button(filter_frame, text="Clear Filters", command=lambda: (date_from.delete(0, tk.END), date_from.insert(0, "2000-01-01"), date_to.delete(0, tk.END), date_to.insert(0, datetime.now().strftime("%Y-%m-%d")), report_account_cb.set(""), cash_accounts_entry.delete(0, tk.END))).pack(side="left", padx=6)

def set_report_view(columns, headings, numeric_cols=None):
    numeric_cols = set(numeric_cols or [])
    report_tree["columns"] = columns
    for col in columns:
        report_tree.heading(col, text="")
        report_tree.column(col, anchor="w")
    for col, hd in zip(columns, headings):
        anchor = "e" if col in numeric_cols else "w"
        report_tree.heading(col, text=hd, anchor=anchor)
        report_tree.column(col, anchor=anchor)

def get_report_dates():
    start = date_from.get().strip()
    end = date_to.get().strip()
    if not is_valid_date(start) or not is_valid_date(end):
        messagebox.showerror("Error", "Dates must be in YYYY-MM-DD format.")
        return None
    if start > end:
        messagebox.showerror("Error", "From date must be before To date.")
        return None
    return start, end

def refresh_dashboard():
    assets = fetch_one("""
        SELECT SUM(t.debit - t.credit) FROM transactions t
        JOIN accounts a ON t.account_id=a.id WHERE a.type='Asset'
    """)[0] or 0

    liabilities = fetch_one("""
        SELECT SUM(t.credit - t.debit) FROM transactions t
        JOIN accounts a ON t.account_id=a.id WHERE a.type='Liability'
    """)[0] or 0

    income = fetch_one("""
        SELECT SUM(t.credit - t.debit) FROM transactions t
        JOIN accounts a ON t.account_id=a.id WHERE a.type='Income'
    """)[0] or 0

    expense = fetch_one("""
        SELECT SUM(t.debit - t.credit) FROM transactions t
        JOIN accounts a ON t.account_id=a.id WHERE a.type='Expense'
    """)[0] or 0

    cards["Total Assets"].config(text=fmt(assets))
    cards["Total Liabilities"].config(text=fmt(liabilities))
    cards["Net Profit"].config(text=fmt(income - expense))

def generate_trial_balance():
    report_tree.delete(*report_tree.get_children())
    dates = get_report_dates()
    if not dates:
        return
    start, end = dates
    current_report_title.set("Trial Balance")
    set_report_view(["Account", "Debit", "Credit"], ["Account", "Debit", "Credit"], numeric_cols=["Debit", "Credit"])
    rows = fetch_all("""
        SELECT a.name, SUM(t.debit), SUM(t.credit)
        FROM transactions t JOIN accounts a ON t.account_id = a.id
        JOIN vouchers v ON t.voucher_id=v.id
        WHERE v.date BETWEEN ? AND ?
        GROUP BY a.name
        ORDER BY a.name
    """, (start, end))
    for i, r in enumerate(rows):
        insert_tree_row(report_tree, (r[0], fmt(r[1] or 0), fmt(r[2] or 0)), i)
    set_status("Trial Balance generated")

def generate_pl():
    report_tree.delete(*report_tree.get_children())
    dates = get_report_dates()
    if not dates:
        return
    start, end = dates
    current_report_title.set("Profit & Loss")
    set_report_view(["Account", "Debit", "Credit"], ["Account", "Debit", "Credit"], numeric_cols=["Debit", "Credit"])
    income = fetch_all("""
        SELECT a.name, SUM(t.credit - t.debit)
        FROM transactions t JOIN accounts a ON t.account_id=a.id
        JOIN vouchers v ON t.voucher_id=v.id
        WHERE a.type='Income' AND v.date BETWEEN ? AND ?
        GROUP BY a.name
    """, (start, end))
    expense = fetch_all("""
        SELECT a.name, SUM(t.debit - t.credit)
        FROM transactions t JOIN accounts a ON t.account_id=a.id
        JOIN vouchers v ON t.voucher_id=v.id
        WHERE a.type='Expense' AND v.date BETWEEN ? AND ?
        GROUP BY a.name
    """, (start, end))
    total_income = sum([r[1] or 0 for r in income])
    total_expense = sum([r[1] or 0 for r in expense])
    row_idx = 0
    for r in income:
        insert_tree_row(report_tree, (r[0], "0.00", fmt(r[1] or 0)), row_idx)
        row_idx += 1
    for r in expense:
        insert_tree_row(report_tree, (r[0], fmt(r[1] or 0), "0.00"), row_idx)
        row_idx += 1
    net = total_income - total_expense
    if net >= 0:
        insert_tree_row(report_tree, ("Net Profit", "0.00", fmt(net)), row_idx)
    else:
        insert_tree_row(report_tree, ("Net Loss", fmt(abs(net)), "0.00"), row_idx)
    set_status("Profit & Loss generated")

def generate_bs():
    report_tree.delete(*report_tree.get_children())
    dates = get_report_dates()
    if not dates:
        return
    start, end = dates
    current_report_title.set("Balance Sheet")
    set_report_view(["Account", "Debit", "Credit"], ["Account", "Debit", "Credit"], numeric_cols=["Debit", "Credit"])
    assets = fetch_all("""
        SELECT a.name, SUM(t.debit - t.credit)
        FROM transactions t JOIN accounts a ON t.account_id=a.id
        JOIN vouchers v ON t.voucher_id=v.id
        WHERE a.type='Asset' AND v.date BETWEEN ? AND ?
        GROUP BY a.name
    """, (start, end))
    liabilities = fetch_all("""
        SELECT a.name, SUM(t.credit - t.debit)
        FROM transactions t JOIN accounts a ON t.account_id=a.id
        JOIN vouchers v ON t.voucher_id=v.id
        WHERE a.type='Liability' AND v.date BETWEEN ? AND ?
        GROUP BY a.name
    """, (start, end))
    row_idx = 0
    for r in assets:
        insert_tree_row(report_tree, (r[0], fmt(r[1] or 0), "0.00"), row_idx)
        row_idx += 1
    for r in liabilities:
        insert_tree_row(report_tree, (r[0], "0.00", fmt(r[1] or 0)), row_idx)
        row_idx += 1
    set_status("Balance Sheet generated")

def generate_general_ledger():
    report_tree.delete(*report_tree.get_children())
    dates = get_report_dates()
    if not dates:
        return
    start, end = dates
    current_report_title.set("General Ledger")
    account_name = report_account_cb.get().strip()
    if not account_name:
        messagebox.showerror("Error", "Select an account for General Ledger.")
        return
    row = fetch_one("SELECT id, type FROM accounts WHERE name=?", (account_name,))
    if not row:
        messagebox.showerror("Error", "Selected account not found.")
        return
    account_id, account_type = row
    rows = fetch_all("""
        SELECT v.date, v.description, t.debit, t.credit
        FROM transactions t
        JOIN vouchers v ON t.voucher_id=v.id
        WHERE t.account_id = ? AND v.date BETWEEN ? AND ?
        ORDER BY v.date, v.id, t.id
    """, (account_id, start, end))
    set_report_view(
        ["Date", "Description", "Debit", "Credit", "Balance"],
        ["Date", "Description", "Debit", "Credit", "Balance"],
        numeric_cols=["Debit", "Credit", "Balance"],
    )
    balance = 0.0
    for i, r in enumerate(rows):
        debit = r[2] or 0
        credit = r[3] or 0
        delta = (debit - credit) if account_type in ("Asset", "Expense") else (credit - debit)
        balance += delta
        insert_tree_row(
            report_tree,
            (r[0], r[1] or "", fmt(debit), fmt(credit), fmt(balance)),
            i,
        )
    set_status("General Ledger generated")

def generate_cash_flow():
    report_tree.delete(*report_tree.get_children())
    dates = get_report_dates()
    if not dates:
        return
    start, end = dates
    current_report_title.set("Cash Flow")
    raw_cash = cash_accounts_entry.get().strip()
    cash_names = [n.strip() for n in raw_cash.split(",") if n.strip()]
    if not cash_names:
        cash_names = [
            r[0] for r in fetch_all(
                "SELECT name FROM accounts WHERE LOWER(name) LIKE '%cash%' OR LOWER(name) LIKE '%bank%'"
            )
        ]
    if not cash_names:
        messagebox.showerror("Error", "No cash accounts found. Enter names in Cash Accts.")
        return
    set_report_view(
        ["Account", "Inflow", "Outflow"],
        ["Account", "Inflow", "Outflow"],
        numeric_cols=["Inflow", "Outflow"],
    )
    placeholders = ",".join(["?"] * len(cash_names))
    rows = fetch_all(f"""
        SELECT a.name, SUM(t.debit), SUM(t.credit)
        FROM transactions t
        JOIN accounts a ON t.account_id = a.id
        JOIN vouchers v ON t.voucher_id=v.id
        WHERE a.name IN ({placeholders})
          AND v.date BETWEEN ? AND ?
        GROUP BY a.name
        ORDER BY a.name
    """, (*cash_names, start, end))
    total_in = 0
    total_out = 0
    for i, r in enumerate(rows):
        inflow = r[1] or 0
        outflow = r[2] or 0
        total_in += inflow
        total_out += outflow
        insert_tree_row(report_tree, (r[0], fmt(inflow), fmt(outflow)), i)
    insert_tree_row(report_tree, ("Net Cash", fmt(total_in), fmt(total_out)), len(rows))
    set_status("Cash Flow generated")

def export_csv():
    path = filedialog.asksaveasfilename(defaultextension=".csv")
    if not path:
        return
    with open(path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(list(report_tree["columns"]))
        for row in report_tree.get_children():
            writer.writerow(report_tree.item(row)["values"])
    set_status("Exported to Excel (CSV)")

def get_report_table_data():
    columns = list(report_tree["columns"])
    if not columns:
        return None, None, None
    headings = [report_tree.heading(col)["text"] for col in columns]
    rows = [report_tree.item(row)["values"] for row in report_tree.get_children()]
    return columns, headings, rows

def build_report_pdf(path):
    _, headings, rows = get_report_table_data()
    if not headings:
        messagebox.showerror("Error", "No report data to export.")
        return False
    data = [headings] + rows
    doc = SimpleDocTemplate(path, pagesize=landscape(LETTER), title=current_report_title.get())
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph(current_report_title.get(), styles["Title"]))
    dates = get_report_dates()
    if dates:
        elements.append(Paragraph(f"From {dates[0]} To {dates[1]}", styles["Normal"]))
    elements.append(Spacer(1, 12))
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b1220")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(table)
    doc.build(elements)
    return True

def export_pdf():
    path = filedialog.asksaveasfilename(defaultextension=".pdf")
    if not path:
        return
    try:
        if build_report_pdf(path):
            set_status("PDF exported")
    except Exception:
        logging.exception("PDF export failed")
        messagebox.showerror("Error", "Failed to export PDF. See rias.log.")

def print_report():
    try:
        tmp_path = os.path.join(tempfile.gettempdir(), "rias_report_print.pdf")
        if not build_report_pdf(tmp_path):
            return
        try:
            os.startfile(tmp_path, "print")
            set_status("Print job sent")
        except Exception:
            os.startfile(tmp_path, "open")
            set_status("Opened PDF for printing")
    except Exception:
        logging.exception("Print failed")
        messagebox.showerror("Error", "Failed to print report. See rias.log.")

actions = ttk.Frame(frame_reports)
actions.pack(side="bottom", fill="x", pady=10)
ttk.Button(actions, text="Trial Balance", command=generate_trial_balance).grid(row=0, column=0, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="Profit & Loss", command=generate_pl).grid(row=0, column=1, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="Balance Sheet", command=generate_bs).grid(row=0, column=2, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="General Ledger", command=generate_general_ledger).grid(row=0, column=3, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="Cash Flow", command=generate_cash_flow).grid(row=0, column=4, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="Export CSV", command=export_csv).grid(row=1, column=0, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="Refresh Dashboard", command=refresh_dashboard, style="Primary.TButton").grid(row=1, column=1, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="Export PDF", command=export_pdf).grid(row=1, column=2, padx=5, pady=4, sticky="w")
ttk.Button(actions, text="Print", command=print_report).grid(row=1, column=3, padx=5, pady=4, sticky="w")

# Reports Treeview
report_tree = create_treeview(
    frame_reports,
    ["Account", "Debit", "Credit"],
    ["Account", "Debit", "Credit"],
    height=15,
    numeric_cols=["Debit", "Credit"],
    pack=False,
)
report_tree.pack(fill="both", expand=True, pady=12)

# ================= VOUCHER HISTORY =================
frame_history = ttk.Frame(content_area)
frames["Voucher History"] = frame_history
ttk.Label(frame_history, text="Voucher History", style="Header.TLabel").pack(anchor="w")
ttk.Label(frame_history, text="Review saved vouchers and their lines", style="SubHeader.TLabel").pack(anchor="w", pady=(2, 12))

history_filter = ttk.Frame(frame_history)
history_filter.pack(fill="x", pady=6)
ttk.Label(history_filter, text="From:").pack(side="left")
hist_from = ttk.Entry(history_filter, width=12)
hist_from.pack(side="left", padx=5)
hist_from.insert(0, "2000-01-01")
ttk.Label(history_filter, text="To:").pack(side="left")
hist_to = ttk.Entry(history_filter, width=12)
hist_to.pack(side="left", padx=5)
hist_to.insert(0, datetime.now().strftime("%Y-%m-%d"))
ttk.Button(history_filter, text="Today", command=lambda: hist_to.delete(0, tk.END) or hist_to.insert(0, datetime.now().strftime("%Y-%m-%d"))).pack(side="left", padx=6)

ttk.Label(history_filter, text="Search").pack(side="left", padx=(12, 4))
hist_search = ttk.Entry(history_filter, width=24)
hist_search.pack(side="left")

history_actions = ttk.Frame(frame_history)
history_actions.pack(fill="x", pady=6)

vouchers_tree = create_treeview(
    frame_history,
    ["ID", "Date", "Description", "Debit", "Credit"],
    ["Voucher ID", "Date", "Description", "Total Debit", "Total Credit"],
    height=10,
    numeric_cols=["Debit", "Credit"],
)

ttk.Label(frame_history, text="Voucher Lines", style="SubHeader.TLabel").pack(anchor="w", pady=(8, 4))
lines_tree = create_treeview(
    frame_history,
    ["Account", "Debit", "Credit"],
    ["Account", "Debit", "Credit"],
    height=10,
    numeric_cols=["Debit", "Credit"],
)

def refresh_voucher_history():
    vouchers_tree.delete(*vouchers_tree.get_children())
    lines_tree.delete(*lines_tree.get_children())
    if not is_valid_date(hist_from.get().strip()) or not is_valid_date(hist_to.get().strip()):
        messagebox.showerror("Error", "Dates must be in YYYY-MM-DD format.")
        return
    search_text = hist_search.get().strip()
    rows = fetch_all(
        """
        SELECT v.id, v.date, v.description, SUM(t.debit) AS total_debit, SUM(t.credit) AS total_credit
        FROM vouchers v
        JOIN transactions t ON v.id = t.voucher_id
        WHERE v.date BETWEEN ? AND ?
          AND (v.description LIKE ? OR v.id LIKE ?)
        GROUP BY v.id, v.date, v.description
        ORDER BY v.date DESC, v.id DESC
        """,
        (hist_from.get(), hist_to.get(), f"%{search_text}%", f"%{search_text}%"),
    )
    for i, r in enumerate(rows):
        insert_tree_row(vouchers_tree, (r[0], r[1], r[2] or "", fmt(r[3] or 0), fmt(r[4] or 0)), i)
    set_status("Voucher history refreshed")

def on_voucher_select(event):
    selection = vouchers_tree.selection()
    if not selection:
        return
    values = vouchers_tree.item(selection[0], "values")
    if not values:
        return
    voucher_id = values[0]
    lines_tree.delete(*lines_tree.get_children())
    rows = fetch_all(
        """
        SELECT a.name, t.debit, t.credit
        FROM transactions t
        JOIN accounts a ON t.account_id = a.id
        WHERE t.voucher_id = ?
        """,
        (voucher_id,),
    )
    for i, r in enumerate(rows):
        insert_tree_row(lines_tree, (r[0], fmt(r[1] or 0), fmt(r[2] or 0)), i)

vouchers_tree.bind("<<TreeviewSelect>>", on_voucher_select)
vouchers_tree.bind("<Double-1>", lambda e: load_voucher_for_edit())
vouchers_tree.bind("<Delete>", lambda e: delete_selected_voucher())
def load_voucher_for_edit():
    global current_voucher_id
    selection = vouchers_tree.selection()
    if not selection:
        set_status("Select a voucher to edit")
        return
    values = vouchers_tree.item(selection[0], "values")
    if not values:
        return
    voucher_id = values[0]
    row = fetch_one("SELECT date, description FROM vouchers WHERE id=?", (voucher_id,))
    if not row:
        set_status("Voucher not found")
        return
    date_entry.delete(0, tk.END)
    date_entry.insert(0, row[0])
    desc_entry.delete(0, tk.END)
    desc_entry.insert(0, row[1] or "")
    entries.clear()
    tx_rows = fetch_all("""
        SELECT a.id, a.name, t.debit, t.credit
        FROM transactions t
        JOIN accounts a ON t.account_id = a.id
        WHERE t.voucher_id = ?
        ORDER BY t.id
    """, (voucher_id,))
    for r in tx_rows:
        entries.append((r[0], r[1], r[2] or 0, r[3] or 0))
    rebuild_voucher_tree()
    update_voucher_totals()
    current_voucher_id = voucher_id
    set_voucher_mode(True)
    show_frame("Voucher Entry", frame_voucher)
    set_status(f"Editing voucher {voucher_id}")

def delete_selected_voucher():
    selection = vouchers_tree.selection()
    if not selection:
        set_status("Select a voucher to delete")
        return
    values = vouchers_tree.item(selection[0], "values")
    if not values:
        return
    voucher_id = values[0]
    if not messagebox.askyesno("Confirm", f"Delete voucher {voucher_id}?"):
        return
    try:
        execute_query("DELETE FROM vouchers WHERE id=?", (voucher_id,))
    except sqlite3.Error:
        set_status("Failed to delete voucher")
        return
    refresh_voucher_history()
    refresh_dashboard()
    set_status("Voucher deleted")

ttk.Button(history_actions, text="Refresh", command=refresh_voucher_history).pack(side="left", padx=5)
ttk.Button(history_actions, text="Edit Voucher", command=load_voucher_for_edit).pack(side="left", padx=5)
ttk.Button(history_actions, text="Delete Voucher", command=delete_selected_voucher).pack(side="left", padx=5)
ttk.Button(history_actions, text="Clear Search", command=lambda: (hist_search.delete(0, tk.END), refresh_voucher_history())).pack(side="left", padx=5)
hist_search.bind("<KeyRelease>", lambda e: refresh_voucher_history())

# ================= NAVIGATION =================
ttk.Label(sidebar, text="RIAS", style="Sidebar.TLabel").pack(anchor="w")
ttk.Label(sidebar, text="Accounting Suite", background=PRIMARY, foreground="#cbd5f5", font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 18))
nav_button("Accounts", frame_accounts)
nav_button("Voucher Entry", frame_voucher)
nav_button("Reports", frame_reports)
nav_button("Voucher History", frame_history)

# ================= DROPDOWNS =================
def refresh_dropdowns():
    names = [r[0] for r in fetch_all("SELECT name FROM accounts")]
    current_acc = account_cb.get()
    current_report = report_account_cb.get()
    account_cb["values"] = names
    report_account_cb["values"] = names
    if current_acc in names:
        account_cb.set(current_acc)
    if current_report in names:
        report_account_cb.set(current_report)

refresh_accounts()
refresh_dashboard()
refresh_voucher_history()
show_frame("Accounts", frame_accounts)
backup_database("startup")
check_db_integrity()

def focus_search():
    if current_frame_name == "Accounts":
        search_acc.focus_set()
    elif current_frame_name == "Voucher History":
        hist_search.focus_set()

def on_ctrl_s(event=None):
    if current_frame_name == "Voucher Entry":
        save_voucher()
        return "break"

def on_ctrl_l(event=None):
    if current_frame_name == "Voucher Entry":
        add_line()
        return "break"

def on_ctrl_n(event=None):
    if current_frame_name == "Voucher Entry":
        clear_voucher_form()
        return "break"

def on_ctrl_f(event=None):
    focus_search()
    return "break"

def on_enter(event=None):
    focus = root.focus_get()
    if current_frame_name == "Accounts":
        if focus in (acc_name, acc_type):
            add_account()
            return "break"
        if focus == search_acc:
            refresh_accounts(search_acc.get().strip())
            return "break"
    if current_frame_name == "Voucher Entry":
        if focus in (account_cb, debit_entry, credit_entry):
            add_line()
            return "break"
    if current_frame_name == "Reports":
        if focus in (date_from, date_to):
            generate_trial_balance()
            return "break"
    if current_frame_name == "Voucher History":
        if focus == hist_search:
            refresh_voucher_history()
            return "break"

def run_if_frame(frame_name, action):
    if current_frame_name == frame_name:
        action()
        return "break"

root.bind("<Control-s>", on_ctrl_s)
root.bind("<Control-l>", on_ctrl_l)
root.bind("<Control-f>", on_ctrl_f)
root.bind("<Control-n>", on_ctrl_n)
root.bind("<Return>", on_enter)
root.bind("<Control-1>", lambda e: (show_frame("Accounts", frame_accounts), "break"))
root.bind("<Control-2>", lambda e: (show_frame("Voucher Entry", frame_voucher), "break"))
root.bind("<Control-3>", lambda e: (show_frame("Reports", frame_reports), "break"))
root.bind("<Control-4>", lambda e: (show_frame("Voucher History", frame_history), "break"))
root.bind("<Control-u>", lambda e: run_if_frame("Accounts", update_account))
root.bind("<Control-d>", lambda e: run_if_frame("Accounts", delete_account))
root.bind("<Control-k>", lambda e: run_if_frame("Accounts", clear_account_form))
root.bind("<Control-r>", lambda e: run_if_frame("Reports", refresh_dashboard) or run_if_frame("Voucher History", refresh_voucher_history))
root.bind("<Control-e>", lambda e: run_if_frame("Reports", export_csv) or run_if_frame("Voucher History", load_voucher_for_edit))
root.bind("<Control-p>", lambda e: run_if_frame("Reports", print_report))
root.bind("<Control-Shift-E>", lambda e: run_if_frame("Reports", export_pdf))
root.bind("<Alt-t>", lambda e: run_if_frame("Reports", generate_trial_balance))
root.bind("<Alt-p>", lambda e: run_if_frame("Reports", generate_pl))
root.bind("<Alt-b>", lambda e: run_if_frame("Reports", generate_bs))
root.bind("<Alt-g>", lambda e: run_if_frame("Reports", generate_general_ledger))
root.bind("<Alt-c>", lambda e: run_if_frame("Reports", generate_cash_flow))

root.mainloop()
db.close()
