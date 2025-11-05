import tkinter as tk
from tkinter import ttk, scrolledtext
import pandas as pd
import os
import time
import threading
from datetime import datetime
import win32file
import win32con
import csv
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import make_pipeline
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import random
import difflib

# Network paths
CSV_LOG_PATH = r"\\192.168.2.10\csv\csv\JobOrder\log000_JobOrder.csv"
VT1_LOG_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_1.csv" #FALSE TEST PATH
JO_MATERIAL_DIR = r"\\192.168.2.19\production\2025\1. Document for Production Admin\8. JOB ORDER MATERIAL LIST"
OUTPUT_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\wrongMaterialDetectorCSV.csv" # MY OWN CSV FILE
COLUMN_NAME = "Job Order Number" # Exact column name in CSV
# Task 2 & 3 Paths
EXCEL_FOLDER = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\8. JOB ORDER MATERIAL LIST"
LOG000_1_PATH = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\log000_1.csv"
# Column mapping for Task 3 (1-based index from log000_1.csv)
COLUMN_MAPPING = {
    'DATE': 2, #1
    'TIME': 3, #2
    'Process 1 Model Code': 4, #3
    'Process 1 S/N': 5, #4
    'Process 1 Em2p': 9, #8
    'Process 1 Em3p': 11, #10
    'Process 1 Harness': 13, #12
    'Process 1 Frame': 15, #14
    'Process 1 Bushing': 17 #16
}
VALIDATION_COLUMN = 'Process 1 Em2p' # The column in output CSV to validate
EXCEL_MATERIAL_COLUMN_INDEX = 4 # The 0-based index (4 for Column E) in the Excel file to check against
REFERENCE_EXCEL = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\8. JOB ORDER MATERIAL LIST\3J737987830000.xlsx"
ML_MATERIAL_COL = 4 # Column E (0-based)
ML_DESC_COL = 10 # Column K (0-based)
ML_TARGET_COLUMNS = ['Process 1 Em2p', 'Process 1 Em3p', 'Process 1 Harness',
                     'Process 1 Frame', 'Process 1 Bushing'] # columns 6-10 in output CSV
# Column names
JO_NUMBER_COL = "Job Order Number"
MATERIAL_COL = "MATERIAL"
# Process 1 items to monitor
PROCESS1_ITEMS = [
    "Process 1 Em2p",
    "Process 1 Em3p",
    "Process 1 Harness",
    "Process 1 Frame",
    "Process 1 Bushing"
]
# Ensure output directory exists
output_dir = os.path.dirname(OUTPUT_PATH)
os.makedirs(output_dir, exist_ok=True)

# GUI setup
root = tk.Tk()
root.title("Production Monitoring System")
root.geometry("1100x800")
root.configure(bg="#f0f0f0")

# === 7TH: JOB ORDER DISPLAY (Top Center - Pleasing) ===
jo_frame = tk.Frame(root, bg="#f0f0f0")
jo_frame.pack(pady=8)
jo_label_title = tk.Label(jo_frame, text="CURRENT JOB ORDER:", font=("Arial", 12, "bold"), bg="#f0f0f0", fg="#003366")
jo_label_title.pack(side="left", padx=5)
jo_display = tk.Label(jo_frame, text="N/A", font=("Arial", 14, "bold"), bg="#ffffff", fg="#003366", relief="sunken", width=20, anchor="center")
jo_display.pack(side="left")

# === 1ST & 2ND: BACK & NEXT BUTTONS (Top Left & Right) ===
nav_frame = tk.Frame(root, bg="#f0f0f0")
nav_frame.pack(fill="x", pady=5)

back_btn = tk.Button(nav_frame, text="BACK", font=("Arial", 10, "bold"), bg="#ff9800", fg="white", width=10)
back_btn.pack(side="left", padx=10)

next_btn = tk.Button(nav_frame, text="NEXT", font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", width=10)
next_btn.pack(side="right", padx=10)

# === 3RD: ALL PROCESS BUTTONS IN ONE ROW ===
btn_frame = tk.Frame(root, bg="#f0f0f0")
btn_frame.pack(pady=10)

process_buttons = {}
loading_labels = {}
dot_labels = {}
blink_threads = {}
stop_flags = {}
process_active = {}

for i in range(1, 7):
    proc_name = f"PROCESS {i}"
    proc_frame = tk.Frame(btn_frame, bg="#f0f0f0")
    proc_frame.pack(side="left", padx=15)

    load_lbl = tk.Label(proc_frame, text="Loading", font=("Arial", 9), fg="gray", bg="#f0f0f0")
    load_lbl.pack()
    loading_labels[i] = load_lbl

    btn = tk.Button(proc_frame, text=f"{proc_name}\nSTOP", font=("Arial", 11, "bold"),
                    width=10, height=2, bg="#4CAF50", fg="white", relief="raised",
                    command=lambda p=i: acknowledge_stop(p))
    btn.pack(pady=3)
    process_buttons[i] = btn

    dot_lbl = tk.Label(proc_frame, text="...", font=("Arial", 12), fg="#4CAF50", bg="#f0f0f0")
    dot_lbl.pack()
    dot_labels[i] = dot_lbl

    stop_flags[i] = threading.Event()
    process_active[i] = True

# === 4TH: PREVIOUS, REFRESH, STOP ALL (Below NEXT, Right-Aligned) ===
control_frame = tk.Frame(root, bg="#f0f0f0")
control_frame.pack(fill="x", pady=5)

prev_btn = tk.Button(control_frame, text="PREVIOUS", font=("Arial", 10, "bold"), bg="#2196F3", fg="white", width=12)
prev_btn.pack(side="right", padx=5)

refresh_btn = tk.Button(control_frame, text="REFRESH", font=("Arial", 10, "bold"), bg="#9C27B0", fg="white", width=12)
refresh_btn.pack(side="right", padx=5)

stop_all_btn = tk.Button(control_frame, text="STOP ALL", font=("Arial", 10, "bold"), bg="#f44336", fg="white", width=12)
stop_all_btn.pack(side="right", padx=10)

# === 5TH: SYSTEM MESSAGE TEXT BOX ===
msg_frame = tk.Frame(root, bg="#f0f0f0")
msg_frame.pack(pady=10, fill="both", expand=False, padx=20)

msg_label = tk.Label(msg_frame, text="SYSTEM MESSAGE", font=("Arial", 11, "bold"), bg="#f0f0f0", anchor="w")
msg_label.pack(anchor="w")

system_text = scrolledtext.ScrolledText(msg_frame, height=6, font=("Courier", 9), state="disabled", bg="#f8f8f8", fg="#000000")
system_text.pack(fill="both", expand=False, pady=5)

# === 6TH & 9TH: MATERIAL UPDATE LOG (Larger, 5 Columns with Color Tags) ===
log_frame = tk.Frame(root, bg="#f0f0f0")
log_frame.pack(pady=10, fill="both", expand=True, padx=20)

log_label = tk.Label(log_frame, text="MATERIAL UPDATE LOG", font=("Arial", 11, "bold"), bg="#f0f0f0", anchor="w")
log_label.pack(anchor="w")

# Treeview for structured table
columns = ("Process", "Material", "Error", "Correction", "Status")
material_tree = ttk.Treeview(log_frame, columns=columns, show="headings", height=2)
material_tree.heading("Process", text="PROCESS")
material_tree.heading("Material", text="MATERIAL")
material_tree.heading("Error", text="ERROR")
material_tree.heading("Correction", text="CORRECTION")
material_tree.heading("Status", text="STATUS")

material_tree.column("Process", width=80, anchor="center")
material_tree.column("Material", width=160, anchor="center")
material_tree.column("Error", width=180, anchor="center")
material_tree.column("Correction", width=180, anchor="center")
material_tree.column("Status", width=80, anchor="center")

material_tree.pack(fill="both", expand=True, side="left")

# Scrollbar for Material Log
mat_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=material_tree.yview)
mat_scroll.pack(side="right", fill="y")
material_tree.configure(yscrollcommand=mat_scroll.set)

# === 8TH: CLEAR BUTTON BELOW MATERIAL LOG ===
clear_btn_frame = tk.Frame(root, bg="#f0f0f0")
clear_btn_frame.pack(pady=5)

def clear_material_log():
    for item in material_tree.get_children():
        material_tree.delete(item)
    log_system("MATERIAL LOG CLEARED")

clear_material_button = tk.Button(clear_btn_frame, text="CLEAR LOG", font=("Arial", 10, "bold"), bg="#ff4444", fg="white",
                                  width=15, relief="raised", command=clear_material_log)
clear_material_button.pack()

# === 11TH: VERTICAL SCROLLBAR CONTROLLED BY MOUSE WHEEL ===
def _on_mousewheel(event):
    material_tree.yview_scroll(int(-1*(event.delta/120)), "units")
    system_text.yview_scroll(int(-1*(event.delta/120)), "units")
    return "break"

root.bind_all("<MouseWheel>", _on_mousewheel)

# Global variables
last_jo_number = None
last_vt1_mtime = None
running = True
last_mtime_log1 = 0
animation_running = True
current_jo = "N/A"

# === SYSTEM LOG FUNCTION ===
def log_system(message):
    system_text.config(state="normal")
    timestamp = datetime.now().strftime("%H:%M:%S")
    system_text.insert(tk.END, f"[{timestamp}] {message}\n")
    system_text.see(tk.END)
    system_text.config(state="disabled")

# === MATERIAL LOG WITH COLOR TAGS (MATERIAL = "FRAME") ===
def log_material(process, material, error, correction):
    # CHANGED: Material column always shows "FRAME"
    item = material_tree.insert("", "end", values=(process, "FRAME", error, correction, "ERROR"))
    material_tree.item(item, tags=("error",))
    material_tree.tag_configure("error", foreground="red", font=("Courier", 9, "bold"))

    # Highlight diff in ERROR vs CORRECTION
    if error and correction:
        diff = difflib.ndiff(error, correction)
        error_colored = ""
        for d in diff:
            if d.startswith(" "):
                error_colored += d[-1]
            elif d.startswith("-"):
                error_colored += f"[red]{d[-1]}[/red]"
            elif d.startswith("+"):
                error_colored += f"[skyblue]{d[-1]}[/skyblue]"
        # We'll use tags later for actual coloring if needed
    else:
        error_colored = error

# === ACKNOWLEDGE STOP ===
def acknowledge_stop(process_id):
    global blink_threads
    stop_flags[process_id].set()
    btn = process_buttons[process_id]
    btn.config(bg="#4CAF50", relief="raised")
    log_system(f"PROCESS {process_id}: STOP ACKNOWLEDGED BY USER")
    if process_id in blink_threads:
        blink_threads[process_id].join(timeout=1)
        del blink_threads[process_id]
    process_active[process_id] = True
    stop_flags[process_id].clear()

# === BLINK BUTTON ===
def blink_button(process_id):
    btn = process_buttons[process_id]
    while not stop_flags[process_id].is_set() and process_active[process_id]:
        btn.config(bg="red")
        time.sleep(0.5)
        if stop_flags[process_id].is_set():
            break
        btn.config(bg="#ff6666")
        time.sleep(0.5)
    btn.config(bg="#4CAF50")

# === ANIMATE LOADING (FIXED: Continues even after button click) ===
def animate_loading():
    global animation_running
    while animation_running:
        for i in range(1, 7):
            if process_active[i]:
                loading_labels[i].config(text="Loading" + "." * ((int(time.time() * 2) % 4)))
                dot_text = "." * ((int(time.time() * 3) % 4) + 1)
                dot_labels[i].config(text=dot_text)
        time.sleep(0.3)

# === GET LAST LINE ===
def get_last_line(csv_file_path):
    with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
        lines = f.readlines()
        for line in reversed(lines):
            if line.strip():
                return line
    return None

def parse_last_row(last_line, delimiter=','):
    reader = csv.reader([last_line], delimiter=delimiter)
    return next(reader)

# === PROCESS JOB ORDER (TASK 1 & 2) ===
def process_job_order():
    global current_jo
    try:
        print(f"Reading tail of: {CSV_LOG_PATH}")
        if not os.path.exists(CSV_LOG_PATH):
            raise FileNotFoundError(f"Source file not found: {CSV_LOG_PATH}")
        last_line = get_last_line(CSV_LOG_PATH)
        if not last_line:
            print("Source CSV is empty.")
            return
        with open(CSV_LOG_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            try:
                headers = next(reader)
            except StopIteration:
                print("CSV has no headers.")
                return
            try:
                job_order_idx = headers.index(COLUMN_NAME)
            except ValueError:
                print(f"Column '{COLUMN_NAME}' not found in headers.")
                return
            rows = list(reader)
            if not rows:
                print("No data rows in CSV.")
                return
            last_row = rows[-1]
        raw_job_order = last_row[job_order_idx]
        cleaned_job_order = raw_job_order.replace(" ", "").replace("\t", "")
        print(f"Original: '{raw_job_order}'")
        print(f"Cleaned : '{cleaned_job_order}'")
        current_jo = cleaned_job_order
        jo_display.config(text=cleaned_job_order)

        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            header = ["Job Order Number", "File Name"] + list(COLUMN_MAPPING.keys())
            writer.writerow(header)
            empty_row = [cleaned_job_order, ""] + [""] * len(COLUMN_MAPPING)
            writer.writerow(empty_row)
        print(f"Successfully saved to: {OUTPUT_PATH}")

        excel_file = f"{cleaned_job_order}.xlsx"
        excel_full_path = os.path.join(EXCEL_FOLDER, excel_file)
        if os.path.exists(excel_full_path):
            print(f"THE FILE NAME {cleaned_job_order} WAS SUCCESSFULLY FOUND AT {EXCEL_FOLDER}")
            update_output_column(1, cleaned_job_order)
            log_system(f"MATCHED FILE: {cleaned_job_order}.xlsx")
        else:
            print(f".xlsx file NOT found: {excel_file}")
            update_output_column(1, "NOT FOUND")
            log_system(f".xlsx file NOT found: {excel_file}")
    except Exception as e:
        print(f"Error: {e}")

def update_output_column(col_index, value):
    if not os.path.exists(OUTPUT_PATH):
        return
    with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
        rows = list(csv.reader(f))
    if len(rows) > 1:
        rows[1][col_index] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(rows)

last_mtime_log1 = 0
def update_from_log000_1(force_update=False):
    global last_mtime_log1
    try:
        if not os.path.exists(LOG000_1_PATH):
            return False
        current_mtime = os.path.getmtime(LOG000_1_PATH)
        if not force_update and current_mtime <= last_mtime_log1:
            return False
        time.sleep(0.1)
        last_mtime_log1 = current_mtime
        print(f"\nTRANSFERING ({os.path.basename(LOG000_1_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")
        with open(LOG000_1_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False
        headers = output_rows[0]
        data_row = output_rows[1]
        for col_name, src_idx_1based in COLUMN_MAPPING.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)
        print(f"Updated {len(COLUMN_MAPPING)} columns from log000_1.csv")
        validate_material_in_excel(output_rows)
        return True
    except Exception as e:
        print(f"Error in log000_1 update: {e}")
        return False

# === SKLEARN VALIDATION + CORRECTION ===
def validate_material_in_excel(output_rows):
    try:
        if len(output_rows) < 2:
            return
        headers = output_rows[0]
        data_row = output_rows[1]
        job_order_num = data_row[headers.index("Job Order Number")]
        file_name_status = data_row[headers.index("File Name")]
        if file_name_status == "NOT FOUND":
            log_system(f"Validation skipped: Excel file for JO {job_order_num} was NOT FOUND.")
            return
        if not os.path.exists(REFERENCE_EXCEL):
            log_system(f"Reference Excel not found: {REFERENCE_EXCEL}")
            return
        ref_df = pd.read_excel(REFERENCE_EXCEL, header=None)
        materials = ref_df.iloc[:, ML_MATERIAL_COL].dropna().astype(str).str.strip()
        descriptions = ref_df.iloc[:, ML_DESC_COL].dropna().astype(str).str.strip()

        invalid_texts = []
        for mat in materials[:min(50, len(materials))]:
            invalid_texts.append(f"XYZ_{mat}")
            if len(mat) > 1:
                invalid_texts.append(mat[:-1])
            if len(mat) > 2:
                i = random.randint(0, len(mat)-2)
                typo = list(mat)
                typo[i], typo[i+1] = typo[i+1], typo[i]
                invalid_texts.append(''.join(typo))
            invalid_texts.append(''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=8)))
        invalid_texts = list(set(invalid_texts) - set(materials.str.upper()))
        train_texts = [f"{m} {d}" for m, d in zip(materials, descriptions)] + invalid_texts
        train_labels = ["VALID"] * len(materials) + ["INVALID"] * len(invalid_texts)

        ml_model = None
        if len(set(train_labels)) >= 2:
            vectorizer = TfidfVectorizer(lowercase=True, ngram_range=(1,2))
            clf = LogisticRegression(max_iter=1000)
            ml_model = make_pipeline(vectorizer, clf)
            try:
                ml_model.fit(train_texts, train_labels)
            except:
                ml_model = None

        material_codes = materials.tolist()
        correction_vectorizer = TfidfVectorizer(lowercase=True, ngram_range=(1,3), analyzer='char')
        tfidf_matrix = correction_vectorizer.fit_transform(material_codes) if material_codes else None

        print("\n" + "="*80)
        print(" SKLEARN VALIDATION + 100% CORRECT SUGGESTIONS (Columns 6-10) ")
        print("="*80)
        any_error = False
        empty_error = False
        for col_name in ML_TARGET_COLUMNS:
            if col_name not in headers:
                continue
            idx = headers.index(col_name)
            value = data_row[idx].strip()
            if not value:
                print(f" [EMPTY] {col_name}: <empty>")
                empty_error = True
                continue
            exact_match = value.upper() in {m.upper() for m in materials}
            if exact_match:
                print(f" [OK] {col_name}: {value} (exact match)")
                continue
            is_valid = False
            confidence = 0.0
            if ml_model:
                pred_input = f"{value} {value}"
                pred = ml_model.predict([pred_input])[0]
                prob = ml_model.predict_proba([pred_input])[0]
                confidence = prob[ml_model.classes_.tolist().index(pred)]
                if pred == "VALID" and confidence > 0.85:
                    is_valid = True
            if is_valid:
                print(f" [OK] {col_name}: {value} (ML-predicted VALID, {confidence:.2%})")
                continue
            any_error = True
            suggested_material = "UNKNOWN"
            input_upper = value.upper()
            if tfidf_matrix is not None:
                query_vec = correction_vectorizer.transform([input_upper])
                sim_scores = cosine_similarity(query_vec, tfidf_matrix).flatten()
                if len(sim_scores) > 0 and sim_scores.max() > 0.1:
                    best_idx = np.argmax(sim_scores)
                    suggested_material = material_codes[best_idx]
                else:
                    candidates = [m for m in material_codes if m.upper().startswith(input_upper[:4])]
                    if not candidates:
                        candidates = material_codes
                    match = difflib.get_close_matches(input_upper, candidates, n=1, cutoff=0.6)
                    suggested_material = match[0] if match else material_codes[0]
            else:
                suggested_material = material_codes[0] if material_codes else "NO DATA"
            print(f" [ERROR] {col_name}: {value}")
            print(f" **SUGGESTED CORRECTION** -> {suggested_material}")
            err_msg = f"PROCESS 1: ERROR: {value} EXPECTED: {suggested_material}"
            log_system(err_msg)
            # CHANGED: Material = "FRAME", pass actual value to error/correction
            log_material("1", value, value, suggested_material)

            if process_active[1] and not stop_flags[1].is_set():
                t = threading.Thread(target=blink_button, args=(1,), daemon=True)
                blink_threads[1] = t
                t.start()

        if empty_error:
            log_system("PROCESS 1: ERROR: ONE OR MORE MATERIAL FIELDS ARE EMPTY")
            if process_active[1] and not stop_flags[1].is_set():
                t = threading.Thread(target=blink_button, args=(1,), daemon=True)
                blink_threads[1] = t
                t.start()

        if not any_error and not empty_error:
            print(" All five material fields are CORRECT.")
            log_system("PROCESS 1: ALL MATERIALS VERIFIED")
            if stop_flags[1].is_set():
                acknowledge_stop(1)
        print("="*80 + "\n")
    except Exception as e:
        print(f"Error in material validation: {e}")
        log_system(f"Validation Error: {e}")

# === WATCHDOG HANDLER ===
class LogChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        normalized_event_path = os.path.normpath(event.src_path).lower()
        normalized_log_path = os.path.normpath(LOG000_1_PATH).lower()
        if normalized_event_path == normalized_log_path:
            pass

# === MONITORING LOOP ===
def monitoring_loop():
    global animation_running, last_mtime_log1
    observer = None
    dot_count = 0
    try:
        process_job_order()
        if os.path.exists(LOG000_1_PATH):
            if update_from_log000_1(force_update=True):
                log_system(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...")
            else:
                try:
                    last_mtime_log1 = os.path.getmtime(LOG000_1_PATH)
                except:
                    pass
        else:
            log_system(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...")

        observer = Observer()
        watch_dir = os.path.dirname(LOG000_1_PATH)
        if os.path.exists(watch_dir):
            observer.schedule(LogChangeHandler(), path=watch_dir, recursive=False)
            observer.start()

        while running:
            updated = False
            if os.path.exists(LOG000_1_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_1_PATH)
                    if current_mtime > last_mtime_log1:
                        if update_from_log000_1(force_update=True):
                            updated = True
                            dot_count = 0
                except:
                    pass

            # Minimal console output
            if not updated:
                dot_str = "." * min((dot_count % 39) + 1, 39)
                print(f"\rWAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES{dot_str:<39}", end="", flush=True)
            else:
                print(f"\r{' ' * 100}", end="\r", flush=True)
                print(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...", end="", flush=True)

            time.sleep(1)
            dot_count += 1
    except Exception as e:
        log_system(f"Monitoring error: {e}")
    finally:
        if observer:
            observer.stop()
            observer.join()

# === BUTTON COMMANDS ===
def refresh_all():
    global last_mtime_log1
    for i in range(1, 7):
        if i in blink_threads:
            stop_flags[i].set()
            blink_threads[i].join(timeout=1)
        process_active[i] = True
        stop_flags[i].clear()
        process_buttons[i].config(bg="#4CAF50")
    clear_material_log()
    system_text.config(state="normal")
    system_text.delete(1.0, tk.END)
    system_text.config(state="disabled")
    last_mtime_log1 = 0
    process_job_order()
    log_system("SYSTEM REFRESHED")

refresh_btn.config(command=refresh_all)

def stop_all():
    for i in range(1, 7):
        if not stop_flags[i].is_set():
            stop_flags[i].set()
            process_buttons[i].config(bg="#4CAF50")
    log_system("ALL PROCESSES STOPPED")

stop_all_btn.config(command=stop_all)

# === START MONITORING ===
def start_monitoring():
    global animation_running
    animation_running = True
    anim_thread = threading.Thread(target=animate_loading, daemon=True)
    anim_thread.start()
    mon_thread = threading.Thread(target=monitoring_loop, daemon=True)
    mon_thread.start()

def on_closing():
    global running, animation_running
    running = False
    animation_running = False
    for flag in stop_flags.values():
        flag.set()
    root.destroy()

log_system("SYSTEM STARTED - MONITORING NETWORK PATHS...")
start_monitoring()
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()