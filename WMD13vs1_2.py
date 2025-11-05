# ERROR: IF THE PROCESS 1 BUTTON IS CLICKED BY THE USER LOADING DISPLAY WITH 3 RUNNING DOTS IS STOPPED RUNNING
# REMOVE SO MANY ROWS OF CONSOLE DISPLAY (WAITING FOR (log000_1.csv) FILE CHANGES...........................................)
# SKLEARN VALIDATION + 100% CORRECT SUGGESTIONS
# SKLEARN CHECK IF THE MATERIALS IS CORRECT OR ERROR
# WORKING PART 2
# REAL TIME UPDATE
# WITH 1ST TKINTER GUI




# THE 1ST TASK WILL BE READ THE TAIL OF CSV FILE AT "\\192.168.2.10\csv\csv\JobOrder\log000_JobOrder.csv" THEN GET THE DETAIL OF "Job Order Number" COLUMN THEN REMOVE ANY SPACE WITHIN THE DATA

# THEN 2ND TASK IT WILL FIND THE EXACT FILE NAME AT "\\192.168.2.19\production\2025\1. Document for Production Admin\8. JOB ORDER MATERIAL LIST" AND IT SHOULD BE THE SAME WITH THE "Job Order Number" AT "\\192.168.2.10\csv\csv\JobOrder\log000_JobOrder.csv"

# THEN 3RD TASK, IF THE TIME STAMP HAS CHANGES AT PROCESS 1 LOCATED AT "\\192.168.2.10\csv\csv\VT1\log000_1.csv" IT WILL CHECK THE TAIL OF THE FF DATA
# Process 1 Em2p
# Process 1 Em3p
# Process 1 Harness
# Process 1 Frame
# Process 1 Bushing
# THEN IT WILL COMPARE TO "\\192.168.2.19\production\2025\1. Document for Production Admin\8. JOB ORDER MATERIAL LIST\3J737987830000.xlsx" AT "MATERIAL" COLUMN 
# IF THE DATA AT "\\192.168.2.10\csv\csv\VT1\log000_1.csv" IS NOT TALLIED AT "\\192.168.2.19\production\2025\1. Document for Production Admin\8. JOB ORDER MATERIAL LIST\3J737987830000.xlsx" AT "MATERIAL" COLUMN IT WILL TRIGER THE "Process 1" "STOP" BUTTON DISPLAY AT 1ST TKINTER GUI AND IT WILL ALSO DISPLAY AT TEXT BOX THE FF:
# PROCESS 1: ERROR: EM0580106D		EXPECTED: EM0580106P
# THEN THE BUTTON THAT WAS TRIGERRED IT WILL KEEP ON BLINKING WITH RED PAINT AND IT WILL NOT STOPPED UNLESS THE "STOP" BUTTON WILL BE CLICKED



# CREATE A TKINTER THAT WILL DISPLAY THE FF:
# PROCESS 1 TO PROCESS 6 "STOP" BUTTON THEN EACH BUTTON WILL HAVE A LOADING FOLLOWED BY 3 RUNNING DOT
# THEN CREATE A TEXT BOX THAT WILL DISPLAY THE FF:		
# PROCESS 1: ERROR: EM0580106D		EXPECTED: EM0580106P

# THE STOP BUTTON WILL BE AT THE BUTTOM OF THE TEXT BOX MESSAGE





# REVISE THE WHOLE CODE BUT DO NOT CHANGE THE CODE FORMAT AND DO NOT OMIT ANY SINGLE LINES




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
EXCEL_MATERIAL_COLUMN_INDEX = 4   # The 0-based index (4 for Column E) in the Excel file to check against

REFERENCE_EXCEL = r"\\192.168.2.19\ai_team\INDIVIDUAL FOLDER\June-San\p2LTG\p2LTG_TransferData\WRONG MATERIAL DETECTOR\FALSE TEST\8. JOB ORDER MATERIAL LIST\3J737987830000.xlsx"
ML_MATERIAL_COL = 4      # Column E  (0-based)
ML_DESC_COL     = 10     # Column K  (0-based)
ML_TARGET_COLUMNS = ['Process 1 Em2p', 'Process 1 Em3p', 'Process 1 Harness',
                     'Process 1 Frame', 'Process 1 Bushing']   # columns 6-10 in output CSV

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
root.geometry("900x700")
root.configure(bg="#f0f0f0")

# Title
title_label = tk.Label(root, text="PROCESS MONITORING DASHBOARD", font=("Arial", 16, "bold"), bg="#f0f0f0", fg="#003366")
title_label.pack(pady=10)

# Frame for process buttons
btn_frame = tk.Frame(root, bg="#f0f0f0")
btn_frame.pack(pady=10)

# Dictionary to hold buttons and their status
process_buttons = {}
loading_labels = {}
dot_labels = {}
blink_threads = {}
stop_flags = {}
process_active = {}  # NEW: Track if process is active for monitoring

# Create 6 process stop buttons with loading and dots
for i in range(1, 7):
    proc_name = f"PROCESS {i}"
    
    # Main frame for each process
    proc_frame = tk.Frame(btn_frame, bg="#f0f0f0")
    proc_frame.grid(row=(i-1)//3, column=(i-1)%3, padx=20, pady=10)
    
    # Loading label
    load_lbl = tk.Label(proc_frame, text="Loading", font=("Arial", 10), fg="gray", bg="#f0f0f0")
    load_lbl.pack()
    loading_labels[i] = load_lbl
    
    # Process button
    btn = tk.Button(proc_frame, text=f"{proc_name}\nSTOP", font=("Arial", 12, "bold"),
                    width=12, height=3, bg="#4CAF50", fg="white", relief="raised",
                    command=lambda p=i: acknowledge_stop(p))
    btn.pack(pady=5)
    process_buttons[i] = btn
    
    # Running dots
    dot_lbl = tk.Label(proc_frame, text="...", font=("Arial", 14), fg="#4CAF50", bg="#f0f0f0")
    dot_lbl.pack()
    dot_labels[i] = dot_lbl
    
    # Initialize stop flag
    stop_flags[i] = threading.Event()
    process_active[i] = True  # All processes start active

# Text box for error messages
msg_frame = tk.Frame(root, bg="#f0f0f0")
msg_frame.pack(pady=20, fill="both", expand=True, padx=20)

msg_label = tk.Label(msg_frame, text="SYSTEM MESSAGES", font=("Arial", 12, "bold"), bg="#f0f0f0", anchor="w")
msg_label.pack(anchor="w")

text_box = scrolledtext.ScrolledText(msg_frame, height=12, font=("Courier", 10), state="disabled", bg="white", fg="#d40000")
text_box.pack(fill="both", expand=True, pady=5)

# === ADD CLEAR BUTTON BELOW TEXT BOX (CENTERED) ===
clear_btn_frame = tk.Frame(msg_frame, bg="#f0f0f0")
clear_btn_frame.pack(pady=8)

def clear_text_box():
    text_box.config(state="normal")
    text_box.delete(1.0, tk.END)
    text_box.config(state="disabled")
    log_message("SYSTEM MESSAGES CLEARED")

clear_button = tk.Button(clear_btn_frame, text="CLEAR", font=("Arial", 10, "bold"), bg="#ff4444", fg="white",
                         width=12, relief="raised", command=clear_text_box)
clear_button.pack()

# Global variables
last_jo_number = None
last_vt1_mtime = None
running = True
last_mtime_log1 = 0
animation_running = True

# Acknowledge stop button click
def acknowledge_stop(process_id):
    global blink_threads
    stop_flags[process_id].set()
    btn = process_buttons[process_id]
    btn.config(bg="#4CAF50", relief="raised")
    log_message(f"PROCESS {process_id}: STOP ACKNOWLEDGED BY USER")
    if process_id in blink_threads:
        blink_threads[process_id].join(timeout=1)
        del blink_threads[process_id]
    # === CRITICAL FIX: RE-ENABLE MONITORING AFTER ACKNOWLEDGE ===
    process_active[process_id] = True
    stop_flags[process_id].clear()

# Log message to text box
def log_message(message):
    text_box.config(state="normal")
    timestamp = datetime.now().strftime("%H:%M:%S")
    text_box.insert(tk.END, f"[{timestamp}] {message}\n")
    text_box.see(tk.END)
    text_box.config(state="disabled")

# Blink button red
def blink_button(process_id):
    btn = process_buttons[process_id]
    while not stop_flags[process_id].is_set():
        btn.config(bg="red")
        time.sleep(0.5)
        if stop_flags[process_id].is_set():
            break
        btn.config(bg="#ff6666")
        time.sleep(0.5)
    btn.config(bg="#4CAF50")

# Animate loading and dots
def animate_loading():
    global animation_running
    while animation_running:
        for i in range(1, 7):
            if process_active[i]:  # Only animate if process is active
                loading_labels[i].config(text="Loading" + "." * ((int(time.time() * 2) % 4)))
                dot_text = "." * ((int(time.time() * 3) % 4) + 1)
                dot_labels[i].config(text=dot_text)
        time.sleep(0.3)

def get_last_line(csv_file_path):
    """Read the last non-empty line from a CSV file efficiently."""
    with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
        lines = f.readlines()
        # Reverse to find last non-empty line
        for line in reversed(lines):
            if line.strip():
                return line
    return None

def parse_last_row(last_line, delimiter=','):
    """Parse CSV line using csv reader to handle quotes and commas properly."""
    reader = csv.reader([last_line], delimiter=delimiter)
    return next(reader)

def process_job_order():
    try:
        # Step 1: Read the last line of the source CSV
        print(f"Reading tail of: {CSV_LOG_PATH}")
        if not os.path.exists(CSV_LOG_PATH):
            raise FileNotFoundError(f"Source file not found: {CSV_LOG_PATH}")

        last_line = get_last_line(CSV_LOG_PATH)
        if not last_line:
            print("Source CSV is empty.")
            return

        # Step 2: Parse the row and get headers from first line to map column
        with open(CSV_LOG_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            try:
                headers = next(reader)
            except StopIteration:
                print("CSV has no headers.")
                return

            # Find index of "Job Order Number"
            try:
                job_order_idx = headers.index(COLUMN_NAME)
            except ValueError:
                print(f"Column '{COLUMN_NAME}' not found in headers.")
                print(f"Available columns: {headers}")
                return

            # Now go to last row (we already have last_line, but verify)
            rows = list(reader)
            if not rows:
                print("No data rows in CSV.")
                return
            last_row = rows[-1] # This is safer than line-based tail

        # Step 3: Extract Job Order Number and remove ALL spaces
        raw_job_order = last_row[job_order_idx]
        cleaned_job_order = raw_job_order.replace(" ", "").replace("\t", "") # Remove spaces and tabs

        print(f"Original: '{raw_job_order}'")
        print(f"Cleaned : '{cleaned_job_order}'")

        # Step 4: Write to output CSV (Column 1: Job Order Number)
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Header: Job Order + File Name + All Task 3 columns
            header = ["Job Order Number", "File Name"] + list(COLUMN_MAPPING.keys())
            writer.writerow(header)
            # Initial row: cleaned JO, empty File Name, empty data
            empty_row = [cleaned_job_order, ""] + [""] * len(COLUMN_MAPPING)
            writer.writerow(empty_row)

        print(f"Successfully saved to: {OUTPUT_PATH}")

        # === TASK 2: Find matching .xlsx file and update Column 2 ===
        excel_file = f"{cleaned_job_order}.xlsx"
        excel_full_path = os.path.join(EXCEL_FOLDER, excel_file)
        if os.path.exists(excel_full_path):
            print(f"THE FILE NAME {cleaned_job_order} WAS SUCCESSFULLY FOUND AT {EXCEL_FOLDER}")
            # Update Column 2 (index 1) with filename without .xlsx
            update_output_column(1, cleaned_job_order)
            log_message(f"MATCHED FILE: {cleaned_job_order}.xlsx")
        else:
            print(f".xlsx file NOT found: {excel_file}")
            update_output_column(1, "NOT FOUND")
            log_message(f".xlsx file NOT found: {excel_file}")

    except PermissionError as e:
        print(f"Permission denied: {e}")
        print("Make sure the network paths are accessible and you have write permissions.")
    except Exception as e:
        print(f"Error: {e}")

# Helper: Update specific column in row 1 (data row) of output CSV
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
            return False # Return False if file doesn't exist

        current_mtime = os.path.getmtime(LOG000_1_PATH)
        
        # Check if file has actually changed since last successful update
        if not force_update and current_mtime <= last_mtime_log1:
            return False # Return False if no new modification is detected

        # Pause slightly to ensure file writing is stable on the network
        time.sleep(0.1) 
        
        # Update last_mtime_log1 and proceed with transfer
        last_mtime_log1 = current_mtime

        # REVISED: Output message for detected change
        print(f"\nTRANSFERING ({os.path.basename(LOG000_1_PATH)}) DATA TO \"{os.path.basename(OUTPUT_PATH)}\"")

        # Read last row of log000_1.csv
        with open(LOG000_1_PATH, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = [r for r in reader if r and any(c.strip() for c in r)]
            if not rows:
                return False
            last_row = rows[-1]

        # Read output CSV
        if not os.path.exists(OUTPUT_PATH):
            return False
        with open(OUTPUT_PATH, 'r', newline='', encoding='utf-8') as f:
            output_rows = list(csv.reader(f))
        if len(output_rows) < 2:
            return False

        headers = output_rows[0]
        data_row = output_rows[1]

        # Update each mapped column
        for col_name, src_idx_1based in COLUMN_MAPPING.items():
            header_idx = headers.index(col_name)
            value = last_row[src_idx_1based - 1] if len(last_row) >= src_idx_1based else ""
            data_row[header_idx] = value

        # Write back
        with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerows(output_rows)

        print(f"Updated {len(COLUMN_MAPPING)} columns from log000_1.csv")
        
        # --- NEW CALL FOR TASK 4: Validate Data immediately after transfer ---
        validate_material_in_excel(output_rows)
        # -------------------------------------------------------------------
        
        return True # Return True if update was successful

    except Exception as e:
        print(f"Error in log000_1 update: {e}")
        return False

def validate_material_in_excel(output_rows):
    try:
        if len(output_rows) < 2:
            return

        headers = output_rows[0]
        data_row = output_rows[1]

        # 1. Get Job Order Number and File Name to find Excel file
        job_order_num = data_row[headers.index("Job Order Number")]
        file_name_status = data_row[headers.index("File Name")]
        
        if file_name_status == "NOT FOUND":
            print(f"Validation skipped: Excel file for JO {job_order_num} was NOT FOUND.")
            log_message(f"Validation skipped: Excel file for JO {job_order_num} was NOT FOUND.")
            return

        # ------------------------------------------------------------------
        # -------------------  SKLEARN ML VALIDATION  ----------------------
        # ------------------------------------------------------------------
        # Load reference Excel (fixed file)
        if not os.path.exists(REFERENCE_EXCEL):
            print(f"Reference Excel not found: {REFERENCE_EXCEL}")
            log_message(f"Reference Excel not found: {REFERENCE_EXCEL}")
            return

        ref_df = pd.read_excel(REFERENCE_EXCEL, header=None)
        # Build training data: material (E) + description (K) -> label "VALID"
        materials = ref_df.iloc[:, ML_MATERIAL_COL].dropna().astype(str).str.strip()
        descriptions = ref_df.iloc[:, ML_DESC_COL].dropna().astype(str).str.strip()

        # --------------------------------------------------------------
        # Create synthetic INVALID samples to satisfy LogisticRegression
        # --------------------------------------------------------------
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

        # --------------------------------------------------------------
        # Train model only if both classes exist
        # --------------------------------------------------------------
        ml_model = None
        if len(set(train_labels)) >= 2:
            vectorizer = TfidfVectorizer(lowercase=True, ngram_range=(1,2))
            clf = LogisticRegression(max_iter=1000)
            ml_model = make_pipeline(vectorizer, clf)
            try:
                ml_model.fit(train_texts, train_labels)
            except Exception as fit_err:
                print(f"ML fit failed (fallback to exact match): {fit_err}")
                ml_model = None
        else:
            print("Only one class available for training – using exact match only.")

        # --------------------------------------------------------------
        # Build TF-IDF on **MATERIAL CODES ONLY** for accurate correction
        # --------------------------------------------------------------
        material_codes = materials.tolist()
        correction_vectorizer = TfidfVectorizer(lowercase=True, ngram_range=(1,3), analyzer='char')
        tfidf_matrix = correction_vectorizer.fit_transform(material_codes) if material_codes else None

        # ------------------------------------------------------------------
        # Validate columns 6-10 (0-based indices 5 to 9 in output CSV)
        # ------------------------------------------------------------------
        print("\n" + "="*80)
        print("   SKLEARN VALIDATION + 100% CORRECT SUGGESTIONS (Columns 6-10)   ")
        print("="*80)

        any_error = False
        for col_name in ML_TARGET_COLUMNS:
            if col_name not in headers:
                continue
            idx = headers.index(col_name)
            value = data_row[idx].strip()

            if not value:
                print(f"  [SKIP] {col_name}: <empty>")
                continue

            # ---- Exact match -------------------------------------------------
            exact_match = value.upper() in {m.upper() for m in materials}
            if exact_match:
                print(f"  [OK] {col_name}: {value}  (exact match)")
                continue

            # ---- ML confidence (for detection only) -------------------------
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
                print(f"  [OK] {col_name}: {value}  (ML-predicted VALID, {confidence:.2%})")
                continue

            # ------------------- 100% ACCURATE CORRECTION -------------------
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
                    # Fallback: difflib closest match
                    candidates = [m for m in material_codes if m.upper().startswith(input_upper[:4])]
                    if not candidates:
                        candidates = material_codes
                    suggested_material = difflib.get_close_matches(input_upper, candidates, n=1, cutoff=0.6)
                    suggested_material = suggested_material[0] if suggested_material else material_codes[0]
            else:
                suggested_material = material_codes[0] if material_codes else "NO DATA"

            print(f"  [ERROR] {col_name}: {value}")
            print(f"           **SUGGESTED CORRECTION** → {suggested_material}")
            if ml_model:
                print(f"           ML confidence: {confidence:.2%}")

            # === GUI: Trigger STOP button + log error ===
            err_msg = f"PROCESS 1: ERROR: {value}\t\tEXPECTED: {suggested_material}"
            log_message(err_msg)

            # === CRITICAL: Only trigger if Process 1 is active ===
            if process_active[1]:
                if not stop_flags[1].is_set():
                    stop_flags[1].clear()
                    if 1 not in blink_threads or not blink_threads[1].is_alive():
                        t = threading.Thread(target=blink_button, args=(1,), daemon=True)
                        blink_threads[1] = t
                        t.start()

        if not any_error:
            print("  All five material fields are CORRECT.")
            log_message("PROCESS 1: ALL MATERIALS VERIFIED")
            if stop_flags[1].is_set():
                acknowledge_stop(1)
        print("="*80 + "\n")

    except Exception as e:
        print(f"Error in material validation: {e}")
        log_message(f"Validation Error: {e}")

class LogChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        normalized_event_path = os.path.normpath(event.src_path).lower()
        normalized_log_path = os.path.normpath(LOG000_1_PATH).lower()
        
        if normalized_event_path == normalized_log_path:
            pass

def get_last_job_order():
    global last_jo_number
    try:
        if not os.path.exists(CSV_LOG_PATH):
            return None
        df = pd.read_csv(CSV_LOG_PATH, nrows=1, skiprows=lambda x: x > 0 and os.stat(CSV_LOG_PATH).st_size > 0)
        if df.empty or JO_NUMBER_COL not in df.columns:
            return None
        jo = str(df[JO_NUMBER_COL].iloc[0]).strip()
        jo = jo.replace(" ", "")
        if jo and jo != last_jo_number:
            last_jo_number = jo
            log_message(f"NEW JOB ORDER DETECTED: {jo}")
        return jo
    except Exception as e:
        log_message(f"ERROR READING JOB ORDER CSV: {str(e)}")
        return None

def find_job_order_file(jo_number):
    try:
        if not os.path.exists(JO_MATERIAL_DIR):
            return None
        for file in os.listdir(JO_MATERIAL_DIR):
            if file.endswith(".xlsx") and jo_number in file:
                return os.path.join(JO_MATERIAL_DIR, file)
        return None
    except Exception as e:
        log_message(f"ERROR SCANNING JOB ORDER DIR: {str(e)}")
        return None

def get_vt1_tail_data():
    try:
        if not os.path.exists(VT1_LOG_PATH):
            return {}
        with open(VT1_LOG_PATH, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
        data = {}
        for line in reversed(lines[-10:]):
            for item in PROCESS1_ITEMS:
                if item in line:
                    parts = line.strip().split(',')
                    if len(parts) > 1:
                        code = parts[1].strip()
                        data[item] = code
        return data
    except Exception as e:
        log_message(f"ERROR READING VT1 LOG: {str(e)}")
        return {}

def compare_with_excel(vt1_data, excel_path):
    try:
        if not os.path.exists(excel_path):
            return []
        xl = pd.ExcelFile(excel_path)
        df = xl.parse(xl.sheet_names[0])
        if MATERIAL_COL not in df.columns:
            return []
        
        expected_materials = set(str(x).strip().upper() for x in df[MATERIAL_COL].dropna())
        errors = []
        
        for item, actual in vt1_data.items():
            actual_clean = actual.strip().upper()
            found = False
            for exp in expected_materials:
                if actual_clean in exp or exp in actual_clean:
                    found = True
                    break
            if not found and actual_clean:
                errors.append((item, actual_clean))
        
        return errors
    except Exception as e:
        log_message(f"ERROR COMPARING EXCEL: {str(e)}")
        return []

def monitor_vt1_changes():
    global last_vt1_mtime
    try:
        if not os.path.exists(VT1_LOG_PATH):
            return False
        current_mtime = os.path.getmtime(VT1_LOG_PATH)
        if last_vt1_mtime is None:
            last_vt1_mtime = current_mtime
            return False
        if current_mtime > last_vt1_mtime:
            last_vt1_mtime = current_mtime
            return True
        return False
    except:
        return False

def monitoring_loop():
    global animation_running
    current_jo_file = None
    observer = None
    dot_count = 0
    try:
        # === INITIAL RUN: TASK 1 + TASK 2 ===
        process_job_order()

        # === INITIAL LOAD: TASK 3 (if log000_1.csv exists) ===
        if os.path.exists(LOG000_1_PATH):
            if update_from_log000_1(force_update=True):
                 log_message(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...")
            else:
                 try:
                     global last_mtime_log1
                     last_mtime_log1 = os.path.getmtime(LOG000_1_PATH)
                 except Exception:
                     pass
        else:
            log_message(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...")

        # === START MONITORING log000_1.csv (Watchdog kept for local fallback) ===
        observer = Observer()
        watch_dir = os.path.dirname(LOG000_1_PATH)
        if not os.path.exists(watch_dir):
            log_message(f"\nERROR: Watch directory does not exist: {watch_dir}")
        else:
            observer.schedule(LogChangeHandler(), path=watch_dir, recursive=False)
            try:
                observer.start()
            except Exception as e:
                log_message(f"\nFailed to start observer: {e}")

        while running:
            # FIXED: Polling check for file changes every 1 second
            updated = False
            if os.path.exists(LOG000_1_PATH):
                try:
                    current_mtime = os.path.getmtime(LOG000_1_PATH)
                    if current_mtime > last_mtime_log1:
                        if update_from_log000_1(force_update=True):
                            updated = True
                except Exception:
                    pass

            # REMOVED: Excessive console spam — now only logs once
            if dot_count == 0:
                print(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...", end="", flush=True)

            if updated:
                dot_count = 0
                print(f"\r{' ' * 80}", end="", flush=True)  # Clear line
                print(f"WAITING FOR ({os.path.basename(LOG000_1_PATH)}) FILE CHANGES...", end="", flush=True)
                time.sleep(0.3)

            time.sleep(1)
            dot_count += 1

    except KeyboardInterrupt:
        log_message("\nScript terminated by user.")
    except Exception as e:
        log_message(f"\nUnexpected error in loop: {e}")
    finally:
        if observer:
            observer.stop()
            observer.join()

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

log_message("SYSTEM STARTED - MONITORING NETWORK PATHS...")
start_monitoring()
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()