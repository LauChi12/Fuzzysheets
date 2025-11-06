# THIS IS THE BRANCH VERSION (CSV file on local machine) 

import pandas as pd
from rapidfuzz import fuzz
#import gspread#
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog #new for CSV version
import threading
import os
import sys


# --- 0. CONFIGURATION/AUTHENTICATION ---
# allows for google login instead of saved token in program file (saved for posterity)
#from oauth2client import file, client, tools
#import gspread
#import os

#def get_gspread_client():
    #try:
       # token_path = os.path.expanduser('~/.gspread_token.json')
        #flow = client.flow_from_clientsecrets('client_secrets.json',
                                              #scope=['https://www.googleapis.com/auth/spreadsheets',
                                                     #'https://www.googleapis.com/auth/drive'])
        #storage = file.Storage(token_path)
        #creds = storage.get()
        #if not creds or creds.invalid:
            #creds = tools.run_flow(flow, storage)
        #return gspread.authorize(creds)

    #except Exception as e:
        #import tkinter.messagebox as messagebox
        #messagebox.showerror("Authentication Error", f"Could not authenticate with Google Sheets.\n\n{type(e).__name__}: {e}")
        #return None
        
"""sets default values for UI input fields"""
OUTPUT_WORKSHEET_TITLE = 'Fuzzy_Match_Results_OUTPUT'
COLUMN_TO_FILTER_DEFAULT = 'Service Name'
TARGET_TERM_DEFAULT = 'Example Search Term'
THRESHOLD_DEFAULT = 90
SHEET_IDENTIFIER_DEFAULT = 'sample_data.csv'

#--- 1. UI FNs ---

#Updates status label and adds a colored log entry
def update_status(message, color='black'):
    if 'status_label' in globals() and status_label.winfo_exists():
        root.after(0, lambda: status_label.config(text=message, fg=color))
        
    if 'log_box' in globals() and log_box.winfo_exists():
        root.after(0, lambda: log_box.insert(tk.END, message + "\n", color))
        root.after(0, lambda: log_box.see(tk.END))
        
#Opens a dialogue to select/upload the CSV file OR Excel workbook
def browse_file():
    file_path = filedialog.askopenfilename(
        defaultextension= ".csv",
        filetypes=[("CSV files", ".csv"), ("Microsoft Excel Workbook", "*.xlsx")]
        )
    #checks if there was actually a file selected then clears the sample text, filling it with the selected file/path 
    if file_path:
        entry_sheet_id.delete(0, tk.END)
        entry_sheet_id.insert(0, file_path)

#Wrap the core fuzzy matching logic in a separate thread to keep the UI running during processing/Google API authentication"""
def run_filter_threaded():
    #disable button while processing
    run_button.config(state=tk.DISABLED, text="Processing...")
    update_status("Starting filtering. Please wait for output file to be written...", 'blue')

    #get the values from UI input fields
    try:
        sheet_id = entry_sheet_id.get().strip()
        worksheet_title = entry_worksheet_title.get().strip()
        column_filter = entry_column_filter.get().strip()
        target_term = entry_target_term.get().strip()
        threshold = int(entry_threshold.get())

        #input violation
        if not all([sheet_id, worksheet_title, column_filter, target_term]):
            raise ValueError("All main fields must be filled out.")
        if threshold < 0 or threshold > 100:
            raise ValueError("Threshold must be between 0 and 100.")

    except ValueError as e:
        update_status(f"Configuration Error: {e}", 'red')
        run_button.config(state=tk.NORMAL, text="Run Filter")
        return
    
    #processing in new thread to keep from blocking the execution of the process
    thread = threading.Thread(
        target=run_filter,
        args=(sheet_id, worksheet_title, column_filter, target_term, threshold)
    )
    thread.start()
        
def run_filter(sheet_id, worksheet_title, column_filter, target_term, threshold):
    """core logic connecting, processing, and writing data"""
    

    global OUTPUT_WORKSHEET_TITLE #using global name for output sheet
    #GOOGLE SHEETS VERSION ONLY
    #try:
        #---1. Authenticate (OAuth 2.0 User Credentials)---
        #users don't have to reauthorize google each time they login: the token is stored locally
        #gc = get_gspread_client()
        #---2. Read the data ---
        #update_status(f"Connecting to spreadsheet ID/URL: {sheet_id}", 'blue')
        #is input url or name??
        #if "docs.google.com/spreadsheets" in sheet_id:
            #sheet = gc.open_by_url(sheet_id)
        #else:
            #sheet = gc.open(sheet_id)

        #worksheet = sheet.worksheet(worksheet_title)

    try:
        file_path = sheet_id
        update_status(f"Reading file: {file_path}", 'blue')

        if not os.path.exists(file_path):
            update_status(f"Error: file not found in path: '{file_path}'.", 'red')
            return
        #now read the file, determine if supported file
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.lower().endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            update_status("That file is not supported by this program. Please select a .csv or .xlsx file.")
            return

        #GOOGLE SHEETS ONLY getting all records and returning data as list of dictionaries(records)
        #data = worksheet.get_all_records()
        #df = pd.DataFrame(data)

        #preflight validation
        if column_filter not in df.columns:
            update_status(f"Error: Column '{column_filter}' not found in sheet. Check header row.", 'red')
            return

        #---3. Filtering the data ---
        update_status(f"Processing {len(df)} rows and filtering by '{target_term}' at threshold {threshold}...", 'blue')

        #fuzzy match ratio
        df['Ratio'] = df[column_filter].apply(lambda x: fuzz.ratio(str(x).lower(), target_term.lower()))
        #filter based on threshold input
        filtered_df = df[df['Ratio'] >= threshold]
        output_data = filtered_df[[column_filter, 'Ratio']].reset_index(names=['Original Row Index'])

        #---4. Write data
        #determine dynamic size of output sheet(add room for header)
        num_rows_needed =  len(output_data) + 1
        num_cols_needed = len(output_data.columns)

        #GOOGLE SHEETS ONLY: delete old tab if one exists - I put a note in the readme to let users know about this
        #update_status(f"Preparing output worksheet: {OUTPUT_WORKSHEET_TITLE}...", 'blue')
        #try:
            #old_ws = sheet.worksheet(OUTPUT_WORKSHEET_TITLE)
            #sheet.del_worksheet(old_ws)
        #except gspread.WorksheetNotFound:
            #pass #this is what we want to happen
 
        #actually make the sheet
        #output_ws = sheet.add_worksheet(
            #title = OUTPUT_WORKSHEET_TITLE,
            #rows=str(num_rows_needed),
            #cols=str(num_cols_needed)
       # )

        #write the header first, then put in data
        #output_ws.update([output_data.columns.values.tolist()] + output_data.values.tolist())

        #notify user in the UI log 
        matches_found = len(output_data)

        if matches_found > 0:
            update_status(f"Success! Found {matches_found} matching records. Results saved to '{OUTPUT_WORKSHEET_TITLE}'.", 'green')
        else:
            update_status(f"Completed: No matching records found for '{target_term}' at threshold {threshold}.", 'orange')

    #weird errors not caused by data
    except Exception as e:
        update_status(f"ERROR: {type(e).__name__}: {e}. Check sheet names or sharing permissions.", 'red')
    finally:
        root.after(0, lambda: run_button.config(state=tk.NORMAL, text="Run Filter"))

            
#--- 2. TKINTER UI SETUP ---
#---2a. Create main window
root = tk.Tk()
root.title("Data Quality Utility (Fuzzy Match)")
root.geometry("650x700")

#---2b. Tag boxes with fun colors
def setup_log_tags():
    log_box.tag_config('red', foreground='red')
    log_box.tag_config('blue', foreground='blue')
    log_box.tag_config('green', foreground='green')
    log_box.tag_config('orange', foreground='orange')

#---2c. Input box frame
input_frame = tk.LabelFrame(root, text="Configuration & Target", padx=10, pady=10)
input_frame.pack(padx=10, pady=10, fill="x")

#---2d. Create UI Elements
row_index = 0

#where they put the input spreadsheet info
tk.Label(input_frame, text="1. Input CSV File Name").grid(row=row_index, column=0, sticky="w",pady=2)
entry_sheet_id = tk.Entry(input_frame, width=50)
entry_sheet_id.insert(0, SHEET_IDENTIFIER_DEFAULT)
entry_sheet_id.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
row_index += 1

#tk.Label(input_frame, text="2. Input Tab Name").grid(row=row_index, column=0, sticky="w",pady=2)
#entry_worksheet_title = tk.Entry(input_frame, width=50)
#entry_worksheet_title.insert(0, "Sheet1")
#entry_worksheet_title.grid(row=row_index, column=1, padx=5, pady=2, sticky="ew")
#row_index += 1

browse_button = tk.Button(input_frame, text="Browse files..", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=2, sticky="e")
row_index += 1

tk.Label(input_frame, text="3. Target Column Name").grid(row=row_index, column=0, sticky="w",pady=2)
entry_column_filter = tk.Entry(input_frame, width=50)
entry_column_filter.insert(0, COLUMN_TO_FILTER_DEFAULT)
entry_column_filter.grid(row=row_index, column=1, padx=5, pady=2, sticky="ew")
row_index += 1

tk.Label(input_frame, text="4. Search Term").grid(row=row_index, column=0, sticky="w",pady=2)
entry_target_term = tk.Entry(input_frame, width=50)
entry_target_term.insert(0, TARGET_TERM_DEFAULT)
entry_target_term.grid(row=row_index, column=1, padx=5, pady=2, sticky="ew")
row_index += 1

tk.Label(input_frame, text="5. Similarity Threshold (eg,: % Match)").grid(row=row_index, column=0, sticky="w",pady=2)
entry_threshold = tk.Entry(input_frame, width=50)
entry_threshold.insert(0,str(THRESHOLD_DEFAULT))
entry_threshold.grid(row=row_index, column=1, padx=5, pady=2, sticky="ew")


#grid expands to fit input field
input_frame.grid_columnconfigure(1, weight=1)

#status field and run button, also make the button text visible on all OS (hopefully)
run_button = tk.Button(
    root,
    text="Run Filter",
    command=run_filter_threaded, bg="#1E88E5", fg="white",font=('Arial', 12, 'bold'), relief=tk.FLAT, activebackground="#1565C0", activeforeground="white")

run_button.pack(pady=15, fill="x", padx=10)
status_label = tk.Label(root, text="Ready to Run. Please fill out fields and click Run.", bd=1, relief=tk.SUNKEN, anchor="w", font=('Arial', 10))
status_label.pack(side=tk.BOTTOM, fill=tk.X)

#log box for output info
tk.Label(root, text="Process Log:", anchor="w").pack(fill="x", padx=10, pady=(10,0))
log_box = scrolledtext.ScrolledText(root, height=12, state=tk.NORMAL, font=('Consolas', 9))
log_box.pack(fill="both", padx=10, pady=5, expand=True)

#set up log colors and start UI
setup_log_tags()


if __name__ == '__main__':
    root.mainloop()







                

        
