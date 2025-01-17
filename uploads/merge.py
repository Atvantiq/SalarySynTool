import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Notebook, Style
import pandas as pd
import os

# Global variables
attendance1_file = ""
salary1_file = ""
attendance2_file = ""
salary2_file = ""
salary_file = None
bank_file = None

def upload_attendance1():
    global attendance1_file
    attendance1_file = filedialog.askopenfilename(title="Select 1st Month Attendance File", filetypes=[("CSV files", "*.csv")])
    if attendance1_file:
        messagebox.showinfo("File Uploaded", "1st Month Attendance file uploaded successfully!")

def upload_salary1():
    global salary1_file
    salary1_file = filedialog.askopenfilename(title="Select 1st Month Salary File", filetypes=[("CSV files", "*.csv")])
    if salary1_file:
        messagebox.showinfo("File Uploaded", "1st Month Salary file uploaded successfully!")

def upload_attendance2():
    global attendance2_file
    attendance2_file = filedialog.askopenfilename(title="Select 2nd Month Attendance File", filetypes=[("CSV files", "*.csv")])
    if attendance2_file:
        messagebox.showinfo("File Uploaded", "2nd Month Attendance file uploaded successfully!")

def upload_salary2():
    global salary2_file
    salary2_file = filedialog.askopenfilename(title="Select 2nd Month Salary File", filetypes=[("CSV files", "*.csv")])
    if salary2_file:
        messagebox.showinfo("File Uploaded", "2nd Month Salary file uploaded successfully!")

def upload_salary_file():
    global salary_file
    salary_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if salary_file:
        messagebox.showinfo("File Uploaded", f"Salary file {os.path.basename(salary_file)} uploaded successfully!")

def upload_bank_file():
    global bank_file
    bank_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if bank_file:
        messagebox.showinfo("File Uploaded", f"Bank file {os.path.basename(bank_file)} uploaded successfully!")
        
def merge_and_process(attendance_file, salary_file):
    try:
        # Mapping months to their number of days
        month_days = {
            'Jan': 31, 'Feb': 28, 'Mar': 31, 'Apr': 30,
            'May': 31, 'Jun': 30, 'Jul': 31, 'Aug': 31,
            'Sep': 30, 'Oct': 31, 'Nov': 30, 'Dec': 31
        }

        # Read files
        attendance_df = pd.read_csv(attendance_file)
        salary_df = pd.read_csv(salary_file)
        
         # Merge attendance and salary data
        merged_df = pd.merge(attendance_df, salary_df, on='EMP ID', how='inner')
        
        # Extract month from the attendance data
        month_year = attendance_df.iloc[0][2]  # the column is named 'Month'
        month = month_year.split('-')[0]       # Extracting 'Jul' from 'Jul-24'
        days_in_month = month_days[month]

        # Dynamically define date columns
        date_columns_1_15 = merged_df.columns[3:18]  # For 1st to 15th
        if days_in_month == 30:
            date_columns_16_31 = merged_df.columns[18:33]  # For 16th to 30th
        elif days_in_month == 31:
            date_columns_16_31 = merged_df.columns[18:34]  # For 16th to 31st
        elif days_in_month == 28:
            date_columns_16_31 = merged_df.columns[18:31]  # For 16th to 28th
        else:
            raise Exception("Invalid number of days in the month.")

   

        # Calculate present days
        merged_df['Present Days 1-15'] = merged_df[date_columns_1_15].apply(lambda row: (row != 'A').sum(), axis=1)
        merged_df['Present Days 16-31'] = merged_df[date_columns_16_31].apply(lambda row: (row != 'A').sum(), axis=1)
        merged_df['Present Days'] = merged_df['Present Days 1-15'] + merged_df['Present Days 16-31']

        # Calculate salary components
        merged_df['Net Payable'] = pd.to_numeric(merged_df['Net Payable'], errors='coerce')
        merged_df['Day Salary'] = merged_df.apply(
            lambda row: row['Net Payable'] / row['Present Days'] if row['Present Days'] > 0 else 0, axis=1)
        merged_df['Total First Half Salary'] = merged_df['Day Salary'] * merged_df['Present Days 1-15']
        merged_df['Total Second Half Salary'] = merged_df['Day Salary'] * merged_df['Present Days 16-31']

        return merged_df
    except Exception as e:
        raise Exception(f"Error during processing: {str(e)}")

def process_and_save():
    try:
        if not (attendance1_file and salary1_file and attendance2_file and salary2_file):
            raise Exception("All four files must be uploaded before processing!")

        # Process both months' data
        processed1 = merge_and_process(attendance1_file, salary1_file)
        processed2 = merge_and_process(attendance2_file, salary2_file)

        # Merge the outputs
        final_df = pd.merge(processed1, processed2, on='EMP ID', how='outer', suffixes=('_Month1', '_Month2'))
        

        # Remove duplicate columns
        columns_to_drop = [col for col in final_df.columns if col.endswith('_Month2') and col[:-7] in final_df.columns]
        final_df.drop(columns=columns_to_drop, inplace=True)
        
        # combining Sum of Total Second Half Salary from current month and Total First Half Salary from next month
        final_df['Combined Salary'] = final_df['Total Second Half Salary_Month1'] + final_df['Total First Half Salary_Month2']


        # Save the final file
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if save_path:
            final_df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"Final merged file saved successfully at {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def download_icici_sheet():
    if not salary_file or not bank_file:
        messagebox.showerror("Error", "Please upload both salary and bank files.")
        return

    salary_df = pd.read_csv(salary_file)
    bank_df = pd.read_csv(bank_file)

    try:
        # Ensure merged data is correct
        merged_df = pd.merge(salary_df, bank_df, on="Emp ID", how="inner")

        # Process data to prepare `icici_data`
        icici_data = []
        for _, row in merged_df.iterrows():
            pymt_mode = "IFT" if row["IFSC Code"].startswith("ICIC") else "NEFT"
            icici_data.append([
                "",#PYMT_PROD_TYPE_CODE
                pymt_mode,  # PYMT_MODE
                "342105001430",  # DEBIT_ACC_NO (example value)
                row["Name"],  # BNF_NAME
                row["Bank Account No."],  # BENE_ACC_NO
                row["IFSC Code"],  # BENE_IFSC
                row["Combined Salary"],  # AMOUNT
                "Salary ",  # DEBIT_NARR (example value)
                "Salary",  # CREDIT_NARR (example value)
                "",  # MOBILE_NUM (example empty field)
                "",  # EMAIL_ID (example empty field)
                "Salary ",  # REMARK (example value)
                "09-10-2024",  # PYMT_DATE (example value)
                "",#remark
                "",
                "",
                "",
                "",
                ""
            ])

        # Debugging to ensure `icici_data` rows are correct
        # print("ICICI Data Rows:")
        # for row in icici_data:
        #     print(len(row), row)

        # Create DataFrame
        icici_df = pd.DataFrame(icici_data, columns=[
            "PYMT_PROD_TYPE_CODE",
            "PYMT_MODE", "DEBIT_ACC_NO", "BNF_NAME", "BENE_ACC_NO", 
            "BENE_IFSC", "AMOUNT", "DEBIT_NARR", "CREDIT_NARR", 
            "MOBILE_NUM", "EMAIL_ID", "REMARK", "PYMT_DATE","REF_NO","ADDL_INFO1","ADDL_INFO2","ADDL_INFO3","ADDL_INFO4","ADDL_INFO5"
        ])
        # print("ICICI DataFrame Created Successfully")
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if save_path:
            icici_df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"ICICI Sheet saved to {save_path}")
    

    except Exception as e:
        print(f"An error occurred: {e}")

def download_bob_sheet():
    if not salary_file or not bank_file:
        messagebox.showerror("Error", "Please upload both salary and bank files.")
        return

    salary_df = pd.read_csv(salary_file)
    bank_df = pd.read_csv(bank_file)

    merged_df = pd.merge(salary_df, bank_df, left_on="Emp ID", right_on="Emp ID", how="inner")

    bob_data = []
    for _, row in merged_df.iterrows():
        transaction_type = "IFT" if row["IFSC Code"].startswith("BARB") else "NEFT"
        bob_data.append([
            "",
            "08/01/2025",  # Value Date
            "NEFT",  # Message Type
            "27800500000353",  # Default Debit Account Number
            row["Name"],
            row["Combined Salary"],
            row["IFSC Code"],
            row["Bank Account No."],
            transaction_type,
            "",
            "",
            "",
            "",
            "",
            "salary",
            "salary"
        ])

    bob_df = pd.DataFrame(bob_data, columns=["CUSTOM_DETAILS1",
        "Value Date", "Message Type", "Debit Account No.", "Beneficiary Name", "Payment Amount",
        "Beneficiary Bank Swift Code / IFSC Code", "Beneficiary Account No.", "Transaction Type Code",
        "CUSTOM_DETAILS2","CUSTOM_DETAILS3","CUSTOM_DETAILS4","CUSTOM_DETAILS5","CUSTOM_DETAILS6","Remarks","Purpose Of Payment"
    ])

    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if save_path:
        bob_df.to_csv(save_path, index=False)
        messagebox.showinfo("Success", f"BOB Sheet saved to {save_path}")

def create_gui():
    root = tk.Tk()
    root.title("Employee Management Tool")
    root.geometry("700x600")
    root.configure(bg="#f0f4f7")

    style = Style()
    style.theme_use("clam")
    style.configure("TNotebook", tabposition='n', background="#dce6f1", borderwidth=0)
    style.configure("TNotebook.Tab", font=("Arial", 12), padding=[10, 5], background="#d0e1f9", foreground="#003366")
    style.map("TNotebook.Tab", background=[("selected", "#4CAF50")], foreground=[("selected", "white")])

    notebook = Notebook(root)

    # Tab 1: Attendance and Salary Processing
    tab1 = tk.Frame(notebook, bg="#eaf3fa")
    tk.Label(tab1, text="Attendance and Salary Processor", font=("Helvetica", 18, "bold"), pady=10, bg="#eaf3fa", fg="#003366").pack()

    tk.Button(tab1, text="Upload 1st Month Attendance", command=upload_attendance1, bg="#4CAF50", fg="white", font=("Helvetica", 12), padx=10, pady=5).pack(pady=10)
    tk.Button(tab1, text="Upload 1st Month Salary", command=upload_salary1, bg="#4CAF50", fg="white", font=("Helvetica", 12), padx=10, pady=5).pack(pady=10)
    tk.Button(tab1, text="Upload 2nd Month Attendance", command=upload_attendance2, bg="#4CAF50", fg="white", font=("Helvetica", 12), padx=10, pady=5).pack(pady=10)
    tk.Button(tab1, text="Upload 2nd Month Salary", command=upload_salary2, bg="#4CAF50", fg="white", font=("Helvetica", 12), padx=10, pady=5).pack(pady=10)
    tk.Button(tab1, text="Process and Save", command=process_and_save, bg="#003366", fg="white", font=("Helvetica", 14, "bold"), padx=15, pady=10).pack(pady=20)
    notebook.add(tab1, text="Attendance & Salary")

    # Tab 2: Bank Sheet Generation
    tab2 = tk.Frame(notebook, bg="#f4f9fc")
    tk.Label(tab2, text="Bank Sheet Generator", font=("Helvetica", 18, "bold"), pady=10, bg="#f4f9fc", fg="#003366").pack()

    tk.Button(tab2, text="Upload Salary File", command=upload_salary_file, bg="#4CAF50", fg="white", font=("Helvetica", 12), padx=10, pady=5).pack(pady=10)
    tk.Button(tab2, text="Upload Bank Details File", command=upload_bank_file, bg="#4CAF50", fg="white", font=("Helvetica", 12), padx=10, pady=5).pack(pady=10)
    tk.Button(tab2, text="Download ICICI Sheet", command=download_icici_sheet, bg="#28a745", fg="white", font=("Helvetica", 14, "bold"), padx=15, pady=10).pack(pady=15)
    tk.Button(tab2, text="Download BOB Sheet", command=download_bob_sheet, bg="#007bff", fg="white", font=("Helvetica", 14, "bold"), padx=15, pady=10).pack(pady=15)
    notebook.add(tab2, text="Bank Sheets")

    notebook.pack(expand=True, fill="both", padx=10, pady=10)
    root.mainloop()

# Run the GUI
create_gui()
