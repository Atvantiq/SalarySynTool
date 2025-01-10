import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Style, Progressbar


# Function to merge and process individual sheets
def merge_and_process(attendance_file, salary_file):
    try:
        # Merge attendance and salary data
        attendance_df = pd.read_csv(attendance_file)
        salary_df = pd.read_csv(salary_file)
        merged_df = pd.merge(attendance_df, salary_df, on='EMP ID', how='inner')
        

        # Process data 
        date_columns_1_15 = merged_df.columns[3:18]
        date_columns_16_31 = merged_df.columns[18:34]

        merged_df['Present Days 1-15'] = merged_df[date_columns_1_15].apply(lambda row: (row != 'A').sum(), axis=1)
        merged_df['Present Days 16-31'] = merged_df[date_columns_16_31].apply(lambda row: (row != 'A').sum(), axis=1)
        merged_df['Present Days'] = merged_df['Present Days 1-15'] + merged_df['Present Days 16-31']

        merged_df['Net Payable'] = pd.to_numeric(merged_df['Net Payable'], errors='coerce')
        merged_df['Day Salary'] = merged_df.apply(
            lambda row: row['Net Payable'] / row['Present Days'] if row['Present Days'] > 0 else 0, axis=1)
        merged_df['Total First Half Salary'] = merged_df['Day Salary'] * merged_df['Present Days 1-15']
        merged_df['Total Second Half Salary'] = merged_df['Day Salary'] * merged_df['Present Days 16-31']
        
        
        return merged_df
    except Exception as e:
        raise Exception(f"Error during processing: {str(e)}")


# Global variables to store file paths
attendance1_file = ""
salary1_file = ""
attendance2_file = ""
salary2_file = ""


# Functions to upload files
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


# Final processing function
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
        
        # Add new column: Sum of Total Second Half Salary from current month and Total First Half Salary from next month
        # merged_df['Total First Half Salary_Month2'] = merged_df['Total First Half Salary'].shift(-1)
        # merged_df['Total Second Half Salary_Month1'] = merged_df['Total Second Half Salary']
        final_df['Combined Salary'] = final_df['Total Second Half Salary_Month1'] + final_df['Total First Half Salary_Month2']


        # Save the final file
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if save_path:
            final_df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"Final merged file saved successfully at {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


#  GUI
def create_gui():
    root = tk.Tk()
    root.title("Employee Attendance and Salary Processor")
    root.geometry("500x400")

    # Style for progress bar
    style = Style()
    style.configure("TProgressbar", thickness=10)

    # Header
    header_label = tk.Label(root, text="Attendance and Salary Processor of atvantiq networks", font=("Helvetica", 16, "bold"))
    header_label.pack(pady=10)

    # Instructions
    instructions_label = tk.Label(
        root,
        text="Upload files in sequence:\n1. 1st Month Attendance\n2. 1st Month Salary\n3. 2nd Month Attendance\n4. 2nd Month Salary",
        font=("Helvetica", 10), wraplength=450, justify="center"
    )
    instructions_label.pack(pady=10)

    # Buttons for file uploads
    tk.Button(root, text="Upload 1st Month Attendance", command=upload_attendance1).pack(pady=5)
    tk.Button(root, text="Upload 1st Month Salary", command=upload_salary1).pack(pady=5)
    tk.Button(root, text="Upload 2nd Month Attendance", command=upload_attendance2).pack(pady=5)
    tk.Button(root, text="Upload 2nd Month Salary", command=upload_salary2).pack(pady=5)

    # Process and save button
    tk.Button(root, text="Process and Save", command=process_and_save, bg="#4CAF50", fg="white", font=("Helvetica", 12)).pack(pady=20)

    root.mainloop()


# Run the GUI
create_gui()
