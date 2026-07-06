import os
import sys
import tempfile
import re
import warnings

# Ignore default openpyxl warnings for a silent execution
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    import pandas as pd
    import matplotlib.pyplot as plt
    from openpyxl.drawing.image import Image
except ImportError as e:
    print(f"\n❌ Error: A required library is missing: {e}")
    print("Please install them by running: pip install pandas openpyxl matplotlib pillow")
    input("Press Enter to exit...")
    sys.exit()

def browse_file():
    """Open a dialog to select an Excel file and read its columns"""
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, filepath)
        try:
            # Quick read of column names to update the combobox
            df = pd.read_excel(filepath, nrows=0)
            combo_columns['values'] = list(df.columns)
            if list(df.columns):
                combo_columns.current(0)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read the file: {e}")

def process_data():
    """Main function for sorting, validation, and dashboard creation"""
    input_file = entry_file_path.get()
    category_column = combo_columns.get()
    
    if not input_file or not category_column:
        messagebox.showwarning("Warning", "Please select a file and a categorization column.")
        return

    try:
        # 1. Read Data
        df = pd.read_excel(input_file)
        original_count = len(df)
        
        # 2. Data Validation
        is_error = df.isnull().any(axis=1) 
        
        df['Audit Notes'] = '✅ Valid'
        df.loc[is_error, 'Audit Notes'] = '⚠️ Missing or Empty Field'
        
        valid_data = df[~is_error].copy()
        error_data = df[is_error].copy()
        
        cleaned_count = len(valid_data)
        errors_count = len(error_data)

        if cleaned_count == 0:
            messagebox.showerror("Error", "All data contains errors or is empty. Cannot complete sorting.")
            return

        # 3. Dashboard Preparation
        summary_data = []
        grouped_data = valid_data.groupby(category_column, dropna=False)
        
        output_file = os.path.splitext(input_file)[0] + "_Analyzed.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            seen_sheets = {"Dashboard_Summary": True, "⚠️ Needs Review": True}
            
            # Process each valid category
            for category_name, category_df in grouped_data:
                count = len(category_df)
                total_salary = category_df['Salary'].sum() if 'Salary' in category_df.columns else 0
                
                summary_data.append({
                    'Category / Department': category_name,
                    'Number of Records': count,
                    'Total Salary': total_salary
                })
                
                # Clean invalid characters for Excel sheet names
                safe_name = re.sub(r'[\\*?:/\[\]]', '_', str(category_name)).strip()
                if not safe_name:
                    safe_name = "Category"
                
                base_sheet_name = safe_name[:31]
                sheet_name = base_sheet_name
                
                counter = 1
                while sheet_name.lower() in [s.lower() for s in seen_sheets.keys()]:
                    suffix = f"_{counter}"
                    sheet_name = base_sheet_name[:31 - len(suffix)] + suffix
                    counter += 1
                    
                seen_sheets[sheet_name] = True
                category_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Create Dashboard (Summary Sheet)
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="Dashboard_Summary", index=False)
            
            # Save records needing manual review in a separate sheet
            if errors_count > 0:
                error_data.to_excel(writer, sheet_name="⚠️ Needs Review", index=False)
            
            # 4. Data Visualization (Pie Chart)
            plt.figure(figsize=(8, 6))
            plt.pie(summary_df['Number of Records'], labels=summary_df['Category / Department'].astype(str), autopct='%1.1f%%', startangle=140)
            plt.title('Distribution of Valid Records by Category')
            
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                chart_path = tmpfile.name
                
            plt.savefig(chart_path)
            plt.close()
            
            # Insert the chart into Excel
            workbook = writer.book
            summary_sheet = workbook["Dashboard_Summary"]
            img = Image(chart_path)
            summary_sheet.add_image(img, 'E2') 
            
        # Clean up temporary chart image
        if os.path.exists(chart_path):
            try:
                os.remove(chart_path)
            except Exception:
                pass

        messagebox.showinfo("Success", 
                            f"Operation completed successfully!\n\n"
                            f"Total Records: {original_count}\n"
                            f"Sorted Valid Records: {cleaned_count}\n"
                            f"Records Needing Review: {errors_count}\n\n"
                            f"File saved as:\n{os.path.basename(output_file)}")
                            
    except Exception as e:
        messagebox.showerror("Unexpected Error", str(e))

# ==========================================
# GUI Setup
# ==========================================
entry_file_path = None
combo_columns = None

def main():
    global entry_file_path, combo_columns
    root = tk.Tk()
    root.title("Smart Data Analyzer & Sorter")
    root.geometry("500x250")
    try:
        root.eval('tk::PlaceWindow . center')
    except Exception:
        pass

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill=tk.BOTH)

    # File Selection
    file_frame = tk.Frame(frame)
    file_frame.pack(fill=tk.X, pady=5)
    tk.Label(file_frame, text="Select Excel File:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
    entry_file_path = tk.Entry(file_frame, width=35)
    entry_file_path.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    btn_browse = tk.Button(file_frame, text="Browse...", command=browse_file, width=10)
    btn_browse.pack(side=tk.LEFT)

    # Column Selection
    col_frame = tk.Frame(frame)
    col_frame.pack(fill=tk.X, pady=(15, 5))
    tk.Label(col_frame, text="Select Categorization Column:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
    combo_columns = ttk.Combobox(col_frame, state="readonly")
    combo_columns.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

    # Process Button
    btn_process = tk.Button(frame, text="Start Sorting & Analysis", bg="#2196F3", fg="white", font=("Arial", 11, "bold"), command=process_data)
    btn_process.pack(pady=30, fill=tk.X)

    root.mainloop()

if __name__ == "__main__":
    main()