import tkinter as tk
from tkinter import filedialog, messagebox
import os
import zipfile
import xml.etree.ElementTree as ET

def convert_format(input_str):
    parts = input_str.split('$')
    hashValue = parts[4]
    saltValue = parts[6]
    spinCount = parts[8]
    output_str = f"$office$2016$0${spinCount}${saltValue}${hashValue}"
    return output_str

def extract_sheet_protection_data_for_all_sheets(excel_file_path, output_directory):
    excel_base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    output_directory = os.path.join(output_directory, excel_base_name + "_SheetProtectionData")

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    try:
        with zipfile.ZipFile(excel_file_path, 'r') as z:
            sheet_files = [f for f in z.namelist() if f.startswith('xl/worksheets/')]

            for sheet_file in sheet_files:
                with z.open(sheet_file) as f:
                    xml_content = f.read()

                    tree = ET.fromstring(xml_content)
                    namespace = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    sheet_protection = tree.find('.//main:sheetProtection', namespace)

                    output_file_name = os.path.basename(sheet_file).split('.')[0] + '_sheetProtectionData.txt'
                    output_file_path = os.path.join(output_directory, output_file_name)

                    if sheet_protection is not None:
                        # Format attribute-value pairs and prepend a '$'
                        protection_data_str = "$" + "$".join([f'{attr}${value}' for attr, value in sheet_protection.attrib.items()])
                        converted_data_str = convert_format(protection_data_str)
                    else:
                        converted_data_str = "No sheetProtection data found."

                    with open(output_file_path, 'w', encoding='utf-8') as out_file:
                        out_file.write(converted_data_str)
        messagebox.showinfo("Success", "Processing completed successfully.")
    except zipfile.BadZipFile:
        messagebox.showerror("Error", "The file is not a valid zip file or is corrupted.")
    except FileNotFoundError:
        messagebox.showerror("Error", "The file does not exist at the specified path.")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        output_directory = filedialog.askdirectory(mustexist=True, title="Select Output Directory")
        if output_directory:
            extract_sheet_protection_data_for_all_sheets(file_path, output_directory)

app = tk.Tk()
app.title("Excel Sheet Protection Data Extractor")

select_file_button = tk.Button(app, text="Select Excel File", command=select_excel_file)
select_file_button.pack(pady=20)

app.mainloop()
