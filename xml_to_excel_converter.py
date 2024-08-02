import os
import sys
import threading
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def xml_to_excel(xml_file, excel_file, progress):
    try:
        # Parse the XML file
        tree = ET.parse(xml_file)
        xml_root = tree.getroot()

        # Create a list to store the data
        data = []

        # Iterate through each <transaction> element
        for transaction in xml_root.findall('transaction'):
            # Create a dictionary for each transaction
            transaction_data = {}
            for elem in transaction:
                transaction_data[elem.tag] = elem.text
            data.append(transaction_data)

        # Create a DataFrame from the list of dictionaries
        df = pd.DataFrame(data)

        # Write the DataFrame to an Excel file
        df.to_excel(excel_file, index=False)
        return True
    except Exception as e:
        # Show error and stop further processing
        messagebox.showerror("Error", f"Failed to convert {xml_file} to Excel. Error: {e}")
        return False

def process_files():
    disable_buttons()  # Disable all buttons before processing
    files = file_listbox.get(0, tk.END)
    if files:
        total_files = len(files)
        progress_bar['maximum'] = total_files
        progress_bar['value'] = 0
        processed_files = 0

        while files:
            file = files[0]
            excel_file = file.replace('.xml', '.xlsx')
            if xml_to_excel(file, excel_file, progress_bar):
                file_listbox.delete(0)  # Remove the first file from the Listbox
                files = file_listbox.get(0, tk.END)  # Update the list of files
                processed_files += 1
                progress_bar['value'] = processed_files
                root.update_idletasks()
            else:
                break  # Stop processing if an error occurs
        
        if processed_files == total_files:
            # This ensures the progress bar fills up completely only if no errors occurred
            messagebox.showinfo("Success", "Files have been successfully converted to Excel.")
        elif processed_files > 0:
            # If some files were processed, but an error occurred
            messagebox.showinfo("Partial Success", f"{processed_files} files have been successfully converted to Excel.")
        
        # Reset the progress bar after processing
        progress_bar['value'] = 0
        enable_buttons()  # Re-enable buttons after processing
    # Update file count after processing
    update_file_count()

def export_to_excel():
    # Run file processing in a separate thread to keep the GUI responsive
    threading.Thread(target=process_files, daemon=True).start()

def update_file_count():
    """Update the file count label based on the number of files in the listbox."""
    file_count = file_listbox.size()
    file_count_label.config(text=f"Imported XML Files: {file_count}")

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("XML files", "*.xml")])
    if files:
        file_listbox.delete(0, tk.END)  # Clear any existing files in the listbox
        for file in files:
            file_listbox.insert(tk.END, file)
        export_button.config(state=tk.NORMAL)
        update_file_count()  # Update the file count

def clear_files():
    file_listbox.delete(0, tk.END)
    export_button.config(state=tk.DISABLED)  # Disable export button as the list is empty
    update_file_count()  # Update the file count

def disable_buttons():
    """Disable all buttons."""
    select_button.config(state=tk.DISABLED)
    clear_button.config(state=tk.DISABLED)
    export_button.config(state=tk.DISABLED)

def enable_buttons():
    """Enable all buttons."""
    select_button.config(state=tk.NORMAL)
    clear_button.config(state=tk.NORMAL)
    export_button.config(state=tk.NORMAL)

# Create the main application window
root = tk.Tk()
root.title("XML to Excel Converter")

# Load and resize the icon image
icon_image_path = resource_path('icon.png')
icon_image = Image.open(icon_image_path)
icon_image = icon_image.resize((100, 100), Image.Resampling.LANCZOS)
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(False, icon_photo)

# Set the icon for the taskbar (requires .ico format)
icon_ico_path = resource_path('icon.ico')
root.iconbitmap(icon_ico_path)

# Create a stylish frame
frame = ttk.Frame(root, padding="20")
frame.pack(expand=True, fill=tk.BOTH)

# Display the icon at the top center
icon_label = ttk.Label(frame, image=icon_photo)
icon_label.pack(pady=10)

# Title label
title_label = ttk.Label(frame, text="XML to Excel Converter", font=("Arial", 16, "bold"))
title_label.pack(pady=10)

# Description label
description_label = ttk.Label(frame, text="Select XML files to convert them to Excel format.", font=("Arial", 12))
description_label.pack(pady=10)

# Create a select button with padding and styling
select_button = ttk.Button(frame, text="Select XML Files", command=select_files)
select_button.pack(pady=10, ipadx=10, ipady=5)

# Create a clear button with padding and styling
clear_button = ttk.Button(frame, text="Clear List", command=clear_files)
clear_button.pack(pady=10, ipadx=10, ipady=5)

# Create a frame for the Listbox and Scrollbar
listbox_frame = ttk.Frame(frame)
listbox_frame.pack(pady=10, fill=tk.BOTH, expand=True)

# Listbox to display selected files
file_listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=10)
file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Create a vertical scrollbar and link it to the Listbox
scrollbar = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=file_listbox.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Configure the Listbox to use the scrollbar
file_listbox.config(yscrollcommand=scrollbar.set)

# Create an export button with padding and styling, initially disabled
export_button = ttk.Button(frame, text="Export to Excel", command=export_to_excel, state=tk.DISABLED)
export_button.pack(pady=10, ipadx=10, ipady=5)

# Create a progress bar
progress_bar = ttk.Progressbar(frame, orient='horizontal', mode='determinate', length=300)
progress_bar.pack(pady=10)

# Create a label to display the count of imported files
file_count_label = ttk.Label(frame, text="Imported XML Files: 0", font=("Arial", 12))
file_count_label.pack(pady=10)

# Create a footer label for credentials and copyright
footer_label = ttk.Label(frame, text="Version 1.0 | Lasindu Semasingha | Associated Motor Finance Co PLC | © 2024", font=("Arial", 10, "italic"))
footer_label.pack(side=tk.BOTTOM, pady=10)

# Start the Tkinter event loop
root.mainloop()
