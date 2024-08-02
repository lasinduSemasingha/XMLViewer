import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk

def xml_to_excel(xml_file, excel_file):
    try:
        # Parse the XML file
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Create a list to store the data
        data = []

        # Iterate through each <transaction> element
        for transaction in root.findall('transaction'):
            # Create a dictionary for each transaction
            transaction_data = {}
            for elem in transaction:
                transaction_data[elem.tag] = elem.text
            data.append(transaction_data)

        # Create a DataFrame from the list of dictionaries
        df = pd.DataFrame(data)

        # Write the DataFrame to an Excel file
        df.to_excel(excel_file, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert {xml_file} to Excel. Error: {e}")

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("XML files", "*.xml")])
    if files:
        for file in files:
            excel_file = file.replace('.xml', '.xlsx')
            xml_to_excel(file, excel_file)
        messagebox.showinfo("Success", "Files have been successfully converted to Excel.")

# Create the main application window
root = tk.Tk()
root.title("XML to Excel Converter")

# Set window size
root.geometry("400x200")

# Load an icon image (optional)
# icon_image = Image.open('icon.png')
# icon_photo = ImageTk.PhotoImage(icon_image)
# root.iconphoto(False, icon_photo)

# Create a stylish frame
frame = ttk.Frame(root, padding="20")
frame.pack(expand=True, fill=tk.BOTH)

# Title label
title_label = ttk.Label(frame, text="XML to Excel Converter", font=("Arial", 16, "bold"))
title_label.pack(pady=10)

# Description label
description_label = ttk.Label(frame, text="Select XML files to convert them to Excel format.", font=("Arial", 12))
description_label.pack(pady=10)

# Create a select button with padding and styling
select_button = ttk.Button(frame, text="Select XML Files", command=select_files)
select_button.pack(pady=20, ipadx=10, ipady=5)

# Start the Tkinter event loop
root.mainloop()
