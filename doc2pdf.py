import os
import tkinter as tk
from tkinter import filedialog
import convertapi
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("API_KEY")

if API_KEY is None:
    raise Exception("API_KEY not found.")

convertapi.api_secret = API_KEY

def convert_docx_to_pdf(input_file, output_file):
    try:
        result = convertapi.convert('pdf', {'File': input_file})
        result.file.save(output_file)
        return True
    except Exception as e:
        print(f"Error converting DOCX to PDF: {e}")
        return False

def select_input_file():
    input_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, input_file)

def select_output_file():
    output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_file)

def convert_button_click():
    input_file = input_entry.get()
    output_file = output_entry.get()

    if not input_file or not output_file:
        result_label.config(text="Please select input and output files.")
        return

    if convert_docx_to_pdf(input_file, output_file):
        result_label.config(text=f"Conversion complete. PDF saved as {output_file}")
    else:
        result_label.config(text="Conversion failed. Please check the input file.")

root = tk.Tk()
root.title("DOC2PDF Converter")

input_label = tk.Label(root, text="Input DOCX File:")
input_label.pack()

input_entry = tk.Entry(root, width=40)
input_entry.pack()

input_button = tk.Button(root, text="Browse", command=select_input_file)
input_button.pack()

output_label = tk.Label(root, text="Output PDF File Location:")
output_label.pack()

output_entry = tk.Entry(root, width=40)
output_entry.pack()

output_button = tk.Button(root, text="Save As", command=select_output_file)
output_button.pack()

convert_button = tk.Button(root, text="Convert", command=convert_button_click)
convert_button.pack()

result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()