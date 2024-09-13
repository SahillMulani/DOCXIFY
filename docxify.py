import tkinter as tk
import customtkinter
import tkinterDnD
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os

customtkinter.set_ctk_parent_class(tkinterDnD.Tk)

customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

# Create the main window
root = customtkinter.CTk()
root.geometry("500x200")
root.title("Docxify (PDF to DOCx)")

# Global variable to store the PDF file path
pdf_file = None

# Function to select PDF file
def select_pdf_file():
    global pdf_file
    pdf_file = filedialog.askopenfilename(
        title="Select PDF file", 
        filetypes=[("PDF files", "*.pdf")]
    )
    if pdf_file:
        pdf_label.configure(text=f"Selected: {os.path.basename(pdf_file)}")

# Function to convert PDF to DOCX
def convert_pdf_to_docx():
    if not pdf_file:
        messagebox.showwarning("No File", "Please select a PDF file.")
        return
    
    # Get the folder and file name of the selected PDF
    pdf_folder = os.path.dirname(pdf_file)
    pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]
    
    # Set the DOCX file path (same name and folder as the PDF)
    docx_file = os.path.join(pdf_folder, pdf_name + ".docx")
    
    try:
        # Convert PDF to DOCX
        cv = Converter(pdf_file)
        cv.convert(docx_file, start=0, end=None)
        cv.close()

        messagebox.showinfo("Success", f"Conversion successful!\nDOCX saved at {docx_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert: {str(e)}")

# UI Elements
pdf_button = customtkinter.CTkButton(root, 
                                     text="Select PDF", 
                                     command=select_pdf_file, 
                                     text_color="white",  # Text color
                                     fg_color="#1E90FF",  # Button background color (DodgerBlue)
                                     hover_color="#00BFFF",  # Hover color (DeepSkyBlue)
                                     font=("Space Grotesk", 18, "bold"))  # Custom font style
pdf_button.pack(pady=10, padx=30)

pdf_label = customtkinter.CTkLabel(root, 
                                   text="No PDF selected", 
                                   text_color="#eeeedd",  # Lighter color for dark mode
                                   font=("Space Grotesk", 14, "bold"))  # Custom font style
pdf_label.pack(pady=10, padx=30)


convert_button = customtkinter.CTkButton(root, 
                                         text="Convert to DOCx", 
                                         command=convert_pdf_to_docx, 
                                         text_color="white",  # Text color
                                         fg_color="#1E90FF",  # LimeGreen background color
                                         hover_color="#98FB98",  # PaleGreen hover color
                                         font=("Space Grotesk", 18, "bold"))  # Font styling
convert_button.pack(pady=10, padx=30)


# Run the application
root.mainloop()