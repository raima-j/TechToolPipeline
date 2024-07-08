import os
import tkinter as tk
from tkinter import filedialog
from pdf2image import convert_from_path
import pytesseract
from concurrent.futures import ThreadPoolExecutor

num=0
import PIL.Image
PIL.Image.MAX_IMAGE_PIXELS = 933120000

# Function to convert PDFs to text files
def convert_to_text():
    folder_path = input_folder.get()
    save_path = output_folder.get()
    poppler_path=poppler.get()
    
    

    if folder_path and save_path:
        # Poppler's bin folder
       
        # Tesseract executable
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"

        # List of PDF files in the input folder
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        total=len(pdf_files)
        # List of existing text files in the output folder
        text_files = [f for f in os.listdir(save_path) if f.endswith('.txt')]
        error_occurred = False

        # Function to process each PDF file
        # Function to process each PDF file
        def process_pdf(pdf_file):
            global num
            
            name = os.path.splitext(pdf_file)[0]
            try:
                if f"{name}.txt" in text_files:
                    print(f"\n________________________________________________\nFile {name}.txt already extracted.")
                    num+=1
                    return
                images = convert_from_path(os.path.join(folder_path, pdf_file), poppler_path=poppler_path)
                text = ""
                for count, page in enumerate(images):
                    text += pytesseract.image_to_string(page)
                with open(os.path.join(save_path, f"{name}.txt"), 'w', encoding='utf-8') as txt_file:
                    txt_file.write(text)
                    
                num+=1
                print(f"\n____________________________________________________\nFile {name}.txt processed successfully.\nFiles remaining:{total-num}\n")
            except Exception as e:
                print(f"\n____________________________________________________\nError processing {pdf_file}: {e}")

                

        # Process PDFs using multithreading
        with ThreadPoolExecutor() as executor:
            executor.map(process_pdf, pdf_files)
            

        output_folder_label = tk.Label(root, text="Conversion Complete!",bg='maroon',fg='white')
        output_folder_label.grid(row=3, column=2)
    else:
        output_folder_label = tk.Label(root, text="Please provide input and output folder paths.",bg='maroon',fg='white')
        output_folder_label.grid(row=3, column=0)

# Creating the main window
root = tk.Tk()
root.title("PDF to Text Converter")
root.configure(bg='maroon')
root.iconbitmap(r'L:\Technical Tools\logo (1).ico')


#Username for device
poppler_label = tk.Label(root, text="Path to Poppler's bin:",bg='maroon',fg='white')
poppler_label.grid(row=0, column=0)
poppler = tk.Entry(root, width=50)
poppler.grid(row=0, column=1)


# Input folder
input_folder_label = tk.Label(root, text="Input Folder:",bg='maroon',fg='white')
input_folder_label.grid(row=1, column=0)
input_folder = tk.Entry(root, width=50)
input_folder.grid(row=1, column=1)
input_folder_button = tk.Button(root, text="Browse", command=lambda: input_folder.insert(tk.END, filedialog.askdirectory()))
input_folder_button.grid(row=1, column=2)

# Output folder
output_folder_label = tk.Label(root, text="Output Folder:",bg='maroon',fg='white')
output_folder_label.grid(row=2, column=0)
output_folder = tk.Entry(root, width=50)
output_folder.grid(row=2, column=1)
output_folder_button = tk.Button(root, text="Browse", command=lambda: output_folder.insert(tk.END, filedialog.askdirectory()))
output_folder_button.grid(row=2, column=2)

# Conversion button
convert_button = tk.Button(root, text="Convert to Text", command=convert_to_text)
convert_button.grid(row=3, column=1)

root.mainloop()
