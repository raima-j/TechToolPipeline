### Extract Tables from PDFs ###

import os
import tkinter as tk
from tkinter import filedialog
import camelot
import pandas as pd

def extract_tables():
    
    start_page = int(start_page_entry.get())
    end_page = int(end_page_entry.get())
    specify_pages = specify_page_entry.get()
    pdf_path = pdf_path_entry.get()
    file_name = file_name_entry.get()
    save_path = save_path_entry.get()

    try:
        engine_kwargs = {'options': {'strings_to_numbers': True}}  #Converting numeric strings to numbers
        
        # Range of pages
        if start_page != 0 and end_page != 0:
    
            tables = camelot.read_pdf(pdf_path, flavor='stream', pages=f'{start_page}-{end_page}', engine_kwargs=engine_kwargs)

            if tables:
                print ("\nExtracting tables...")
                
                if not os.path.exists(save_path):
                    os.makedirs(save_path)
                    
                with pd.ExcelWriter(os.path.join(save_path, f'{file_name}.xlsx'),engine_kwargs=engine_kwargs) as writer:
                    
                    for i, table in enumerate(tables):
                        table_df = table.df        
                        page_number=start_page+i
                        tab_name=f"Page_{page_number}"
                        
                        # Saving tables with tab name.
                        table_df.to_excel(writer, index=False,sheet_name=tab_name)

                status_label.config(text="Tables extracted successfully",bg='maroon',fg='white')
            else:
                print ("\nNo tables were found.")
                status_label.config(text="No tables found",bg='maroon',fg='white')
                
        # Specific pages
        elif specify_pages != 0:
            
            pages = specify_pages.split(',')
            print("\n____________________________________________________\nExtracting tables...")
            tables = camelot.read_pdf(pdf_path, flavor='stream', pages=specify_pages, engine_kwargs=engine_kwargs)

            if tables:
                print("\n_________________________________________________________\nTables extracted.")
                # Create directory if it doesn't exist
                if not os.path.exists(save_path):
                    os.makedirs(save_path)

                with pd.ExcelWriter(os.path.join(save_path, f'{file_name}.xlsx'),engine_kwargs=engine_kwargs) as writer:
                    
                    for i, table in enumerate(tables):
                        table_df = table.df
                        table_number=i+1
                        tab_name=f"Table_{table_number}"

                        table_df.to_excel(writer, index=False,sheet_name=tab_name)

                status_label.config(text="Tables extracted successfully",bg='maroon',fg='white')
            else:
                status_label.config(text="No tables found",bg='maroon',fg='white')
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}",bg='maroon',fg='white')

# GUI

root = tk.Tk()
root.title("Table Extractor")
root.configure(bg='maroon')
#root.iconbitmap(r'L:\Technical Tools\logo (1).ico')


tk.Label(root, text="Start Page:",bg='maroon',fg='white').grid(row=0, column=0)
start_page_entry = tk.Entry(root)
start_page_entry.grid(row=0, column=1)

tk.Label(root, text="End Page:",bg='maroon',fg='white').grid(row=1, column=0)
end_page_entry = tk.Entry(root)
end_page_entry.grid(row=1, column=1)

tk.Label(root, text="(or) Specify Pages (123,12,34...):",bg='maroon',fg='white').grid(row=2, column=0)
specify_page_entry = tk.Entry(root)
specify_page_entry.grid(row=2, column=1)

tk.Label(root, text="PDF Path:",bg='maroon',fg='white').grid(row=3, column=0)
pdf_path_entry = tk.Entry(root)
pdf_path_entry.grid(row=3, column=1)

def browse_pdf_path():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_path_entry.delete(0, tk.END)
    pdf_path_entry.insert(0, pdf_path)

browse_button = tk.Button(root, text="Browse", command=browse_pdf_path)
browse_button.grid(row=3, column=2)

tk.Label(root, text="File Name:",bg='maroon',fg='white').grid(row=4, column=0)
file_name_entry = tk.Entry(root)
file_name_entry.grid(row=4, column=1)

tk.Label(root, text="Save Path:",bg='maroon',fg='white').grid(row=5, column=0)
save_path_entry = tk.Entry(root)
save_path_entry.grid(row=5, column=1)

extract_button = tk.Button(root, text="Extract Tables", command=extract_tables)
extract_button.grid(row=6, column=1)

status_label = tk.Label(root, text="",bg='maroon',fg='white')
status_label.grid(row=7, column=0, columnspan=3)

root.mainloop()
