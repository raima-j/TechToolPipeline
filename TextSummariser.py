from tkinter import filedialog
from transformers import BartTokenizer, BartForConditionalGeneration
import os
import pytesseract
from natsort import natsorted
from pdf2image import convert_from_path
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
import csv
from nltk import sent_tokenize

root=tk.Tk()
root.title("AI Text Summariser")
root.configure(bg='maroon')
root.iconbitmap(r'L:\Technical Tools\logo (1).ico')

print("\n----> Please note")
print ("\nThis program uses a pre-trained BART AI model.\nPlease be aware that the summaries may be inaccurate.\nThis summariser should only be used as a way of understanding contents of files and not as an extractor.\n")

# Model loading
model_name = "facebook/bart-large-cnn"
tokenizer = BartTokenizer.from_pretrained(model_name)
model = BartForConditionalGeneration.from_pretrained(model_name)


def ocr():
    poppler_label.grid()
    poppler.grid()
    summarise_button=tk.Button(root,text='Summarise',command=lambda: begin_ocr(input_folder.get(), poppler.get()))
    summarise_button.grid(row=6,column=1)
    input_folder_label.config(text="Input folder for PDF files:")
    input_folder_button.config(command=lambda: input_folder.insert(tk.END, filedialog.askdirectory()))
    text_label=tk.Label(root,text="Click on Summarise",bg='maroon',fg='white',font=16)
    text_label.grid(row=9,column=0)


def no_ocr():
    poppler_label.grid_remove()
    poppler.grid_remove()
    summarise_button=tk.Button(root,text='Summarise',command=lambda: begin(input_folder.get()))
    summarise_button.grid(row=6,column=1)
    input_folder_label.config(text="Input folder for Text files:")
    input_folder_button.config(command=lambda: input_folder.insert(tk.END, filedialog.askdirectory()))
    text_label=tk.Label(root,text="Click on Summarise",bg='maroon',fg='white',font=16)
    text_label.grid(row=9,column=0)

def summarize_part(text_chunk):
    inputs = tokenizer(text_chunk, return_tensors='pt', max_length=1024, truncation=True)
    input_ids = inputs['input_ids']
    attention_mask = inputs['attention_mask']
    summary_ids = model.generate(input_ids, attention_mask=attention_mask, max_length=1024, num_beams=4, early_stopping=True)
    summary = tokenizer.decode(summary_ids[0], skip_special_tokens=True)
    return summary


#Processing PDF Files
def begin_ocr(input_folder,poppler):
    print ("\nPlease wait while PDF files get processed...\n")
    chunk_size = 1024
    context_chunk = 100

    pdf_files_dir = input_folder
    output_dir = output_folder.get()
    summary_files=[f for f in os.listdir(output_dir) if f.endswith('.txt')]
    poppler_path = poppler
    save_file=save_file_enter.get()
    
    if pdf_files_dir and output_dir and poppler_path and save_file:
        if pdf_files_dir and os.path.exists(poppler_path):
            pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
            
        pdf_files = [f for f in os.listdir(pdf_files_dir) if f.endswith('.pdf')]
        pdf_files = natsorted(pdf_files)
        
        status_label.config(text='Processing files...',bg='maroon',fg='white')
        

        def process_files(batch):
            for filename in batch:
                try:
                    
                    images = convert_from_path(os.path.join(pdf_files_dir, filename), poppler_path=poppler_path)
                    text = ''
                    for count, page in enumerate(images):
                        page_text = pytesseract.image_to_string(page).lower()
                        text += page_text
                    tokens = tokenizer(text, return_tensors='pt', truncation=True)['input_ids'][0]
                    chunks = []
                    for i in range(0, len(tokens), chunk_size - context_chunk):
                        chunks.append(tokens[i:i + chunk_size])

                    summaries = []
                    for chunk in chunks:
                        chunk_text = tokenizer.decode(chunk, skip_special_tokens=True)
                        summaries.append(summarize_part(chunk_text))
                    final_summary = ' '.join(summaries)
                    
                    print(f"\n________________________________________________________\nSummary for {filename}:\n")
                    print (text)
                    print(final_summary)
                    
                    # Save the summaries to a csv file
                    output_file_path = os.path.join(output_dir, f"{save_file}.csv")
                    with open(output_file_path, 'a', encoding='utf-8') as f:
                        writer=csv.writer(f)
                        hyperlink=f'=HYPERLINK("{filename}")'
                        writer.writerow([hyperlink,final_summary])
                        
                except Exception as e:
                    print (f"\nError processing {filename}: \n{e}")
                    
                
        
        # batch processing
        batch_size=20
        pdf_batches=[pdf_files[i:i+batch_size] for i in range(0,len(pdf_files),batch_size)]         
        with ThreadPoolExecutor() as executor:
            executor.map(process_files,pdf_batches)
        status_label.config(text='Processed summaries.',bg='maroon',fg='white') 
        
        
    else:
        status_label.config(text='Please type all details.',bg='maroon',fg='white')
        print ("\n___________________________\nPlease fill all details.\n___________________________\n")
        return



#Prcoessing Text Files
def begin(input_folder):
    print ("\nPlease wait while TEXT files get processed...\n")
    chunk_size = 1024
    context_chunk = 100

    text_files_dir = input_folder
    output_dir = output_folder.get()
    save_file=save_file_enter.get()
    
    if text_files_dir and output_dir and save_file:
        text_files = [f for f in os.listdir(text_files_dir) if f.endswith('.txt')]
        text_files = natsorted(text_files)
        
        
        status_label.config(text='Processing files...',bg='maroon',fg='white')
        print ("\nPlease wait while files get processed...\n")

        def process_files(batch):
            
            for filename in batch:
                try:
                    with open(os.path.join(text_files_dir, filename), 'r', encoding='utf-8') as file:
                        text = file.read()
                        
                    if text=='':
                        print ("The file is empty.\n")
                    
                    tokens = tokenizer(text, return_tensors='pt', truncation=False)['input_ids'][0]
                    chunks = []
                    
                    for i in range(0, len(tokens), chunk_size - context_chunk):
                        chunks.append(tokens[i:i + chunk_size])
                        
                    summaries = []
                    for chunk in chunks:
                        chunk_text = tokenizer.decode(chunk, skip_special_tokens=True)
                        summaries.append(summarize_part(chunk_text))
                    final_summary = ' '.join(summaries)
                    
                    # summary
                    print(f"\n________________________________________________________\nSummary for {filename}:\n")
                    print(final_summary)
                    
                    
                    # Save the summaries to a csv file
                    output_file_path = os.path.join(output_dir, f"{save_file}.csv")
                    with open(output_file_path, 'a', encoding='utf-8') as f:
                        writer=csv.writer(f)
                        hyperlink=f'=HYPERLINK("{filename}")'
                        writer.writerow([hyperlink,final_summary])

                except Exception as e:
                    print (f"\nError processing {filename}: \n{e}")

        batch_size=20
        text_batches=[text_files[i:i+batch_size] for i in range(0,len(text_files),batch_size)]         
        with ThreadPoolExecutor() as executor:
            executor.map(process_files,text_batches)
        status_label.config(text='Processed summaries.',bg='maroon',fg='white') 
        
    else:
        status_label.config(text='Please type all details.',bg='maroon',fg='white')
        print ("\nPlease fill all details.\n")
        return
    
#Interface
poppler_label=tk.Label(root,text="Path to Poppler's bin:",bg='maroon',fg='white')
poppler_label.grid(row=0,column=0)
poppler=tk.Entry(root, width=50)
poppler.grid(row=0,column=1)


input_folder_label = tk.Label(root, text="Input folder for files:", bg='maroon', fg='white')
input_folder_label.grid(row=2, column=0)
input_folder = tk.Entry(root, width=50)
input_folder.grid(row=2, column=1)
input_folder_button = tk.Button(root, text='Browse', command=lambda: input_folder.insert(tk.END, filedialog.askdirectory()))
input_folder_button.grid(row=2, column=2)

output_folder_label=tk.Label(root,text='Output folder directory:',bg='maroon',fg='white')
output_folder_label.grid(row=3,column=0)
output_folder=tk.Entry(root,width=50)
output_folder.grid(row=3,column=1)
output_folder_button=tk.Button(root,text='Browse',command=lambda: output_folder.insert(tk.END,filedialog.askdirectory()))
output_folder_button.grid(row=3,column=2)

save_file=tk.Label(root,text='Save File as:',bg='maroon',fg='white')
save_file.grid(row=4,column=0)
save_file_enter=tk.Entry(root,width=50)
save_file_enter.grid(row=4,column=1)


label=tk.Label(root,text=" ",bg='maroon',fg='white')
label.grid(row=5,column=1)

text_label=tk.Label(root,text="Select an option to summarise from.",bg='maroon',fg='white',font=16)
text_label.grid(row=8,column=0)

convert_button = tk.Button(root, text="Process PDFs (OCR)", command=ocr)
convert_button.grid(row=9, column=1)

convert_button2 = tk.Button(root, text="Process Texts", command=no_ocr)
convert_button2.grid(row=8, column=1)


status_label=tk.Label(root,text="",bg='maroon',fg='white')
status_label.grid(row=10,column=1)

root.mainloop()

