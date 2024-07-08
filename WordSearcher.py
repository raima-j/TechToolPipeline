### Word Searcher ###


import os
import tkinter as tk
from tkinter import filedialog, ttk, Menu
from concurrent.futures import ThreadPoolExecutor
import pytesseract
from pdf2image import convert_from_path
import re
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
from natsort import natsorted
import csv
from collections import Counter
import matplotlib.pyplot as plt
import matplotlib
import seaborn as sns

import PIL.Image
PIL.Image.MAX_IMAGE_PIXELS = 933120000

# Initialize NLTK
nltk.download('punkt')
nltk.download('stopwords')
token_counter = Counter()

print ("\n_________________________________________________________\n")

matplotlib.use('TkAgg')
root = tk.Tk()
root.title("PDF Word Search")
root.configure(bg='maroon')
root.iconbitmap(r'logo.png')


words = []

def keywords():
    with open(r"...Technical Tools\keywords.txt", 'r') as f:
        keywords = f.read().splitlines()
    return keywords

num = 0  

def clear_words():
    for entry in words:
        entry.destroy()
    words.clear()

def get_words():
    clear_words()
    num = int(number.get())
    for i in range(num):
        word_label = tk.Label(root, text=f"Word {i+1}:",bg='maroon',fg='white')
        word_label.grid(row=i+6, column=0)
        word_entry = tk.Entry(root, width=50)
        word_entry.grid(row=i+6, column=1)
        words.append(word_entry)

def reset():
    global num  
    num = 0  
    clear_words()
    input_folder.delete(0, tk.END)
    number.delete(0, tk.END)
    tree.delete(*tree.get_children())

def search_files(start_file, end_file):
    print ("\nPlease wait, extracting...\n")
    global num  
    num = 0
    
    folder_path = input_folder.get()
    poppler_path = poppler.get()
    save_file = save_path_entry.get()
    
    with open(os.path.join(folder_path, f'{save_file}.csv'), 'a',encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["-------------------------------FILE PATH------------------------------", "---PAGE NUM---", "-----WORD-----", "-------------------------------SENTENCE------------------------------"])

    if folder_path and os.path.exists(poppler_path):
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"

        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        pdf_files = natsorted(pdf_files)
        
        indexed_files = {file: idx for idx, file in enumerate(pdf_files)}
        start_index = indexed_files.get(start_file)
        end_index = indexed_files.get(end_file) + 1 if end_file in indexed_files else len(pdf_files)
        pdf_files = pdf_files[start_index:end_index]
        
        total_count = len(pdf_files)
        
        word_patterns = '|'.join([re.escape(word.get()) for word in words])
        word_regex = re.compile(word_patterns, re.IGNORECASE)
        
        all_matches = []

        def process_pdf(pdf_file):
            global num  
            matches = []
            try:
                images = convert_from_path(os.path.join(folder_path, pdf_file), poppler_path=poppler_path)
                for count, page in enumerate(images):
                    page_text = pytesseract.image_to_string(page).lower()
                    
                    new_tokens = keywords()
                    tokens = word_tokenize(page_text)
                    
                    sentences = sent_tokenize(page_text)
                    
                    text = []
                    line = ''
                    token_counter.update(new_tokens)
                    
                    for sentence in sentences:
                        length = len(sentence)
                        for i in range(length):
                            if sentence[i].isalnum() or sentence[i] == ' ':
                                line += sentence[i]
                        text.append(line)
                        line = ''
                        
                        cleaned_text = [line.replace('\n', ' ') for line in text]

                    for sentence in cleaned_text:
                        current_matches = [(os.path.join(folder_path, pdf_file), count+1, match, sentence) for match in word_regex.findall(sentence) if match.lower() not in stopwords.words('english') and not match.isdigit()]
                        matches.extend(current_matches)
                        if current_matches:
                            print ("\nMatch found:\n",current_matches)
                num += 1
                print (f"\nFile {pdf_file} processed successfully.")
                print (f"\nFiles left: {total_count-num}\n___________________________________________________________________________________\n")
                with open(os.path.join(folder_path, f'{save_file}.csv'), 'a') as f:
                    writer = csv.writer(f)
                    for match in matches:
                        hyperlink = f'=HYPERLINK("{match[0]}", "{match[0]}")'
                        writer.writerow([hyperlink, match[1], match[2], match[3]])
            except Exception as e:
                print(f"Error processing '{pdf_file}': {str(e)}")
            return matches

        with ThreadPoolExecutor() as executor:
            for pdf_file in pdf_files:
                matches = process_pdf(pdf_file)
                all_matches.extend(matches)
        
        update_treeview(all_matches)

        output_folder_label = tk.Label(root, text="Search Complete!",bg='maroon',fg='white')
        output_folder_label.grid(row=3, column=2)
        
        top_10_words = token_counter.most_common(10)
    
        common_words, count_words = zip(*top_10_words)
        
        plt.figure(figsize=(12, 8))
        sns.barplot(x=list(common_words), y=list(count_words), palette='rocket')
        plt.title("10 Most Common Words")
        plt.xticks(rotation=45)
        plt.show()
        
    else:
        output_folder_label = tk.Label(root, text="Please provide valid input folder and Poppler path.",bg='maroon',fg='white')
        output_folder_label.grid(row=6, column=0)

def update_treeview(matches):
    tree.delete(*tree.get_children())
    for match in matches:
        tree.insert('', 'end', values=match)
        tree.insert('','end',None)

    tree.bind("<Double-1>", open_file)

def open_file(event):
    selected_item = tree.selection()
    if selected_item:
        filepath = tree.item(selected_item, "values")[0]
        os.startfile(filepath)

def copy_to_clipboard(event):
    selected_item = tree.selection()
    if selected_item:
        selected_text = tree.item(selected_item, "values")[3]
        root.clipboard_clear()
        root.clipboard_append(selected_text)

poppler_label = tk.Label(root, text="Path to Poppler's bin:",bg='maroon',fg='white')
poppler_label.grid(row=0, column=0)
poppler = tk.Entry(root, width=50)
poppler.grid(row=0, column=1)

input_folder_label = tk.Label(root, text="Input Folder:",bg='maroon',fg='white')
input_folder_label.grid(row=1, column=0)
input_folder = tk.Entry(root, width=50)
input_folder.grid(row=1, column=1)
input_folder_button = tk.Button(root, text="Browse", command=lambda: input_folder.insert(tk.END, filedialog.askdirectory()))
input_folder_button.grid(row=1, column=2)

start_file_label = tk.Label(root, text="Start File:",bg='maroon',fg='white')
start_file_label.grid(row=2, column=0)
start_file_entry = tk.Entry(root, width=50)
start_file_entry.grid(row=2, column=1)

end_file_label = tk.Label(root, text="End File:",bg='maroon',fg='white')
end_file_label.grid(row=3, column=0)
end_file_entry = tk.Entry(root, width=50)
end_file_entry.grid(row=3, column=1)

save_path_label = tk.Label(root, text="Save Name:",bg='maroon',fg='white')
save_path_label.grid(row=4, column=0)
save_path_entry = tk.Entry(root, width=50)
save_path_entry.grid(row=4, column=1)

num_words = tk.Label(root, text='Enter number of words to be searched:',bg='maroon',fg='white')
num_words.grid(row=5, column=0)
number = tk.Entry(root, width=50)
number.grid(row=5, column=1)

enter_button = tk.Button(root, text="Enter", command=get_words)
enter_button.grid(row=5, column=2)

search_button = tk.Button(root, text="Search", command=lambda: search_files(start_file_entry.get(), end_file_entry.get()))
search_button.grid(row=5, column=4)

reset_button = tk.Button(root, text="Reset", command=reset)
reset_button.grid(row=5, column=6)

tree = ttk.Treeview(root, columns=('PDF File', 'Page', 'Word', 'Sentence'), show='headings', height=20)
tree.heading('PDF File', text='PDF File')
tree.heading('Page', text='Page')
tree.heading('Word', text='Word')
tree.heading('Sentence', text='Sentence')

scrollbar = ttk.Scrollbar(root, orient='vertical', command=tree.yview)
scrollbar.grid(row=30, column=14, sticky='ns')
tree.configure(yscrollcommand=scrollbar.set)

tree.grid(row=30, column=0, columnspan=14, sticky='nsew', rowspan=2, pady=5, padx=5)

# Binding Ctrl+C for copying selected sentence, menu for clipboard
tree.bind("<Control-c>", copy_to_clipboard)

context_menu = Menu(tree, tearoff=0)
context_menu.add_command(label="Copy", command=lambda: copy_to_clipboard(None))

def show_context_menu(event):
    context_menu.tk_popup(event.x_root, event.y_root)

tree.bind("<Button-3>", show_context_menu)

root.mainloop()
