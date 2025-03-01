# Uses Pdftk cmd line 
import os
import re
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to select input PDF
def select_pdf():
    filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if filepath:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, filepath)

# Function to add a new chapter entry
def add_chapter():
    chapter_title = chapter_entry.get().strip()
    page_number = page_entry.get().strip()

    if not chapter_title or not page_number.isdigit():
        messagebox.showerror("Error", "Enter a valid chapter title and page number!")
        return

    chapter_listbox.insert(tk.END, f"{chapter_title} - Page {page_number}")
    chapters.append((re.sub(r"[\r\n]",' ',chapter_title), int(page_number)))

    chapter_entry.delete(0, tk.END)
    page_entry.delete(0, tk.END)

# Function to add bookmarks and save the new PDF
def add_bookmarks():
    pdf_path = pdf_entry.get()
    if not os.path.exists(pdf_path):
        messagebox.showerror("Error", "Invalid PDF file!")
        return

    output_path = pdf_path.replace(".pdf", "_bookmarked.pdf")
    bookmarks_file = "bookmarks.txt"

    # Create the bookmarks.txt file
    with open(bookmarks_file, "w") as f:
        for title, page in chapters:
            f.write("BookmarkBegin\n")
            f.write(f"BookmarkTitle: {title}\n")
            f.write("BookmarkLevel: 1\n")
            f.write(f"BookmarkPageNumber: {page}\n\n")

    # Run pdftk command
    cmd = f'pdftk "{pdf_path}" update_info "{bookmarks_file}" output "{output_path}"'
    subprocess.run(cmd, shell=True)

    messagebox.showinfo("Success", f"Bookmarks added!\nSaved as {output_path}")

# GUI Setup
root = tk.Tk()
root.title("PDF Bookmark Adder")
root.geometry("450x400")

# PDF Selection
tk.Label(root, text="Select PDF:").pack(pady=5)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(pady=5)
tk.Button(root, text="Browse", command=select_pdf).pack(pady=5)

# Chapter Input
tk.Label(root, text="Chapter Title:").pack()
chapter_entry = tk.Entry(root, width=40)
chapter_entry.pack()

tk.Label(root, text="Page Number:").pack()
page_entry = tk.Entry(root, width=10)
page_entry.pack()

tk.Button(root, text="Add Chapter", command=add_chapter, bg="lightgreen").pack(pady=5)

# Chapter List
tk.Label(root, text="Chapters:").pack()
chapter_listbox = tk.Listbox(root, width=50, height=5)
chapter_listbox.pack()

# Add Bookmarks Button
tk.Button(root, text="Add Bookmarks", command=add_bookmarks, bg="lightblue").pack(pady=10)

# Store chapters
chapters = []

root.mainloop()
