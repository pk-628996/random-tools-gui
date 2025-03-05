import os
import sys
import re
import string
import subprocess
from sys import platform
import tkinter as tk
from tkinter import filedialog, messagebox


# secure Pdfs
# Remove passwords



#pdftk_path=os.path.join(sys._MEIPASS,"pdftk.exe")
pdftk_path = 'pdftk'

def pg_count(pdf):
    if os.path.exists(pdf):
        try:
            num_str = subprocess.run([pdftk_path, pdf, 'dump_data'], capture_output=True, text=True,creationflags=subprocess.CREATE_NO_WINDOW)
            for line in num_str.stdout.splitlines():
                if line.startswith('NumberOfPages'):
                    return re.sub(r"\D", "", line)
        except Exception as e:
            print(f"Error getting page count: {e}")
    return "0"

                    
all_labels = list(string.ascii_uppercase)
all_labels += [a + b for a in string.ascii_uppercase for b in string.ascii_uppercase]
all_labels += [a + b + c for a in string.ascii_uppercase for b in string.ascii_uppercase for c in string.ascii_uppercase]


def is_valid_pdftk_range(range_str, total_pages):
    """
    Validates a given pdftk page range string against the total number of pages.
    """
    range_pattern = re.compile(r'^\d+(-\d+)?(,(\d+(-\d+)?))*$')  # Valid patterns: 1, 2-5, 3,7-9
    even_odd_pattern = re.compile(r'^\d+-\d+(even|odd)$')  # Special case: 1-10even, 3-9odd

    # Parse and validate page numbers
    parts = range_str.split(',')
    parts = [part.strip() for part in parts]
    for part in parts:
        if not (range_pattern.match(part) or even_odd_pattern.match(part)):
            return False
        if 'even' in part or 'odd' in part:
            start, end = map(int, re.findall(r'\d+', part))
            if start < 1 or end > total_pages or start > end:
                return False
        elif '-' in part:
            start, end = map(int, part.split('-'))
            if start < 1 or end > total_pages or start > end:
                return False
        else:
            page = int(part)
            if page < 1 or page > total_pages:
                return False

    return True

def handle_ranges(label,range_str):
    parts = range_str.split(',')
    parts = [part.strip() for part in parts]
    range_handle = []
    for part in parts:
        range_handle.append(f'{label}{part}')
    return range_handle




class PdfMerger(tk.Frame):
    '''Handles merging of pdfs.'''
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(bg='lightyellow')
        self.pack(fill="both", expand=True)
        self.intro = '''First, add your input PDFs. Order them via drag-and-drop. Their pages will be copied and merged to create your new PDF. List the pages to copy using a mix of page numbers (e.g. 1,3,7) and page ranges (e.g. 1-3, 12-4, 20-8). You can also append "even" or "odd" (e.g. 1-10even). You can list the same page numbers or ranges more than once. You can also list the same PDF document more than once.'''
        self.intro_tbox = tk.Text(self,wrap='word', font=('Arial',10),height=4,width=116,state='disabled',bg='lightgreen')
        self.intro_tbox.grid(row=0,column=0,columnspan=3,pady=10,padx=10)
        self.intro_tbox.config(state='normal') 
        self.intro_tbox.tag_configure('center',justify='center',lmargin1=20,lmargin2=20)
        self.intro_tbox.insert('1.0',self.intro) 
        self.intro_tbox.tag_add('center','1.0','end')
        self.intro_tbox.config(state='disabled')
        tk.Label(self, text="Selected pdfs:").grid(row=1,column=0,pady=0,padx=0)
        self.file_list = tk.Listbox(self, width=75, height=15)
        self.file_list.grid(row=2,column=0,pady=0,padx=0)
        self.file_list.bind('<Button-1>',self._on_drag_start)
        self.file_list.bind('<B1-Motion>',self._on_drag_motion)
        self.file_list.bind('<ButtonRelease-1>',self._on_drag_drop)
        tk.Label(self, text="Page Count").grid(row=1,column=1,pady=0,padx=0)
        self.num_list = tk.Listbox(self, width=10, height=15)
        self.num_list.grid(row=2,column=1,pady=0,padx=0)
        tk.Label(self, text="Range - Double click to change").grid(row=1,column=2,pady=2,padx=2)
        self.range_list = tk.Listbox(self, width=40, height=15)
        self.range_list.grid(row=2,column=2,pady=0,padx=0)
        self.range_list.bind('<Double-Button-1>',self.edit_item)
        tk.Button(self, text="Add PDFs", command=self.add_pdfs).grid(row=3,column=0,pady=5)
        tk.Button(self, text="Merge PDFs", command=self.merge_pdfs, bg="lightgreen").grid(row=3,column=1,pady=10)
    def _on_drag_start(self,event):
        widget = event.widget
        self.dragged_file_index = widget.nearest(event.y)
        widget.selection_clear(0,tk.END)
        widget.selection_set(self.dragged_file_index)
    def _on_drag_motion(self,event):
        widget = event.widget
        index = widget.nearest(event.y)
        widget.activate(index)
        if event.y < 20:
            widget.yview_scroll(-1,'units')
        elif event.y > widget.winfo_height() -20:
            widget.yview_scroll(1,'units')
    def _on_drag_drop(self,event):
        widget = event.widget
        self.dropped_at_index = widget.nearest(event.y)
        if self.dropped_at_index != self.dragged_file_index:
            text = widget.get(self.dragged_file_index)
            ntext = self.num_list.get(self.dragged_file_index)
            rtext = self.range_list.get(self.dragged_file_index)
            self.num_list.delete(self.dragged_file_index)
            self.num_list.insert(self.dropped_at_index,ntext)
            self.range_list.delete(self.dragged_file_index)
            self.range_list.insert(self.dropped_at_index,rtext)
            widget.delete(self.dragged_file_index)
            widget.insert(self.dropped_at_index,text)
            widget.selection_clear(0,tk.END)
            widget.selection_set(self.dropped_at_index)
    def add_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for file in files:
            self.file_list.insert(tk.END, file)
            pg_cnt = pg_count(file)
            self.num_list.insert(tk.END,pg_cnt)
            self.range_list.insert(tk.END,f'1-{pg_cnt}')
    def merge_pdfs(self):
        files = list(self.file_list.get(0, tk.END))
        ranges = list(self.range_list.get(0, tk.END))
        if not files:
            messagebox.showerror("Error", "No PDFs selected!")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_path:
            labels = {}
            new_list = []
            for i,fnm in enumerate(files):
                label = all_labels[i]
                labels[label] = ranges[i]
                new_list.append(f'{label}={fnm}')
            cat_list = []
            for label,rang in labels.items():
                cat_list += handle_ranges(label,rang)
            cmd = [pdftk_path,*new_list,'cat',*cat_list,'output',output_path]
            print(cmd)
            f=subprocess.run(cmd,capture_output=True,text=True,creationflags=subprocess.CREATE_NO_WINDOW)
            if f.returncode == 0:
                messagebox.showinfo("Success", f"PDFs Merged!\nSaved as {output_path}")
            else:
                messagebox.showerror('Error',f.stderr)
    def edit_item(self,event):
        try:
            self.selected_index = self.range_list.curselection()[0]
            self.original_txt = self.range_list.get(self.selected_index)
            self.original_pg_cnt = int(self.num_list.get(self.selected_index))
            x,y,width,height = self.range_list.bbox(self.selected_index)
            width=self.range_list.winfo_width()
            self.rentry = tk.Entry(self.range_list)
            self.rentry.insert(0,self.original_txt)
            self.rentry.bind('<Return>',self.save_edit)
            self.rentry.bind('<FocusOut>',self.save_edit)
            self.rentry.place(x=x,y=y,width=width,height=height)
            self.rentry.focus_set()
        except IndexError:
            pass
    def save_edit(self,event):
        new_text = self.rentry.get()
        if not is_valid_pdftk_range(new_text,self.original_pg_cnt):
            messagebox.showerror("Error", "One or More pdfs has invalid range.")
            self.rentry.destroy()
            return
        self.range_list.delete(self.selected_index)
        self.range_list.insert(self.selected_index,new_text)
        self.rentry.destroy()
        
class BookmarkManager(tk.Frame):
    """Handles adding bookmarks to a PDF using pdftk."""
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(bg='lightyellow')
        self.chapters = []
        self.pack(fill="both", expand=True)
        self.slabel = tk.Label(self,text='Select Pdf:')
        self.slabel.pack(pady=2)
        self.sentry = tk.Entry(self,width=60)
        self.sentry.pack(pady=2)
        self.bbutton = tk.Button(self, text="Browse", command=self.select_pdf,bg='gold')
        self.bbutton.pack(pady=5)
        self.tlabel = tk.Label(self,text='Enter Chapter Title')
        self.tlabel.pack()
        self.tentry = tk.Entry(self,width=80)
        self.tentry.pack()
        self.nlabel = tk.Label(self,text='Page Number')
        self.nlabel.pack()
        self.nentry = tk.Entry(self,width=80)
        self.nentry.pack()
        self.addbtn = tk.Button(self, text="Add Chapter", command=self.add_chapter, bg="lightgreen")
        self.addbtn.pack(pady=5)
        self.chstxt = tk.Label(self,text='Chapters')
        self.chstxt.pack()
        self.chapter_listbox = tk.Listbox(self, width=80, height=10)
        self.chapter_listbox.pack()
        self.bkbutton = tk.Button(self, text="Add Bookmarks", command=self.add_bookmarks, bg="lightblue")
        self.bkbutton.pack(pady=10)
    def add_chapter(self):
        chapter_title = self.tentry.get().strip()
        page_number = self.nentry.get().strip()
        if not chapter_title or not page_number.isdigit():
            messagebox.showerror("Error", "Enter a valid chapter title and page number!")
            return
        self.chapter_listbox.insert(tk.END, f"{chapter_title} - Page {page_number}")
        self.chapters.append((re.sub(r"[\r\n]",' ',chapter_title), int(page_number)))
        self.tentry.delete(0, tk.END)
        self.nentry.delete(0, tk.END)
    def select_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.sentry.delete(0, tk.END)
            self.sentry.insert(0, filepath)
    def add_bookmarks(self):
        pdf_path = self.sentry.get()
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Invalid PDF file!")
            return
        output_path = pdf_path[:-4] + "_bookmarked.pdf"
        bookmarks_file = "bookmarks.txt"
        # Create the bookmarks.txt file
        with open(bookmarks_file, "w") as f:
            for title, page in self.chapters:
                f.write("BookmarkBegin\n")
                f.write(f"BookmarkTitle: {title}\n")
                f.write("BookmarkLevel: 1\n")
                f.write(f"BookmarkPageNumber: {page}\n\n")
        # Run pdftk command
        cmd = [pdftk_path,pdf_path,'update_info',bookmarks_file,'output',output_path]
        f = subprocess.run(cmd,capture_output=True,text=True,creationflags=subprocess.CREATE_NO_WINDOW)
        if f.returncode == 0:
            os.remove(bookmarks_file)
            messagebox.showinfo("Success", f"Bookmarks added!\nSaved as {output_path}")
        else:
            messagebox.showerror('Error',str(f.stderr))




class MainApp(tk.Tk):
    """Main application to switch between different PDF functions."""
    def __init__(self):
        super().__init__()
        self.title("PDF Tool")
        self.geometry("1000x500")
        #self.configure(bg='lightyellow')
        # Menu
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        menu_options = tk.Menu(menubar, tearoff=0)
        help_menu = tk.Menu(menubar,tearoff=0)
        menubar.add_cascade(label="Options", menu=menu_options)
        menubar.add_cascade(label='Help',menu=help_menu)
        menu_options.add_command(label="Merge PDFs", command=self.show_pdf_merger)
        menu_options.add_command(label="Add Bookmark", command=self.show_book_marker)
        menu_options.add_separator()
        menu_options.add_command(label="Exit", command=self.quit)
        help_menu.add_command(label='About',command=self.show_about_info)
        help_menu.add_separator()
        help_menu.add_command(label='Exit',command=self.quit)
        self.current_frame = None
        self.show_pdf_merger()
    def show_frame(self, frame_class):
        """Destroys current frame and replaces it with a new one."""
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = frame_class(self)
    def show_pdf_merger(self):
        self.show_frame(PdfMerger)
    def show_book_marker(self):
        self.show_frame(BookmarkManager)
    def show_about_info(self):
        messagebox.showinfo('About','nothing to say.')


if __name__=='__main__':
    if platform == 'win32':
        print('Windows')
        app = MainApp()
        app.mainloop()
