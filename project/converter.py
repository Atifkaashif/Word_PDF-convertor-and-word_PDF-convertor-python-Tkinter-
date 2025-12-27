import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from docx import Document
from docx2pdf import convert as docx2pdf_convert
import os
import platform

# -------- Functions --------
def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_path_var.set(file_path)
        progress_bar_pdf['value'] = 0

def browse_word():
    file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if file_path:
        word_path_var.set(file_path)
        progress_bar_word['value'] = 0

def refresh_pdf():
    pdf_path_var.set("")
    progress_bar_pdf['value'] = 0

def refresh_word():
    word_path_var.set("")
    progress_bar_word['value'] = 0

def convert_pdf_to_word():
    pdf_path = pdf_path_var.get()
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                             filetypes=[("Word Files", "*.docx")])
    if not save_path:
        return

    try:
        doc = Document()
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)
                # Update progress bar
                progress = ((i + 1) / total_pages) * 100
                progress_bar_pdf['value'] = progress
                root.update_idletasks()

        doc.save(save_path)
        messagebox.showinfo("Success", f"✅ PDF successfully converted to Word!\nSaved as:\n{save_path}")
        progress_bar_pdf['value'] = 0

    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert PDF:\n{str(e)}")
        progress_bar_pdf['value'] = 0

def convert_word_to_pdf():
    word_path = word_path_var.get()
    if not word_path:
        messagebox.showerror("Error", "Please select a Word file")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF Files", "*.pdf")])
    if not save_path:
        return

    if platform.system() != "Windows":
        messagebox.showwarning("Platform Warning", "Word→PDF conversion works reliably only on Windows with MS Word installed.")
        return

    try:
        progress_bar_word['value'] = 50
        root.update_idletasks()

        docx2pdf_convert(word_path, save_path)

        progress_bar_word['value'] = 100
        root.update_idletasks()

        messagebox.showinfo("Success", f"✅ Word successfully converted to PDF!\nSaved as:\n{save_path}")
        progress_bar_word['value'] = 0

    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert Word to PDF:\n{str(e)}")
        progress_bar_word['value'] = 0

# -------- Button Hover Effects --------
def on_enter(e, button, hover_color):
    button['background'] = hover_color

def on_leave(e, button, default_color):
    button['background'] = default_color

# -------- Tkinter GUI --------
root = tk.Tk()
root.title("PDF ⇄ Word Converter")
root.geometry("650x400")
root.resizable(False, False)
root.configure(bg="#f0f8ff")

tab_control = ttk.Notebook(root)

# ---------------- Tab 1: PDF → Word ----------------
tab1 = tk.Frame(tab_control, bg="#f0f8ff")
tab_control.add(tab1, text="PDF → Word")

pdf_path_var = tk.StringVar()

tk.Label(tab1, text="Select PDF File:", font=("Helvetica", 14), bg="#f0f8ff").pack(pady=10)
entry_frame_pdf = tk.Frame(tab1, bg="#f0f8ff")
entry_frame_pdf.pack(pady=5)

pdf_entry = tk.Entry(entry_frame_pdf, textvariable=pdf_path_var, width=45, font=("Arial", 12))
pdf_entry.pack(side=tk.LEFT, padx=5)

browse_pdf_btn = tk.Button(entry_frame_pdf, text="Browse", command=browse_pdf, bg="#1e90ff", fg="white",
                           font=("Arial", 11, "bold"), relief="raised", bd=3)
browse_pdf_btn.pack(side=tk.LEFT, padx=5)
browse_pdf_btn.bind("<Enter>", lambda e: on_enter(e, browse_pdf_btn, "#63b3ed"))
browse_pdf_btn.bind("<Leave>", lambda e: on_leave(e, browse_pdf_btn, "#1e90ff"))

refresh_pdf_btn = tk.Button(entry_frame_pdf, text="Refresh", command=refresh_pdf, bg="#ff6347", fg="white",
                            font=("Arial", 11, "bold"), relief="raised", bd=3)
refresh_pdf_btn.pack(side=tk.LEFT, padx=5)
refresh_pdf_btn.bind("<Enter>", lambda e: on_enter(e, refresh_pdf_btn, "#ff7f50"))
refresh_pdf_btn.bind("<Leave>", lambda e: on_leave(e, refresh_pdf_btn, "#ff6347"))

convert_pdf_btn = tk.Button(tab1, text="Convert to Word", command=convert_pdf_to_word, bg="#28a745",
                            fg="white", font=("Arial", 13, "bold"), relief="raised", bd=4, width=25, height=2)
convert_pdf_btn.pack(pady=20)
convert_pdf_btn.bind("<Enter>", lambda e: on_enter(e, convert_pdf_btn, "#45c767"))
convert_pdf_btn.bind("<Leave>", lambda e: on_leave(e, convert_pdf_btn, "#28a745"))

progress_bar_pdf = ttk.Progressbar(tab1, orient='horizontal', length=500, mode='determinate')
progress_bar_pdf.pack(pady=10)

# ---------------- Tab 2: Word → PDF ----------------
tab2 = tk.Frame(tab_control, bg="#f0f8ff")
tab_control.add(tab2, text="Word → PDF")

word_path_var = tk.StringVar()

tk.Label(tab2, text="Select Word File:", font=("Helvetica", 14), bg="#f0f8ff").pack(pady=10)
entry_frame_word = tk.Frame(tab2, bg="#f0f8ff")
entry_frame_word.pack(pady=5)

word_entry = tk.Entry(entry_frame_word, textvariable=word_path_var, width=45, font=("Arial", 12))
word_entry.pack(side=tk.LEFT, padx=5)

browse_word_btn = tk.Button(entry_frame_word, text="Browse", command=browse_word, bg="#1e90ff", fg="white",
                            font=("Arial", 11, "bold"), relief="raised", bd=3)
browse_word_btn.pack(side=tk.LEFT, padx=5)
browse_word_btn.bind("<Enter>", lambda e: on_enter(e, browse_word_btn, "#63b3ed"))
browse_word_btn.bind("<Leave>", lambda e: on_leave(e, browse_word_btn, "#1e90ff"))

refresh_word_btn = tk.Button(entry_frame_word, text="Refresh", command=refresh_word, bg="#ff6347", fg="white",
                             font=("Arial", 11, "bold"), relief="raised", bd=3)
refresh_word_btn.pack(side=tk.LEFT, padx=5)
refresh_word_btn.bind("<Enter>", lambda e: on_enter(e, refresh_word_btn, "#ff7f50"))
refresh_word_btn.bind("<Leave>", lambda e: on_leave(e, refresh_word_btn, "#ff6347"))

convert_word_btn = tk.Button(tab2, text="Convert to PDF", command=convert_word_to_pdf, bg="#28a745",
                             fg="white", font=("Arial", 13, "bold"), relief="raised", bd=4, width=25, height=2)
convert_word_btn.pack(pady=20)
convert_word_btn.bind("<Enter>", lambda e: on_enter(e, convert_word_btn, "#45c767"))
convert_word_btn.bind("<Leave>", lambda e: on_leave(e, convert_word_btn, "#28a745"))

progress_bar_word = ttk.Progressbar(tab2, orient='horizontal', length=500, mode='determinate')
progress_bar_word.pack(pady=10)

# Add tabs to root
tab_control.pack(expand=1, fill="both")

root.mainloop()
