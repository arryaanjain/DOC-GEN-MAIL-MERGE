import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class FileSelector:
    def __init__(self, master, on_file_selected):
        self.master = master
        self.on_file_selected = on_file_selected
        self.frame = tk.Frame(master)
        self.frame.pack(pady=20)

        self.label = tk.Label(self.frame, text="Select a DOCX file:")
        self.label.pack()

        self.select_button = tk.Button(self.frame, text="Browse", command=self.select_file)
        self.select_button.pack()

        self.file_path_label = tk.Label(self.frame, text="")
        self.file_path_label.pack()

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("DOCX files", "*.docx")])
        if file_path:
            self.file_path_label.config(text=file_path)
            self.on_file_selected(file_path)
        else:
            messagebox.showwarning("Warning", "No file selected.")