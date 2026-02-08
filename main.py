import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from converter import convert_pdf_to_pptx
import threading
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to PPT Converter")
        self.root.geometry("500x250")
        self.root.resizable(False, False)

        self.pdf_path = None

        # Style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#ccc")

        # Main Frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text="PDF to PowerPoint Converter", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # File Selection Frame
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)

        self.file_label = ttk.Label(file_frame, text="No file selected", width=40, anchor="w", relief="sunken", padding=(5, 5))
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        select_btn = ttk.Button(file_frame, text="Select PDF", command=self.select_file)
        select_btn.pack(side=tk.RIGHT)

        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=20)

        # Convert Button
        self.convert_btn = ttk.Button(main_frame, text="Convert to PPT", command=self.start_conversion, state=tk.DISABLED)
        self.convert_btn.pack(fill=tk.X, pady=5)

        # Status Label
        self.status_label = ttk.Label(main_frame, text="Ready", foreground="gray")
        self.status_label.pack(pady=5)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.convert_btn.config(state=tk.NORMAL)
            self.status_label.config(text="Ready", foreground="gray")
            self.progress_var.set(0)

    def start_conversion(self):
        if not self.pdf_path:
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint Presentation", "*.pptx")],
            initialfile=os.path.splitext(os.path.basename(self.pdf_path))[0] + ".pptx"
        )

        if not save_path:
            return

        self.convert_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Converting...", foreground="blue")
        
        # Run conversion in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.run_conversion, args=(self.pdf_path, save_path))
        thread.start()

    def run_conversion(self, pdf_path, pptx_path):
        def update_progress(percent):
            self.root.after(0, lambda: self.progress_var.set(percent))

        success, message = convert_pdf_to_pptx(pdf_path, pptx_path, update_progress)
        
        self.root.after(0, lambda: self.conversion_finished(success, message))

    def conversion_finished(self, success, message):
        self.convert_btn.config(state=tk.NORMAL)
        if success:
            self.status_label.config(text="Done!", foreground="green")
            messagebox.showinfo("Success", message)
            # Reset
            self.progress_var.set(0)
            self.status_label.config(text="Ready", foreground="gray")
        else:
            self.status_label.config(text="Error", foreground="red")
            messagebox.showerror("Error", f"An error occurred:\n{message}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
