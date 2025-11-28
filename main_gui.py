import os
import sys  # Added sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from check_docx_engine import generate_html_report, parse_document_sections

# --- Helper to get resource path ---
def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- Helper to load version ---
def load_version():
    """ Load version from version.txt if it exists, else return 'Dev Mode' """
    version_file = get_resource_path("version.txt")
    if os.path.exists(version_file):
        try:
            with open(version_file, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass
    return "Dev Mode"

APP_VERSION = load_version()

class AuditorApp:
    def __init__(self, root):
        self.root = root
        self.style = ttk.Style(theme="litera") 
        # Update Title with Version
        self.root.title(f"Docx Bilingual Auditor (v{APP_VERSION})")
        self.root.geometry("1080x768")
        
        self.chi_path_var = tk.StringVar()
        self.eng_path_var = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        main_container = ttk.Frame(self.root, padding=30)
        main_container.pack(fill=BOTH, expand=YES)

        # Header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=X, pady=(0, 25))
        
        # Display Version in Title Label as well
        title_lbl = ttk.Label(
            header_frame, 
            text=f"Docx Bilingual Auditor v{APP_VERSION}", 
            font=("Helvetica", 24, "bold"),
            bootstyle=PRIMARY
        )
        title_lbl.pack(side=LEFT)
        
        subtitle_lbl = ttk.Label(
            header_frame,
            text="Compare Bold & Underline styles across documents.",
            font=("Helvetica", 10),
            bootstyle=SECONDARY
        )
        subtitle_lbl.pack(side=LEFT, padx=10, pady=(12, 0))

        # Input Frame
        files_frame = ttk.Labelframe(main_container, text=" 📄 Source Documents ", padding=20, bootstyle=INFO)
        files_frame.pack(fill=X, pady=10)
        files_frame.columnconfigure(1, weight=1)

        ttk.Label(files_frame, text="Chinese Version:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky=W, pady=10)
        chi_entry = ttk.Entry(files_frame, textvariable=self.chi_path_var, font=("Helvetica", 10))
        chi_entry.grid(row=0, column=1, sticky=EW, padx=15, pady=10)
        ttk.Button(files_frame, text="Browse...", command=self.select_chi_file, bootstyle="info-outline").grid(row=0, column=2, pady=10)

        ttk.Label(files_frame, text="English Version:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky=W, pady=10)
        eng_entry = ttk.Entry(files_frame, textvariable=self.eng_path_var, font=("Helvetica", 10))
        eng_entry.grid(row=1, column=1, sticky=EW, padx=15, pady=10)
        ttk.Button(files_frame, text="Browse...", command=self.select_eng_file, bootstyle="info-outline").grid(row=1, column=2, pady=10)

        # Action Buttons
        action_frame = ttk.Frame(main_container)
        action_frame.pack(fill=X, pady=25)

        self.btn_run = ttk.Button(
            action_frame, 
            text="🚀 Start Audit Analysis", 
            command=self.start_process,
            bootstyle="success",
            width=25,
            padding=10
        )
        self.btn_run.pack(side=LEFT, padx=(0, 15))
        
        ttk.Button(action_frame, text="Exit Application", command=self.root.quit, bootstyle="secondary-outline", padding=10).pack(side=LEFT)

        # Log Area
        log_label = ttk.Label(main_container, text="Execution Log:", font=("Helvetica", 10, "bold"), bootstyle=SECONDARY)
        log_label.pack(fill=X, pady=(10, 5))

        self.log_text = ScrolledText(
            main_container, 
            height=12, 
            state='disabled', 
            font=("Consolas", 10),
            bg="#f8f9fa",  # Light grey background
            fg="#343a40",  # Dark grey text
            relief="flat", 
            highlightthickness=1, 
            highlightbackground="#dee2e6"
        )
        self.log_text.pack(fill=BOTH, expand=YES)

    def log(self, message, level="info"):
        """Write message to GUI log area"""
        self.log_text.config(state='normal')
        
        tag_name = "normal"
        if "Error" in message or level == "error":
            tag_name = "error"
            self.log_text.tag_config("error", foreground="#dc3545")
        elif "Warning" in message:
             tag_name = "warning"
             self.log_text.tag_config("warning", foreground="#ffc107")
        elif ">>>" in message or "Success" in message:
             tag_name = "highlight"
             self.log_text.tag_config("highlight", foreground="#0d6efd", font=("Consolas", 10, "bold"))

        self.log_text.insert(tk.END, message + "\n", tag_name)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def select_chi_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if filename:
            self.chi_path_var.set(filename)
            self.log(f"Selected CH file: {os.path.basename(filename)}")

    def select_eng_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if filename:
            self.eng_path_var.set(filename)
            self.log(f"Selected EN file: {os.path.basename(filename)}")

    def start_process(self):
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        chi_path = self.chi_path_var.get()
        eng_path = self.eng_path_var.get()

        if not chi_path or not eng_path:
            self.root.after(0, lambda: messagebox.showwarning("Action Required", "Please select both Chinese and English documents first."))
            return

        self.root.after(0, lambda: self.btn_run.config(state="disabled", text="Analyzing... Please Wait"))
        
        def clear_log():
            self.log_text.config(state='normal')
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state='disabled')
        self.root.after(0, clear_log)

        self.root.after(0, lambda: self.log(f">>> Starting Bilingual Audit Process (v{APP_VERSION})...", "highlight"))
        self.root.after(0, lambda: self.log("-" * 50))

        try:
            # log_func wrapper for thread safety
            def thread_safe_log(msg):
                self.root.after(0, lambda: self.log(msg))

            # 1. Extract
            self.root.after(0, lambda: self.log(f"Reading Chinese Doc: {os.path.basename(chi_path)}"))
            chi_pattern = r"^[甲乙丙丁戊己庚辛壬癸(（].*部\s*[：:]"
            
            sections_chi = parse_document_sections(chi_path, "頁碼", chi_pattern, thread_safe_log)

            self.root.after(0, lambda: self.log(f"Reading English Doc: {os.path.basename(eng_path)}"))
            eng_pattern = r"^Part.*[：:]"
            
            sections_eng = parse_document_sections(eng_path, "Page", eng_pattern, thread_safe_log)

            # 2. Generate Report
            output_dir = os.path.dirname(chi_path)
            output_path = os.path.join(output_dir, "Bilingual_Audit_Report.html")
            
            self.root.after(0, lambda: self.log("-" * 50))
            self.root.after(0, lambda: self.log(f"Generating HTML Report...", "highlight"))
            
            generate_html_report(sections_chi, sections_eng, output_path)

            self.root.after(0, lambda: self.log("✅ SUCCESS! Report generated successfully.", "highlight"))
            self.root.after(0, lambda: self.log(f"Location: {output_path}"))
            
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Audit complete!\n\nReport saved to:\n{output_path}"))

        except Exception as e:
            self.root.after(0, lambda: self.log(f"❌ Critical Error: {str(e)}", "error"))
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred:\n{str(e)}"))
        
        finally:
            self.root.after(0, lambda: self.btn_run.config(state="normal", text="🚀 Start Audit Analysis"))

if __name__ == "__main__":
    root = ttk.Window(themename="litera")
    app = AuditorApp(root)
    root.mainloop()