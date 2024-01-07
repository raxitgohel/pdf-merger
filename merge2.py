import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PyPDF2 import PdfMerger
import re
import customtkinter
import os, sys
from PIL import Image, ImageTk

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")

        # Themed style
        self.root.geometry("700x450")
        self.root.update() 
        self.root.resizable(False, False)
        self.style = ttk.Style()
        self.style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Helvetica', 11), background='#2E2E2E', foreground='white') # Modify the font of the body
        self.style.configure("mystyle.Treeview.Heading", font=('Helvetica', 13,'bold'), background='#2E2E2E', foreground='grey') # Modify the font of the headings
        self.style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Set the background image
        img_path = self.resource_path('logo.png')
        self.bg_image = Image.open(img_path)
        self.width, self.height = self.root.winfo_width(), self.root.winfo_height()
        self.resized_image = self.bg_image.resize((self.height, self.height), Image.LANCZOS)

        # Convert the image to Tkinter PhotoImage
        self.tk_image = ImageTk.PhotoImage(self.resized_image)

        # Create a Canvas widget
        self.canvas = tk.Canvas(self.root,bg='Lavender', width=self.width, height=self.height)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Center and display the dimmed image on the canvas
        self.canvas.create_image(self.width//2, self.height//2, image=self.tk_image)
        
        # To store file paths
        self.pdf_files = []

        # View selected files
        self.tree = ttk.Treeview(self.canvas, style="mystyle.Treeview", selectmode="extended", columns=("Selected Files",), show="headings", height=0)
        self.tree.heading("Selected Files", text="Selected Files")
        self.tree.pack(pady=2)

        # Button to select PDF(s)
        self.select_button = customtkinter.CTkButton(master=self.canvas, text="Select PDF(s)", command=self.select_pdf)
        self.select_button.pack(pady=5)

        # Button to merge PDFs
        self.merge_button = customtkinter.CTkButton(master=self.canvas, text="Merge PDFs", command=self.merge_pdfs)
        self.merge_button.pack(pady=10)

        # Reset button
        self.reset_button = customtkinter.CTkButton(master=self.canvas, text="Reset", command=self.reset)
        self.reset_button.pack(pady=10)

        # Context menu for Treeview
        self.context_menu = tk.Menu(self.canvas, tearoff=0)
        self.context_menu.add_command(label="Remove", command=self.remove_selected_item)

        # Bind right-click event to show context menu
        self.tree.bind("<Button-3>", self.show_context_menu)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
    
    def show_context_menu(self, event):
        # Select item on right-click
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def remove_selected_item(self):
        selected_items = self.tree.selection()

        for item in selected_items:
            values = self.tree.item(item, 'values')
            if values:
                file_name = values[0]
                self.pdf_files = [sublist for sublist in self.pdf_files if sublist[1]!=file_name]
                self.tree.delete(item)
        self.update_tree(calledFromReset=True)

    def extract_name(self, file_path):
        pattern = r"([^\/]+)$"
        match = re.search(pattern, file_path)
        if match:
            return match.group(1)
        else:
            return None

    def select_pdf(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])

        if file_paths:
            for file_path in file_paths:
                file_name = self.extract_name(file_path)
                self.pdf_files.append([file_path, file_name])
                self.update_tree(file_name)

    def update_tree(self, file_name=None, calledFromReset=False):
        # Calculate the desired height dynamically
        tree_height = len(self.pdf_files)
        self.tree.config(height=tree_height)

        self.tree["columns"] = ("Selected Files",)
        self.tree.heading("Selected Files", text="Selected Files")

        if file_name:
            self.tree.insert("", "end", values=(file_name,))
        elif not calledFromReset:
            tk.messagebox.showinfo("Error", "No PDF found")


    def merge_pdfs(self):
        if not self.pdf_files:
            tk.messagebox.showinfo("Error", "No PDFs selected to merge.")
            return
        if len(self.pdf_files) <=1:
            tk.messagebox.showinfo("Error", "Select at least two PDFs.")
            return
        
        try:
            merged_pdf = PdfMerger()
            for pdf_file in self.pdf_files:
                merged_pdf.append(pdf_file[0])

            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

            if save_path:
                merged_pdf.write(save_path)
                merged_pdf.close()
                tk.messagebox.showinfo("Success", "PDFs merged successfully.")
            else:
                tk.messagebox.showerror("Error", "Merge cancelled.")
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")

    
    def reset(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.pdf_files = []
        self.update_tree(calledFromReset=True)


if __name__ == "__main__":
    root = customtkinter.CTk()
    app = PDFMergerApp(root)
    root.mainloop()
