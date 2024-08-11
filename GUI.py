import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import fitz
from tkinter import filedialog



class GUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Swift Label Matcher")
        self.root.geometry("800x700")

        self.pdf_file_path = None
        self.pdf_file = None
        self.orders = {}
        self.print_count = 0
        # a global variable to store the selected printer
        self.printer_var = tk.StringVar()

        # the menu bar
        self.menu = tk.Menu(self.root)
        self.file_menu = tk.Menu(self.menu)
        self.file_menu.add_command(label="Open", command=self.open_file)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.root.quit)
        self.menu.add_cascade(label="File", menu=self.file_menu)

        self.printer_menu = tk.Menu(self.menu)
        self.menu.add_cascade(label="Printer", menu=self.printer_menu)
        self.printer_submenu = tk.Menu(self.printer_menu)
        self.printer_menu.add_cascade(label="Select Printer", menu=self.printer_submenu)
        self.display_all_printers()
        self.root.config(menu=self.menu)
        
        # the main content
        self.selected_printer_label_frame = tk.Frame(self.root)
        self.selected_printer_label_frame.pack()
        self.selected_printer_label = tk.Label(self.selected_printer_label_frame, text="Selected Printer: None")
        self.selected_printer_label.grid(row=0, column=0)

        self.file_path_label = tk.Label(self.root, text="PDF File Path: None")
        self.file_path_label.pack()

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(pady=10)
        
        self.user_input_label = tk.Label(self.main_frame, text="Enter Product ID:")
        self.user_input_label.grid(row=0, column=0)

        self.user_input = tk.Entry(self.main_frame)
        self.user_input.grid(row=0, column=1)
        self.user_input.config(width=30)
        self.user_input.bind("<Return>", self.find_print_label)

        self.summary_button = tk.Button(self.main_frame, text="Summary", command=self.show_printing_summary)
        self.summary_button.grid(row=0, column=2, padx=30)

        self.save_button = tk.Button(self.main_frame, text="Save", command=self.save_pdf)
        self.save_button.grid(row=1, column=2)

        self.print_count_label = tk.Label(self.main_frame, text="Print Count: 0")
        self.print_count_label.grid(row=1, column=0)

        self.feedback = tk.Text(self.root, state=tk.DISABLED)
        self.feedback.config(height=40)
        self.feedback.pack()


        self.root.mainloop()

    def open_file(self):
        "Prompt the user to select a PDF file, then extract the product ids from the PDF file."
        self.pdf_file_path = filedialog.askopenfilename(title="Select a PDF file", filetypes=[("PDF files", "*.pdf")])
        self.file_path_label.config(text=f'PDF File: {self.pdf_file_path}')
        self.pdf_file = fitz.open(self.pdf_file_path)
        self.extract_product_ids()

    def extract_product_ids(self):
        "Extract the product ids from the PDF file and store them in the orders dictionary."
        # reset extracted orders and print count if a new PDF file is selected
        if len(self.orders) > 0:
            self.printing_feedback("------------------------------New PDF File Selected-----------------------------")
            self.orders = {}
            self.update_print_count(0)

        for page_num in range(self.pdf_file.page_count):
            page = self.pdf_file[page_num]
            text = page.get_text().splitlines()
            # The product id is the second to last or last element in the list
            # A valid product id must be alphanumeric
            product_id = text[-2] if any(char.isalpha() for char in text[-2]) else text[-1]
            # how many products are there for this product_id in this page
            product_count = 1         
            # One page could have multiple product ids(e.g. SXL023*2+SXL024*1), and we use the first one as a identifier
            end_index = product_id.find('*')
            # if there're more than one product id in the same page, we only take the first one
            if end_index != -1:
                product_count = int(product_id[end_index + 1])
                product_id = product_id[:end_index]                        
            if product_id not in self.orders:                
                self.orders[product_id] = []            
            self.orders[product_id].append([product_count, page_num + 1, False])

        # sort the shipping labels by the number of identifier product present in the page within each product_id category
        for product_id in self.orders:
            self.orders[product_id].sort(key=lambda x: x[0])

        self.pdf_file.close()
        


    def find_print_label(self,event=None):
        "Find the shipping label to print based on the user input."
        if len(self.printer_var.get()) == 0:
            messagebox.showerror("Printer Not Selected", "Please select a printer first")
            return
        if self.pdf_file == None:
            messagebox.showerror("PDF File Not Selected", "Please select a PDF file first")
            return
        target_id = self.user_input.get()
        self.user_input.delete(0, 'end')
        if target_id in self.orders:
            for i in range(len(self.orders[target_id])):
                if not self.orders[target_id][i][2]:
                    page_num = self.orders[target_id][i][1]
                    self.orders[target_id][i][2] = True
                    self.printing_feedback(f"Printing page {page_num}, product id {target_id}: {i + 1} of {len(self.orders[target_id])}")
                    subprocess.run(["lp", "-d", self.printer_var.get(), "-o", f"page-ranges={page_num}", "-o", "print-quality=5", "-o", "orientation-requested=6",self.pdf_file_path])
                    self.update_print_count(self.print_count + 1)
                    break
            # if all labels for the product have been printed. The else block will be executed only if the for loop completes without breaking
            else:
                result = messagebox.askyesno(f"Print Product {target_id} Again?", f"All labels for product {target_id} have been printed. Do you want to print again?")
                if result:
                    for i in range(len(self.orders[target_id])):
                        self.orders[target_id][i][2] = False
                    self.printing_feedback(f"All labels for product {target_id} have been printed. Chose to print again.")
                    self.update_print_count(self.print_count - len(self.orders[target_id]))
                else:
                    self.printing_feedback(f"All labels for product {target_id} have been printed. Chose Not to print again.")
        else:
            self.printing_feedback(f"Product {target_id} not found")


    def printing_feedback(self, message):
        "Display the printing feedback in the feedback text widget."
        self.feedback.config(state=tk.NORMAL)
        self.feedback.insert(tk.END, message + "\n")
        self.feedback.config(state=tk.DISABLED)
        self.feedback.see(tk.END)

    def show_printing_summary(self):
        "Display a summary of the printing in a new window."
        dialog = tk.Toplevel()
        dialog.title("Printing Summary")

        tree = ttk.Treeview(dialog, columns=("Product ID", "Printed", "Total"))
        # Hide the first column (tree column)
        tree.column("#0", width=0, stretch=tk.NO)

        tree.heading("Product ID", text="Product ID")
        tree.heading("Printed", text="Printed")
        tree.heading("Total", text="Total")

        tree.tag_configure("hightlight", background="green")
        for product_id in self.orders:
            total = len(self.orders[product_id])
            printed = sum([1 for label in self.orders[product_id] if label[2]])
            if printed == total:
                tree.insert("", "end", text=product_id, values=(product_id, printed, total), tags=("hightlight",))
            else:
                tree.insert("", "end", text=product_id, values=(product_id, printed, total))

        tree.pack(expand=True, fill="both")

    def get_available_printers(self):
        "Get a list of available printers on the system."
        all_printer_info_list = subprocess.check_output(['lpstat', '-e']).decode('utf-8').splitlines()
        printer_name_list = [line for line in all_printer_info_list]
        return printer_name_list

    def display_all_printers(self):
        "Display all available printers in the printer submenu."
        printer_list = self.get_available_printers()
        for printer in printer_list:
            self.printer_submenu.add_radiobutton(label=printer, variable=self.printer_var, value=printer, command=lambda: self.selected_printer_label.config(text="Selected Printer: " + self.printer_var.get()))

    def update_print_count(self, new_count):
        "Update the print count label to reflect the new print count."
        self.print_count = new_count
        self.print_count_label.config(text=f"Print Count: {self.print_count}")
        
    def sort_pdf_by_product_id(self, path):
        "Sort the PDF file by product id and save the sorted PDF file to the specified path."
        self.pdf_file = fitz.open(self.pdf_file_path)
        saved_pdf_file = fitz.open() # An empty PDF file to store the sorted pages
        sorted_product_id_list = sorted(self.orders.keys())
        for product_id in sorted_product_id_list:
            for _, page_num, _ in self.orders[product_id]:
                # page_num is 1-indexed, and to_page is inclusive
                saved_pdf_file.insert_pdf(self.pdf_file, from_page=page_num - 1, to_page=page_num - 1)
        saved_pdf_file.save(path)
        self.pdf_file.close()
        saved_pdf_file.close()

    def save_pdf(self):
        "Prompt the user to select a path to save the sorted PDF file."
        if self.pdf_file == None:
            messagebox.showerror("PDF File Not Selected", "Please select a PDF file first")
            return
        # "asksaveasfilename" returns an abosolute path to the file
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file_path == "":
            return
        self.sort_pdf_by_product_id(file_path)
GUI()