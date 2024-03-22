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
        self.root.geometry("800x600")

        # set the menu
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
        # a global variable to store the selected printer
        self.printer_var = tk.StringVar()
        self.printer_var.set(None)
        self.display_all_printers()
        self.root.config(menu=self.menu)
        
        self.selected_printer_label_frame = tk.Frame(self.root)
        self.selected_printer_label_frame.pack()
        self.selected_printer_label = tk.Label(self.selected_printer_label_frame, text="Selected Printer:")
        self.selected_printer_label.grid(row=0, column=0)
        self.selected_printer_name_label = tk.Label(self.selected_printer_label_frame, textvariable=self.printer_var)
        self.selected_printer_name_label.grid(row=0, column=1)

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

        self.feedback = tk.Text(self.root, state=tk.DISABLED)
        self.feedback.config(height=40)
        self.feedback.pack()

        self.pdf_file_path = None
        self.pdf_file = None
        self.orders = {}


        self.root.mainloop()

    def open_file(self):
        self.pdf_file_path = filedialog.askopenfilename(title="Select a PDF file", filetypes=[("PDF files", "*.pdf")])
        self.file_path_label.config(text=f'PDF File: {self.pdf_file_path}')
        self.pdf_file = fitz.open(self.pdf_file_path)
        self.extract_product_ids()

    def extract_product_ids(self):
        if len(self.orders) > 0:
            self.printing_feedback("------------------------------New PDF File Selected-----------------------------")
            self.orders = {}
        for page_num in range(self.pdf_file.page_count):
            page = self.pdf_file[page_num]
            text = page.get_text().splitlines()
            # the product id is the second to last or last element in the list
            product_id = text[-2] if any(char.isalpha() for char in text[-2]) else text[-1]
            # the product id could make up of multiple ids, we only consider the first one
            end_index = product_id.find('*')
            product_id = product_id[:end_index if end_index != -1 else len(product_id)]
            self.orders[product_id].append([page_num + 1, False]) if product_id in self.orders else self.orders.update({product_id: [[page_num + 1, False]]})
        self.pdf_file.close()
        


    def find_print_label(self,event=None):
        if self.printer_var.get() == 'None':
            messagebox.showerror("Printer Not Selected", "Please select a printer first")
            return
        if self.pdf_file == None:
            messagebox.showerror("PDF File Not Selected", "Please select a PDF file first")
            return
        target_id = self.user_input.get()
        self.user_input.delete(0, 'end')
        if target_id in self.orders:
            for i in range(len(self.orders[target_id])):
                if not self.orders[target_id][i][1]:
                    page_num = self.orders[target_id][i][0]
                    self.orders[target_id][i][1] = True
                    self.printing_feedback(f"Printing page {page_num}, product id {target_id}: {i + 1} of {len(self.orders[target_id])}")
                    subprocess.run(["lp", "-d", self.printer_var.get(), "-o", f"page-ranges={page_num}", "-o", "print-quality=5", "-o", "orientation-requested=6",self.pdf_file_path])
                    break
            # if all labels for the product have been printed. The else block will be executed only if the for loop completes without breaking
            else:
                result = messagebox.askyesno(f"Print Product {target_id} Again?", f"All labels for product {target_id} have been printed. Do you want to print again?")
                if result:
                    for i in range(len(self.orders[target_id])):
                        self.orders[target_id][i][1] = False
                    self.printing_feedback(f"All labels for product {target_id} have been printed. Printing again...")
                else:
                    self.printing_feedback("All labels for this product have been printed. Will Not print again.")

        else:
            self.printing_feedback(f"Product {target_id} not found")


    def printing_feedback(self, message):
        self.feedback.config(state=tk.NORMAL)
        self.feedback.insert(tk.END, message + "\n")
        self.feedback.config(state=tk.DISABLED)
        self.feedback.see(tk.END)

    def show_printing_summary(self):
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
            printed = sum([1 for label in self.orders[product_id] if label[1]])
            if printed == total:
                tree.insert("", "end", text=product_id, values=(product_id, printed, total), tags=("hightlight",))
            else:
                tree.insert("", "end", text=product_id, values=(product_id, printed, total))

        tree.pack(expand=True, fill="both")

    def get_available_printers(self):
        all_printer_info_list = subprocess.check_output(['lpstat', '-e']).decode('utf-8').splitlines()
        printer_name_list = [line for line in all_printer_info_list]
        return printer_name_list

    def display_all_printers(self):
        printer_list = self.get_available_printers()
        for printer in printer_list:
            self.printer_submenu.add_radiobutton(label=printer, variable=self.printer_var, value=printer)

GUI()