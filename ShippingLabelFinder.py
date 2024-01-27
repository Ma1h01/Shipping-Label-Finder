import pandas as pd
import fitz
import subprocess


pdf_file_path = f'./orders/{input("Enter PDF document name: ")}.pdf'
try:
    pdf_doc = fitz.open(pdf_file_path)
except Exception as e:
    print(e)
    print("Unable to find the PDF document")
    quit()
    
while True:
    target_text = input("Order number: ")
    # The program is excepted to get the input from a external Label Scanner,
    # which prints two extra new line characters for each scan
    # This empty prompt prevents the extra new line character from being used as input to next prompt
    input("")  
    if target_text.lower() == 'q':
        print("Exiting...")
        break
    page_has_found = False
    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        text = page.get_text()

        if target_text in text:
            print(f"Printing page {page_num + 1}")
            page_has_found = True
            subprocess.run(["lp", "-d", "Zebra_Technologies_ZTC_ZP_450_200dpi", "-o", f"page-ranges={page_num+1}", pdf_file_path])
            break
    if not page_has_found:
        print("Label not found")

pdf_doc.close()

