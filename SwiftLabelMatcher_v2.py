import pandas as pd
import fitz
import subprocess

def printSummary(orders):
    for key, value in orders.items():
        num_not_pirnted = 0
        for i in range(len(value)):
            if not value[i][1]:
                num_not_pirnted += 1
        print(f"{key}: total {len(value)}, not printed {num_not_pirnted}")


pdf_file_path = f'./orders/{input("PDF file name： ")}.pdf'
try:
    pdf_doc = fitz.open(pdf_file_path)
except Exception as e:
    print(e)
    print("could not find the file or the file is not a PDF file")
    quit()

orders = {}

for page_num in range(pdf_doc.page_count):
    page = pdf_doc[page_num]
    text = page.get_text().splitlines()
    # the product id is the second to last or last element in the list
    product_id = text[-2] if any(char.isalpha() for char in text[-2]) else text[-1]
    # the product id could make up of multiple ids, we only consider the first one
    end_index = product_id.find('*')
    product_id = product_id[:end_index if end_index != -1 else len(product_id)]
    orders[product_id].append([page_num + 1, False]) if product_id in orders else orders.update({product_id: [[page_num + 1, False]]})

pdf_doc.close()

while True:
    target_id = input("Product ID: ")
    if target_id.lower() == 'q':
        printSummary(orders)
        print("Quitting...")
        break

    if target_id.lower() == 's':
        printSummary(orders)
        continue

    if target_id in orders:
        for i in range(len(orders[target_id])):
            if not orders[target_id][i][1]:
                page_num = orders[target_id][i][0]
                orders[target_id][i][1] = True
                print(f"Printing page {page_num}, product id {target_id}: {i + 1} of {len(orders[target_id])}")
                # subprocess.run(["lp", "-d", "Zebra_Technologies_ZTC_ZP_450_200dpi", "-o", f"page-ranges={page_num}", pdf_file_path])
                break
        else:
            print("All labels for this product have been printed")
            response = input("Enter 'y' to print again, 'n' to continue: ")
            if response.lower() == 'y':
                for i in range(len(orders[target_id])):
                    orders[target_id][i][1] = False
    else:
        print("Product ID not found")






