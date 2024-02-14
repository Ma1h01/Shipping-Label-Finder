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

product_ids = {}
for page_num in range(pdf_doc.page_count):
    page = pdf_doc[page_num]  #  SHG162
    id = page.get_text().splitlines()[-2]  # 合并PDF
    if id not in product_ids:
        product_ids[id] = [[page_num + 1, False]] #[page, printed?]
    else:
        product_ids[id].append([page_num + 1, False])
print(product_ids)

while True:
    target_id = input("Product Id: ")
    if target_id.lower() == 'q':
        print("Exiting...")
        break
    # if target_text.lower() == 'q':
    #     print("Exiting...")
    #     break
    # page_has_found = False
    # for page_num in range(pdf_doc.page_count):
    #     page = pdf_doc[page_num]  #  SHG162
    #     text = page.get_text('text')  # 合并PDF

    #     if target_text in text:
    #         print(f'text type: {type(text)}')
    #         print(f'text: {text}')
    #         print(f'to array: {text.splitlines()}')
    #         print(f'id: {text.splitlines()[-2]}')
    #         print(f"Printing page {page_num + 1}")
    #         page_has_found = True
    #         # subprocess.run(["lp", "-d", "Zebra_Technologies_ZTC_ZP_450_200dpi", "-o", f"page-ranges={page_num+1}", pdf_file_path])
    #         break
    if target_id not in product_ids:
        print("Label not found")
    elif False not in [x[1] for x in product_ids[target_id]]:
        response = input("All labels for this product have been printed, enter 'y' to print again, 'n' to exit: ")
        if response.lower() == 'y':
            for x in product_ids[target_id]:
                x[1] = False
        else:
            print("Exiting...")
            break
    else:
        count = 0
        for x in product_ids[target_id]:
            if not x[1]:
                print(f"Printing page {x[0]}")
                x[1] = True
                count += 1
                print(f'{target_id}: {count} of {len(product_ids[target_id])}')
                # subprocess.run(["lp", "-d", "Zebra_Technologies_ZTC_ZP_450_200dpi", "-o", f"page-ranges={x[0]}", pdf_file_path])
                break
            count += 1

pdf_doc.close()

