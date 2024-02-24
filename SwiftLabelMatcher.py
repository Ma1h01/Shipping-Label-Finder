import pandas as pd
import fitz
import subprocess

def set_up_table(excel_file_path, pdf_file_path):
    try:
        shipping_labels_pdf = fitz.open(pdf_file_path)
        order_details_excel = pd.read_excel(excel_file_path)
    except Exception as e:
        print(e)
        quit()

    # copy the data frame, containing Order Number and ID columns only        
    table = order_details_excel[['Order Number', 'ID']]
    # use the Order Number as the index
    table.set_index('Order Number', inplace=True)    

    for page_num in range(shipping_labels_pdf.page_count):
        page = shipping_labels_pdf[page_num]
        # extract all the text from the page into a list of strings
        order_number = page.get_text().splitlines()
        # since the order number can be in the second to last or last of the list, we need to check
        # order number contains NO alphabets
        order_number = order_number[-1] if (any(char.isalpha() for char in order_number[-2])) else order_number[-2]
        # add the page number and printed status to the table
        table.loc[order_number, ['Page Number', 'Printed']] = page_num + 1, False

    shipping_labels_pdf.close()
    return table

## Start of the program
excel_file_path = f'./orders/{input("Enter Excel document name: ")}.xlsx'
pdf_file_path = f'./orders/{input("Enter PDF document name: ")}.pdf'

table = set_up_table(excel_file_path, pdf_file_path)

while True:
    target_id = input("Product Id: ")
    if target_id.lower() == 'q':
        print("Exiting...")
        break

    # check if there is an order number contains the specified product ID
    filtered_table = table[table['ID'] == target_id]
    if len(filtered_table) == 0:
        print(f"No orders found for product {target_id}")
        continue
    else:
        # check if all orders corresponding to the product ID have been printed
        filtered_table_unprinted = filtered_table[filtered_table['Printed'] == False]
        if len(filtered_table_unprinted) == 0:
            print(f"All labels for product {target_id} have been printed")
            response = input("Enter 'y' to print again, 'n' to continue: ")
            if response.lower() == 'y':
                # reset all the printed status to False
                for index,row in filtered_table.iterrows():
                    table.loc[index, 'Printed'] = False
        else:
            for index,row in filtered_table_unprinted.iterrows():
                print(f"{target_id}: {len(filtered_table) - len(filtered_table_unprinted) + 1} of {len(filtered_table)}")
                print(f"Printing page {int(row['Page Number'])} for order {index} and product {target_id}")
                subprocess.run(["lp", "-d", "Zebra_Technologies_ZTC_ZP_450_200dpi", "-o", f"page-ranges={int(row['Page Number'])}", "-o", "print-quality=5", pdf_file_path])
                table.loc[index, 'Printed'] = True
                break


