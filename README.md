# Shipping Label Finder
## Description
This Python script reduces e-commerce retailers' hassle of shuffling around hundreds and thousands of shipping labels to match the order number with the correct label. Using an external Label Scanner to scan the product ID attached to a product, the program can precisely and promptly locate the matched label and print it out. This program has saved up to 80% of the package processing time.

## Important
1. It's common that a shipping label could involve multiple products (See [example](/images/multiple_products.pdf)). In this case, the first product ID(`id:DSC038`) will be used as the ***identifier product ID***

## How To Use
1. Select and run ```GUI.py```.
   
2. Select a UPS shipping label PDF in the menu. (Users must ensure their shipping labels include a product ID at the bottom. See [example](/images/shipping_label_example.pdf))
   
3. Select a printer in the menu. (Users must ensure the printer is properly connected)
   
4. Everything is all set. Start entering the product ID in the prompt.

## Updates
### 2024-08-11
- Updated the Sort&Save feature to first sort the selected PDF by the identifier product ID in each page, then sort by the number of the identifier products present in each page.

### 2024-05-02
- Added the Sort&Save feature, which sorts the selected PDF by the identifier product ID in each page and creates a sorted PDF.

### 2024-03-23
- Created the GUI for the program.
- Got rid of the order summary spreadsheet requirement, and only the UPS shipping labels PDF is required.
  
### 2024-02-24
- Used pandas to process the order summary spreadsheet instead of solely reading the shipping label PDF.
- Able to correctly match shipping labels that contain multiple products or multiple quantities of the same product.
- The order summary spreadsheet must follow a pre-established format for the program to run properly.
