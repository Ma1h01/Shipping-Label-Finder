# Shipping Label Finder
## Description
This Python script reduces e-commerce retailers' hassle of shuffling around hundreds and thousands of shipping labels to match the order number with the correct label. Using an external Label Scanner to scan the product ID attached to a product, the program can precisely and promptly locate the matched label and print it out. This program has saved up to 80% of the package processing time.

## Updates 
### 2024-2-24
- Use pandas to process the order summary excel, instead of solely reading the shipping label pdf
- Able to correctly match shipping label that contains multiple products or multiple quantity of the same products
- The order summary excel must follow a pre-established format for the program to run properly
