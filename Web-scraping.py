import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox
import re
import os

# Fixed exchange rate from USD to INR
USD_TO_INR_EXCHANGE_RATE = 73.5

def get_amazon_data(search_query, price_from, price_to):
    # Convert price range from INR to USD
    price_from_usd = round(price_from / USD_TO_INR_EXCHANGE_RATE, 2)
    price_to_usd = round(price_to / USD_TO_INR_EXCHANGE_RATE, 2)

    base_url = f'https://www.amazon.com/s?k={search_query.replace(" ", "+")}&low-price={price_from_usd}&high-price={price_to_usd}'
    
    # Adding headers to mimic a regular browser request
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'TE': 'Trailers'
    }

    response = requests.get(base_url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        products = []

        for product in soup.find_all('div', {'class': 's-result-item'}):
            title_tag = product.find('h2')
            price_tag = product.find('span', {'class': 'a-offscreen'})
            link_tag = product.find('a', {'class': 'a-link-normal'})

            if title_tag and price_tag and link_tag:
                title = title_tag.text.strip()
                raw_price = price_tag.text.strip()
                link = "https://www.amazon.com" + link_tag['href']
                
                # Convert price from USD to INR
                try:
                    # Extract numerical values from the raw price string using regex
                    match = re.search(r'(\d+(\.\d+)?)', raw_price)
                    if match:
                        price_value = float(match.group(1))
                        price_inr = round(price_value * USD_TO_INR_EXCHANGE_RATE, 2)
                        products.append({'Title': title, 'Price (INR)': price_inr, 'Link': link})
                    else:
                        print(f"Failed to convert price for {title}. No numerical value found in price string.")
                except ValueError:
                    print(f"Failed to convert price for {title}. Invalid price format.")
                
        return products

    else:
        print(f"Failed to fetch data. Status code: {response.status_code}")
        print(response.text)  # Print HTML content for debugging
        return None

def save_to_excel(data, file_name='amazon_products.xlsx'):
    df = pd.DataFrame(data)
    
    # Write to Excel with formatting
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Add color to the header row
        header_fill = openpyxl.styles.PatternFill(start_color="66B2FF", end_color="66B2FF", fill_type="solid")
        for cell in worksheet['A1:C1']:
            for col in cell:
                col.fill = header_fill
        
        # Add hyperlinks to the 'Link' column
        for row_num, link in enumerate(df['Link'], start=2):
            cell = worksheet.cell(row=row_num, column=3)
            cell.hyperlink = link
            cell.value = 'Link'
            cell.style = 'Hyperlink'
    
    print(f"Data saved to {file_name}")
    return file_name

def open_excel(file_path):
    os.system(f'start excel "{file_path}"')

def on_search_button_click():
    search_query = entry_search.get()
    price_from = entry_price_from.get()
    price_to = entry_price_to.get()
    
    if not price_from or not price_to:
        result_label.config(text="Please enter both lower and upper bounds of the price range.", foreground="red")
        return

    try:
        price_from = float(price_from)
        price_to = float(price_to)
    except ValueError:
        result_label.config(text="Invalid price range. Please enter valid numbers.", foreground="red")
        return

    # Show loading message
    result_label.config(text="Searching...", foreground="blue")
    root.update()

    amazon_data = get_amazon_data(search_query, price_from, price_to)

    if amazon_data:
        file_path = save_to_excel(amazon_data)
        result_label.config(text=f"Data saved to '{file_path}'", foreground="green")
        
        # Display results in Treeview
        display_results_in_treeview(amazon_data)
        
        open_excel_option = messagebox.askyesno("Open Excel", "Do you want to open the Excel file?")
        if open_excel_option:
            open_excel(file_path)
    else:
        # Show error message
        result_label.config(text="Failed to fetch data. Check your internet connection.", foreground="red")

def on_clear_button_click():
    # Clear input fields, Treeview, and result label
    entry_search.delete(0, tk.END)
    entry_price_from.delete(0, tk.END)
    entry_price_to.delete(0, tk.END)
    result_label.config(text="", foreground="green")

    for row in tree.get_children():
        tree.delete(row)

def display_results_in_treeview(data):
    # Clear previous results
    for row in tree.get_children():
        tree.delete(row)

    # Display results in Treeview
    for item in data:
        tree.insert("", "end", values=(item['Title'], item['Price (INR)'], item['Link']))

    result_label.config(text=f"Search Results:", foreground="green")

# Create main window
root = tk.Tk()
root.title("Amazon Scraper")

# Set background color similar to Amazon
root.configure(bg='#232f3e')

# Create and place widgets with styles
label_search = ttk.Label(root, text="Enter the product you want to search on Amazon:", background='#232f3e', foreground='white', font=('Arial', 14, 'bold'))
label_search.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)

entry_search = ttk.Entry(root, width=40, font=('Arial', 12))
entry_search.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)

label_price_range = ttk.Label(root, text="Enter the price range in INR (optional):", background='#232f3e', foreground='white', font=('Arial', 14, 'bold'))
label_price_range.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)

label_price_from = ttk.Label(root, text="From:", background='#232f3e', foreground='white', font=('Arial', 14, 'bold'))
label_price_from.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)

entry_price_from = ttk.Entry(root, width=20, font=('Arial', 12))
entry_price_from.grid(row=1, column=2, padx=10, pady=10, sticky=tk.W)

label_price_to = ttk.Label(root, text="To:", background='#232f3e', foreground='white', font=('Arial', 14, 'bold'))
label_price_to.grid(row=1, column=3, padx=10, pady=10, sticky=tk.W)

entry_price_to = ttk.Entry(root, width=20, font=('Arial', 12))
entry_price_to.grid(row=1, column=4, padx=10, pady=10, sticky=tk.W)

button_search = ttk.Button(root, text="Search", command=on_search_button_click, style='TButton')
button_search.grid(row=2, column=0, columnspan=5, pady=10)

button_clear = ttk.Button(root, text="Clear", command=on_clear_button_click, style='TButton')
button_clear.grid(row=2, column=3, columnspan=2, pady=10)

result_label = ttk.Label(root, text="", foreground="green", background='#232f3e', font=('Arial', 14, 'bold'))
result_label.grid(row=3, column=0, columnspan=5, pady=10)

# Create Treeview to display results
tree = ttk.Treeview(root, columns=("Title", "Price (INR)", "Link"), show="headings")
tree.heading("Title", text="Title")
tree.heading("Price (INR)", text="Price (INR)")
tree.heading("Link", text="Link")
tree.grid(row=4, column=0, columnspan=5, pady=10)

# Add a custom style for the search and clear buttons
style = ttk.Style()
style.configure('TButton', font=('Arial', 12, 'bold'))
style.map('TButton',
          background=[('active', '#FF9900'), ('pressed', '#232f3e')],
          foreground=[('pressed', '#FF9900')])

# Start the main loop
root.mainloop()
