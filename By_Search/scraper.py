import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time
from PIL import Image
import requests
from io import BytesIO
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import os
import re

# Function to create the tkinter GUI
def create_gui():
    root = tk.Tk()
    root.title("Amazon Product Scraper")
    
    # Label and Entry for search query
    tk.Label(root, text="Enter Search Query:").pack()
    query_entry = tk.Entry(root, width=50)
    query_entry.pack(pady=10)
    
    # Button to start scraping
    def start_scraping():
        query = query_entry.get().strip()
        if query:
            root.destroy()
            scrape_amazon(query)
        else:
            messagebox.showwarning("Warning", "Please enter a search query.")

    tk.Button(root, text="Start Scraping", command=start_scraping).pack(pady=10)

    root.mainloop()

# Function to search for products on Amazon and extract details
def scrape_amazon(keyword):
    # Configure Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode (optional)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    # Path to your ChromeDriver
    chrome_driver_path = r""  # Put the path for Chrome Driver

    # Initialize WebDriver
    service = ChromeService(executable_path=chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    products = []
    seen_names = set()  # To keep track of seen product names

    try:
        # Open Amazon homepage
        driver.get("https://www.amazon.in/")
        
        # Locate the search input field and enter the keyword
        search_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'twotabsearchtextbox')))
        search_box.send_keys(keyword)
        search_box.submit()

        while True:
            # Wait until the search results are loaded
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@data-component-type='s-search-result']")))

            # Get the page source and parse with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, "html.parser")

            # Find all product containers
            product_cards = soup.find_all('div', {'data-component-type': 's-search-result'})

            for card in product_cards:
                try:
                    # Extract product name
                    product_name_elem = card.find('span', {'class': 'a-text-normal'})
                    if product_name_elem:
                        product_name = product_name_elem.text.strip()
                    else:
                        continue

                    # Skip if product name already seen to avoid duplicates
                    if product_name in seen_names:
                        continue
                    seen_names.add(product_name)

                    # Extract product image
                    product_image_elem = card.find('img', {'class': 's-image'})
                    if product_image_elem:
                        product_image = product_image_elem['src']
                    else:
                        product_image = "Image not available"

                    # Extract product price
                    product_price_elem = card.find('span', {'class': 'a-price-whole'})
                    if product_price_elem:
                        product_price = product_price_elem.text.strip()
                    else:
                        product_price = "Price not available"

                    # Extract product link
                    product_link_elem = card.find('a', {'class': 'a-link-normal'})
                    if product_link_elem:
                        product_link = "https://www.amazon.in" + product_link_elem['href']
                    else:
                        product_link = "Link not available"

                    try:
                        # Open product link to get more details
                        driver.get(product_link)
                        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "productTitle")))
                        product_soup = BeautifulSoup(driver.page_source, "html.parser")

                        # Extract Amazon Title
                        amazon_title_elem = product_soup.find('span', {'id': 'productTitle'})
                        amazon_title = amazon_title_elem.text.strip() if amazon_title_elem else "Title not available"

                        # Extract Features and Description
                        features_elem = product_soup.find('div', {'id': 'feature-bullets'})
                        features = features_elem.text.strip() if features_elem else "Features not available"

                        # Extract Brand
                        brand_elem = product_soup.find('a', {'id': 'bylineInfo'})
                        brand = brand_elem.text.strip() if brand_elem else "Brand not available"

                        # Extract Quantity Sold Past Month
                        quantity_sold_elem = product_soup.find('span', text=re.compile(r'(\d{3,}|\d{1,2}[0-9]{2})\+ bought in past month', re.IGNORECASE))
                        quantity_sold = quantity_sold_elem.text.strip() if quantity_sold_elem else "Quantity sold past month not available"

                        # Extract Number of Reviews
                        reviews_elem = product_soup.find('span', {'id': 'acrCustomerReviewText'})
                        reviews = reviews_elem.text.strip() if reviews_elem else "Reviews not available"

                    except Exception as e:
                        print(f"Error extracting data for {product_name}: {e}")
                        continue

                    # Append product details to list
                    products.append({
                        'Sr. No': len(products) + 1,
                        'Item Name': product_name,
                        'Features and Description': features,
                        'Brand': brand,
                        'Price': product_price,
                        'Item Photo': product_image,
                        'Quantity Sold Past month': quantity_sold,
                        'Number of Reviews': reviews,
                        'Link': product_link
                    })

                    # Return to search results page
                    driver.back()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@data-component-type='s-search-result']")))

                except Exception as e:
                    print(f"Error processing product: {e}")
                    continue

            # Find the next page button
            next_page_elem = driver.find_elements(By.CLASS_NAME, 's-pagination-next')
            if next_page_elem and 's-pagination-disabled' not in next_page_elem[0].get_attribute('class'):
                next_page_elem[0].click()
                time.sleep(3)  # Wait for the next page to load
            else:
                break

    except Exception as ex:
        print(f"Exception occurred: {str(ex)}")

    finally:
        # Close the driver
        driver.quit()

        # Sort products by number of reviews (placeholder logic)
        products.sort(key=lambda x: int(x['Number of Reviews'].split()[0].replace(',', '')), reverse=True)

        # Reassign serial numbers based on the sorted order
        for idx, product in enumerate(products, start=1):
            product['Sr. No'] = idx

        # Save products to Excel file named "amazon_products"
        save_to_excel(products, 'amazon_products.xlsx') 

# Function to save product details to Excel
def save_to_excel(products, filename):
    df = pd.DataFrame(products)
    
    # Save initial data to Excel
    df.to_excel(filename, index=False)

    # Create directory for product pictures if not exists
    if not os.path.exists('Product Pictures'):
        os.makedirs('Product Pictures')

    # Save images to "Product Pictures" folder
    for idx, product in enumerate(products, start=1):
        img_url = product['Item Photo']
        if img_url != "Image not available":
            response = requests.get(img_url)
            img = Image.open(BytesIO(response.content))
            img = img.resize((100, 100))  # Resize image to fit in cell
            img_path = f'Product Pictures/product_image_{idx}.png'
            img.save(img_path)
            product['Item Photo'] = img_path  # Update product dictionary with image path

    # Load the workbook and sheet
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # Add images to the Excel sheet
    for idx, product in enumerate(products, start=2):  # Start from row 2 (after headers)
        img_path = product['Item Photo']
        if img_path != "Image not available":
            img_openpyxl = OpenpyxlImage(img_path)
            ws.add_image(img_openpyxl, f'H{idx}')  # Assuming 'H' is the column for images

    # Save the workbook
    wb.save(filename)
    print(f"Data saved to {filename}")

# Entry point
if __name__ == "__main__":
    create_gui()
