**Web Scrapper for Amazon**

A Python Utility for scraping Product Listings and Information from Amazon Website into a well organized Excel Workbook

Requirements: 
1. Python (with pip)
2. Requirement Modules (from respective requirements.txt files)
3. Chrome Driver [from here](https://googlechromelabs.github.io/chrome-for-testing/)


**Part I: Scraping by Search**

You can enter a keyword to search, and it will store all the products available on the first page in an excel sheet.
The attributes it will cover are: 

1. Item Name
2. Features and Description
3. Brand
4. Price
5. Item Picture
6. Quantity Sold Past Month
7. Number of Reviews
8. Product Link

Note: Create an excel sheet named "amazon_products", else you can change in code to rename this file.

![image](https://github.com/rahulpandharkar/amazon_scraping/assets/103379268/9e08eb0f-43a2-416b-bc2f-63ca35738816)


**Part II: Scraping by Link**

You can enter any Product Link, and it will store all the relevant information available on the products page in the excel sheet. 
The attributes it wwill cover are: 

1. Item Name
2. Features and Description
3. Brand
4. Price
5. Item Picture
6. Item Dimensions
7. Item Weight
8. Bought Past Month
9. Rating
10. Number of Reviews
11. Imported from
12. Manufacturing Details
13. URL

Note: Create an excel sheet named "product_info", else you can change in code to rename this file.

![image](https://github.com/rahulpandharkar/amazon_scraping/assets/103379268/d6dc63c0-97c2-4641-9c06-d0f79cea5b59)
