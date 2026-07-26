# python-data-analytics-pdf-webscrape-GUI-and-email-excel-automation
Data Analytics | Streamlit Dashboards | PDF Report Generators &amp; Editors | Web Scraping Pipelines | Excel, Email &amp; File Automation

A collection of Data analysis tools wiht interactive Streamlit Dashboard, Web scrapers, PDF extraction and loader, Tkinter based GUI app with form fill, search using sqlite3, Email-Excel-File & Folder Automation scripts built using Python, Streamlit, Pandas, BeautifulSoup, and ReportLab

--------------

## 🛠️ Included Modules

### 1. Data Analytics - Risk Profile Analytics (Interactive Streamlit Data Dashboard)
* reads and cleans raw client and portfolio data
* Categorise risk profiles, Risk diversifications
* pivot tables
* Report Generation - Advisor level-Client Risk profile Summary .CSV and .PNG (pie charts) files
* Streamlit Web Interactive Dashboard - where users can filter metrics, analyze trends, and export summary views


### 2. PDF Extractor & Generator 
#### 2.1 PDF Extractor- 
This is a PDF Invoice Data Extraction & Parsing Pipeline: 
* Vendor Router: Opens PDF invoices using pdfplumber, scans the raw text, and routes the document to a specific parser based on keywords (bioplex, zencorporations, or conincorporated).
* Extracts invoice metadata (Invoice #, Date, Total Amount) and table line items (Quantity, Description, Unit Price, Total) using custom Regular Expressions tuned to each vendor's layout.
* Compiles all processed summaries into an "Invoice Summaries" tab and all extracted item rows into a "Line Items" tab inside a single Excel workbook


#### 2.2 PDF Generator/Filler-
* Uses FillPDF to extract PDF fields and creates excel
* Takes raw customer records from Excel data and automatically generates filled-out Direct Deposit PDF form
* Reads customer records (ddfields.xlsx) via Pandas, formatting dates (MM/DD/YYYY) and handling missing data cleanly
* Programmatically opens fillable PDF templates (chase-dd.pdf), maps data frames to fillable form fields, and toggles dynamic radio buttons (e.g., Checking vs. Savings)
* outputs individual PDFs into a designated output directory


### 3. WebScraper 
#### 3.1 E-Commerce Catalog Scrape (Static)-
* crawls book catalog pages to extract structured product data into an Excel-ready CSV format
* HTML Parsing: Navigates books.toscrape.com using Requests & BeautifulSoup to extract book titles, stock availability, and prices
* Saves normalized output directly into books.csv for seamless spreadsheet processing

#### 3.2 Yahoo Finance Historical Stock Scraper(Dynamic) | Selenium-
* A headless web extraction pipeline that automates retrieving and exporting financial stock tables (eg- NVIDIA)
* Automatically converts standard human dates (YYYY-MM-DD) into Unix timestamps for Yahoo Finance URL parameters
* Use Selenium with custom User-Agent headers to handle dynamic JavaScript rendering and bypass basic anti-scraping blocks
* Uses Pandas to process daily stock rows and dynamic column-width auto-fitting in NVDA_history.xlsx so no values are truncated


### 4. Email, Excel, file and folder Automation-
* Automation to search email with subject, sender and download excel attachments with incremental names
* Combine all the excel attachments into 1 excel file and save into another folder with date time filename format
* create pivot and charts in excel 
* draft and send the master file on email to recipients wit message


### 5. GUI - form filling and search (using TKinter, SQLite3,Excel)- 
* GUI desktop app that lets users fill form, -name, age, courses, agree to rules and submit
* save the form data into excel and SQLite3 database
* another tab in the GUI - lets reviewer see the entries of users (from sqlite db)


### 6. Weather App (using requests)
* A lightweight CLI utility that interacts with the OpenWeatherMap REST API to retrieve and display real-time weather conditions and temperatures for any user-selected city
* Uses the requests library to construct dynamic HTTP GET requests targeting the OpenWeatherMap API endpoints
* Validates response payloads against API error codes (e.g., catching 404 errors for invalid or misspelled city inputs gracefully)



