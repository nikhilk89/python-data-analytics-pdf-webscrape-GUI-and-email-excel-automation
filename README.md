# python-data-analytics-pdf-webscrape-GUI-and-email-excel-automation
Data Analytics | Streamlit Dashboards | PDF Report Generators &amp; Editors | Web Scraping Pipelines | Excel, Email &amp; File Automation

A collection of Data analysis tools with interactive Streamlit Dashboard, Web scrapers, PDF extraction and loader, Tkinter based GUI app with form fill, search using sqlite3, Email-Excel-File & Folder Automation scripts built using Python, Streamlit, Pandas and BeautifulSoup, 

--------------

## 🛠️ Included Modules

### 1. 📊 Risk Profile Analytics (Interactive Streamlit Dashboard)
* **Tech Stack:** `Streamlit`, `Pandas`, `Matplotlib`
* **Overview:** An end-to-end analytics dashboard for wealth management, risk profiling, and client portfolio diversification.
* **Key Features:**
  * **Data Pipeline:** Ingests and cleans raw client portfolio datasets, categorizing risk profiles and evaluating diversification metrics.
  * **Automated Reporting:** Programmatically generates advisor-level client summary reports in `.csv` and exports visualization charts as `.png`.
  * **Interactive UI:** Provides a web dashboard for dynamic metric filtering, trend analysis, and custom data exports.

---

### 2. 📄 PDF Processing Suite (Extraction & Fillable Generators)

#### 2.1 Multi-Vendor Invoice Data Extraction Pipeline
* **Tech Stack:** `pdfplumber`, `re` (Regex), `Pandas`, `OpenPyXL`
* **Overview:** Automated document routing and text extraction pipeline designed to parse multi-layout PDF invoices.
* **Key Features:**
  * **Vendor Router Engine:** Scans raw text using `pdfplumber` and conditionally routes documents to specific parsing logic based on vendor keyword signatures.
  * **Regex Line-Item Parsing:** Extracts invoice metadata (Invoice #, Date, Total) alongside detailed item tables using custom Regular Expressions tuned to vendor layouts.
  * **Dual-Tab Excel Output:** Exports invoice header summaries into an `"Invoice Summaries"` tab and detailed line items into a `"Line Items"` tab in a consolidated Excel workbook.

#### 2.2 PDF Mail-Merge & Direct Deposit Form Generator
* **Tech Stack:** `FillPDF` / `pypdf`, `Pandas`
* **Overview:** Takes client banking records from Excel and automatically populates official fillable PDF forms.
* **Key Features:**
  * **Data Normalization:** Ingests client records (`ddfields.xlsx`) via Pandas, formats dates (`MM/DD/YYYY`), and handles missing fields cleanly.
  * **Interactive Form-Filling:** Maps Excel columns directly to PDF form fields (`chase-dd.pdf`) and toggles dynamic radio buttons (e.g., *Checking* vs. *Savings*).
  * **Batch Generation:** Programmatically outputs individual, client-named PDFs into a designated output directory.


#### 2.3 PDF OCR Text Extraction Script

* **Tech Stack:** `PyMuPDF`, `pytesseract`(Tesseract OCR), `Pillow`(PIL), `JSON` 
* **Overview:** Converts scanned PDF documents into high-resolution images to extract structured text using Tesseract OCR and export it to JSON
* **Key Features:**
  * **High-Res Rendering:** Renders PDF pages into RGB images at 300 DPI via PyMuPDF to maximize OCR accuracy
  * **Configured OCR Engine:** Uses Tesseract with Page Segmentation Mode 6 optimized for standard text blocks
  * **Structured JSON Export:** Formats extracted text by page number and saves the output to local directory (static/ocr.json)

---

### 3. 🌐 Web Scraping & Extraction Pipelines (Static & Dynamic)

#### 3.1 E-Commerce Catalog Scraper (Static Parsing)
* **Tech Stack:** `Requests`, `BeautifulSoup4`, `CSV`
* **Overview:** A web crawling tool for extracting structured product listings from static e-commerce catalogs.
* **Key Features:**
  * **HTML Parsing:** Crawls `books.toscrape.com` to extract titles, stock status, and prices.
  * **Structured Data Export:** Cleans text markup and saves normalized outputs directly into an Excel-ready `books.csv` file.

#### 3.2 Financial Historical Stock Scraper (Dynamic JS Rendering) - Yahoo Finance
* **Tech Stack:** `Selenium`, `Pandas`, `OpenPyXL`
* **Overview:** Headless web scraping pipeline for dynamic financial data tables (e.g., NVIDIA).
* **Key Features:**
  * **Automated URL Parameters:** Programmatically calculates and converts calendar dates (`YYYY-MM-DD`) into Unix timestamps required by Yahoo Finance.
  * **Anti-Scraping Bypass:** Runs Selenium in headless mode with custom User-Agent headers to handle dynamic JavaScript loading.
  * **Formatted Excel Export:** Loads price histories into Pandas and auto-adjusts column widths in `NVDA_history.xlsx` to prevent text truncation.

---

### 4. 📧 Email, Excel & File System Workflow Automation
* **Tech Stack:** `imaplib` / `smtplib`, `Pandas`, `OpenPyXL`
* **Overview:** An end-to-end automation pipeline that ingests incoming email attachments, merges datasets, constructs pivot reports, and dispatches master summary emails.
* **Key Features:**
  * **Attachment Harvester:** Searches inbox messages by subject/sender and downloads Excel attachments using incremental naming rules.
  * **Data Consolidation:** Stacks raw workbook sheets into a single master dataset saved with timestamped filenames (`Consolidated_Report_YYYYMMDD_HHMM.xlsx`).
  * **Automated Pivot Reports & Dispatch:** Programmatically builds native Excel Pivot Tables and embedded charts, drafting and emailing the finalized report to stakeholders automatically.

---

### 5. 🖥️  GUI Application - Desktop Data Entry & Reviewer Portal
* **Tech Stack:** `Tkinter`, `SQLite3`, `Pandas`, `OpenPyXL`
* **Overview:** A desktop graphical interface built for structured data intake, dual persistence, and reviewer administration.
* **Key Features:**
  * **Form Intake:** User-friendly interface capturing user details (Name, Age, Course Selection, Rule Agreement) with input validation.
  * **Dual Persistence:** Automatically writes submitted data simultaneously to an Excel workbook and a local `SQLite3` database.
  * **Reviewer Dashboard:** Features a dedicated secondary tab allowing administrators to query, view, and audit entries directly from the database.

---

### 6. 🌤️ Weather Forecast CLI Utility
* **Tech Stack:** `Requests`, `JSON`
* **Overview:** A lightweight command-line tool connecting to external weather APIs for instant location forecasts.
* **Key Features:**
  * **REST API Querying:** Interacts with the OpenWeatherMap REST API via dynamic HTTP GET requests.
  * **Response Validation:** Handles API status payloads gracefully, returning clear user feedback for invalid city inputs or `404` errors.
