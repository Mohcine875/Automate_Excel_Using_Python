# Japan Sales Excel Automation — Python

## Project Overview

This project automates the extraction and consolidation of sales data from multiple Excel files using Python.

The workflow processes sales files for Japanese cities including Osaka, Tokyo, and Yokohama for the years 2022–2023. Instead of manually opening each workbook and copying values from specific cells or sheets, the script automatically extracts the required information and consolidates it into a single output file.

The objective is to reduce repetitive manual Excel work, improve accuracy, and make reporting workflows faster and more scalable.

---

## Business Problem

Sales data is often stored across multiple Excel files, cities, sheets, and reporting periods.

Manually extracting values from each file can be:

- time-consuming
- repetitive
- error-prone
- difficult to scale
- inconsistent across reporting cycles

This project solves that problem by using Python to automate the extraction and consolidation process.

---

## Business Objective

The automation was designed to:

- process multiple Excel files automatically
- extract predefined data points from specific sheets and cells
- consolidate extracted results into one output workbook
- generate timestamped output files for tracking
- reduce manual reporting time
- improve consistency and accuracy in Excel-based reporting

---

## Tools Used

| Tool / Library | Purpose |
|---|---|
| Python | Automation scripting |
| pandas | Data manipulation and tabular processing |
| openpyxl | Reading and writing Excel files |
| xlwings | Excel automation and workbook interaction |
| datetime | Generating timestamped output filenames |
| Excel | Source files, configuration, and output reports |
| GitHub | Project documentation and version control |

---

## Project Workflow

```text
Excel Sales Files
   ↓
Configuration Workbook
   ↓
Python Automation Script
   ↓
Cell and Sheet Extraction
   ↓
Data Consolidation
   ↓
Timestamped Output File
   ↓
Final Excel Report
```

---

## Key Features

The project includes the following automation features:

- batch processing of multiple Excel files
- extraction of specific values from defined sheets and cells
- use of a configuration workbook to control extraction logic
- automated consolidation of extracted data
- generation of a timestamped output Excel file
- reduced manual copy-paste work
- improved reporting consistency

---

## Project Structure

```text
Automate_Excel_Using_Python/
│
├── Data/
│   ├── Osaka_Sales_2022-2023.xlsx
│   ├── Tokyo_Sales_2022-2023.xlsx
│   └── Yokohama_Sales_2022-2023.xlsx
│
├── Output/
│
├── excel_scraper.py
├── excel_scraper.xlsm
└── README.md
```

---

## Data Files

| File | Description |
|---|---|
| `Osaka_Sales_2022-2023.xlsx` | Sales data workbook for Osaka |
| `Tokyo_Sales_2022-2023.xlsx` | Sales data workbook for Tokyo |
| `Yokohama_Sales_2022-2023.xlsx` | Sales data workbook for Yokohama |
| `excel_scraper.xlsm` | Excel configuration workbook used by the automation |
| `excel_scraper.py` | Main Python automation script |
| `Output/` | Folder where consolidated results are saved |

---

## How the Automation Works

The automation process follows these steps:

1. Reads the configuration settings from `excel_scraper.xlsm`.
2. Identifies the Excel files to process from the `Data/` folder.
3. Opens each sales workbook.
4. Extracts predefined values from specific sheets and cells.
5. Consolidates the extracted data into a structured output.
6. Saves the final result in the `Output/` folder with a timestamped filename.

---

## How to Run the Project

### 1. Install Required Libraries

```bash
pip install pandas openpyxl xlwings
```

### 2. Prepare the Project Files

Make sure the project contains:

```text
excel_scraper.py
excel_scraper.xlsm
Data/
Output/
```

The `Data/` folder should contain the Excel sales files.

### 3. Configure the Excel Settings File

Open:

```text
excel_scraper.xlsm
```

Then update the settings sheet to define:

- input folder
- target workbook names
- sheet names
- cell references
- extraction rules

### 4. Run the Script

```bash
python excel_scraper.py
```

### 5. Review the Output

After execution, the consolidated report is saved in the `Output/` folder with a timestamped filename.

Example:

```text
results_20230904_153000.xlsx
```

---

## Python Skills Demonstrated

This project demonstrates practical Python automation skills, including:

- reading Excel files programmatically
- working with multiple workbooks
- extracting values from specific sheets and cells
- using configuration-driven automation
- consolidating data into a structured output
- generating timestamped result files
- automating repetitive business reporting workflows

---

## Example Use Case

This project can be used in reporting scenarios where sales data is stored across multiple city-specific Excel files.

For example, a business analyst can use this automation to extract monthly or yearly sales metrics from Osaka, Tokyo, and Yokohama files and automatically consolidate them into one final report.

---

## Project Value

This project demonstrates the ability to:

- automate Excel-based reporting workflows
- reduce manual data extraction effort
- improve consistency in recurring reports
- minimize copy-paste errors
- process multiple files efficiently
- apply Python to real business productivity tasks

---

## Technologies

- Python
- pandas
- openpyxl
- xlwings
- Excel
- Automation
- Data Processing
- Reporting Workflow

---

## Project Status

Completed as a Python-based Excel automation project.

---

## Author

**Mohcine Behate**

Python Automation and Data Productivity Portfolio Project
