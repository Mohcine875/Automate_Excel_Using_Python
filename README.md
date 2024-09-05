# Japan Sales Automation Project (2022-2023)

## Project Overview

This project automates the extraction of specific data points from multiple Excel files containing sales data for the Japanese cities of Osaka, Tokyo, and Yokohama for the years 2022-2023. The goal is to simplify the consolidation of sales data across different Excel files by automatically extracting predefined data from cells and sheets, and then compiling them into a single output file.

## Key Features

- **Batch Processing of Excel Files**: Automatically processes all Excel sales files for different cities and extracts relevant information from specified cells and sheets.
- **Data Extraction**: Data is extracted from specific cells in the spreadsheets as defined in a configuration sheet.
- **Automated Data Consolidation**: The extracted data is compiled into a single Excel file with a timestamped filename for easy tracking.
- **Time Efficiency**: This project reduces the time spent manually processing each file and minimizes human errors.

## Technologies Used

- **Python**: For scripting and automation logic.
- **Pandas**: For data manipulation and handling.
- **OpenPyXL**: For reading and writing Excel files.
- **XLWings**: For interacting with Excel and automating tasks.
- **Datetime**: For generating timestamped filenames.

## Project Structure

```bash
.
├── excel_scraper.py                 # Main automation script
├── excel_scraper.xlsm               # Excel file with macros and configuration settings
├── Osaka_Sales_2022-2023.xlsx       # Sales data for Osaka
├── Tokyo_Sales_2022-2023.xlsx       # Sales data for Tokyo
├── Yokohama_Sales_2022-2023.xlsx    # Sales data for Yokohama
├── output                           # Directory where results are saved
└── README.md                        # Project documentation
```

## Data Files

### Osaka_Sales_2022-2023.xlsx
- Contains sales data for the city of Osaka during 2022-2023.

### Tokyo_Sales_2022-2023.xlsx
- Contains sales data for the city of Tokyo during 2022-2023.

### Yokohama_Sales_2022-2023.xlsx
- Contains sales data for the city of Yokohama during 2022-2023.

## How to Run the Project

1. **Prerequisites**:
   - Ensure that Python is installed on your machine.
   - Install the necessary Python packages by running the following command in your terminal:
     ```bash
     pip install pandas openpyxl xlwings
     ```

2. **Place Your Files**:
   - Place the `excel_scraper.py` script, the `excel_scraper.xlsm` Excel file, and the sales data files (`Osaka_Sales_2022-2023.xlsx`, `Tokyo_Sales_2022-2023.xlsx`, `Yokohama_Sales_2022-2023.xlsx`) in the same directory.

3. **Configure the Excel File `excel_scraper.xlsm`**:
   - Open the `excel_scraper.xlsm` file in Excel.
   - Navigate to the "Settings" sheet and configure the input folder (`INPUT_FOLDER`) to the folder where your sales files are located.
   - Ensure the sheet names and cell references are correctly defined in the "tbl_SETTINGS" table.

4. **Run the Script**:
   - Open a terminal in the directory where the files are located.
   - Run the following command to start the automation process:
     ```bash
     python excel_scraper.py
     ```

5. **Check the Output**:
   - Once the script completes execution, an Excel file containing the consolidated data will be generated in the `output` folder. The file will have a timestamped name, such as `results_20230904_153000.xlsx`.

## Example Use Case

The automation script was used to process sales data from Osaka, Tokyo, and Yokohama for the years 2022-2023. It automatically extracted key sales metrics from specific cells in each file and compiled them into a final report, saving hours of manual data extraction and consolidation.

## Conclusion

This automation project enables easy management of large datasets scattered across multiple Excel files, reducing manual tasks and improving efficiency. It is particularly useful in cases where data needs to be regularly extracted and consolidated for reporting or analysis.
