# Automated ETL Pipeline Project – Prince Igweze

## Task

Company XYZ Services Limited keeps a catalogue of their financial statements in Excel sheets every month. These financial statements contain annual dollar amounts paid by month for various sub-categories across 13 provinces in Canada.

## Methodology

Using Python, I designed a script to automatically **Extract**, **Transform**, and **Load** (ETL) the dollar amounts for each sub-category across all provinces. The data is aggregated into a single Excel sheet and loaded into an SQL database for further analysis.

### Key Steps:

1. **Extract**: Retrieve data from monthly Excel sheets.
2. **Transform**: Aggregate and clean the data to ensure consistency and accuracy.
3. **Load**: Insert the transformed data into a consolidated Excel file and an SQL database.

![Alt text](https://github.com/PrinceIgweze/Finance-ETL-Pipeline-/blob/main/Flowchart.png?raw=true)
*Figure 1: ETL Process Diagram illustrating the data flow and processing stages.*

## How to Run the Package

1. **Download the ETL Folders**:
   - Clone or download the repository and place the ETL folders in your desired location.

2. **Configure Parameters**:
   - Navigate to the `parameters` folder and open the `parameter.xlsx` file.
   - Modify the values to fit your local computer and database configuration.

3. **Run the Script**:
   - Open the `main.py` file in your preferred IDE.
   - Execute the script to run the ETL process.


