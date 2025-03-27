# P1 – Excel Grade Summary System

## Overview

This project is a solution to a problem presented by a high school teacher who needs to improve the format of an Excel spreadsheet that tracks student grades for multiple classes. The teacher's existing Excel file contains all student data for different classes in one spreadsheet, with student information in a single column. The task is to automate the process of organizing this data, creating a new Excel file that summarizes and formats the data for each class.

## Features

- **Class-specific worksheets**: Automatically generates a worksheet for each class (e.g., Algebra, Calculus, etc.).
- **Student Data Formatting**: Organizes and formats student data, such as last name, first name, student ID, and grade.
- **Class Summaries**: Generates summary statistics for each class, including:
  - Highest grade
  - Lowest grade
  - Mean grade
  - Median grade
  - Number of students in the class
- **Formatting & Styling**:
  - Adds filters to each worksheet.
  - Bolds the headers of specific columns.
  - Adjusts column widths based on the length of the column headers.
- **Save as New Excel File**: The formatted data is saved in a new Excel file named `formatted_grades.xlsx`.

## Libraries Required

- `openpyxl`: For reading and writing Excel files.
- `openpyxl.styles.Font`: For text formatting in the Excel sheet.

## External Files Required

- The project requires an Excel file with student data, which should be formatted as follows:
  - **Class Name**: The class for which the student is registered.
  - **Student Info**: A single column with student data in the format `LastName_FirstName_StudentID`.
  - **Grade**: A numeric grade between 0 and 100.

You can download example files from the Learning Suite project description to use as input.

## Logical Flow

1. **Reading the Input File**:
   - The program imports one of the provided example Excel files using the `openpyxl` library.
   - The program should be robust enough to handle different files in the same format.

2. **Creating New Worksheets for Each Class**:
   - For each class, a new worksheet is created.
   - The program checks if a sheet already exists for a class and avoids duplicating it.

3. **Organizing Student Data**:
   - Data is parsed from the input file and split into last name, first name, student ID, and grade.
   - This data is appended to the appropriate worksheet.

4. **Adding Filters and Summary Information**:
   - A filter is added to the columns for easier data manipulation.
   - Summary statistics (highest grade, lowest grade, mean grade, median grade, and the number of students) are added in columns F and G.

5. **Formatting**:
   - The headers for columns A, B, C, D, F, and G are bolded.
   - Column widths are adjusted based on the length of the headers.

6. **Saving the Output**:
   - The modified workbook is saved as `formatted_grades.xlsx`.

## Example Output

The program generates a formatted Excel file (`formatted_grades.xlsx`) with separate sheets for each class. The sheets will contain student data and summary statistics for each class.

You can check the example output provided in the Learning Suite assignment description.

## How to Run

1. Clone this repository.
2. Ensure you have the `openpyxl` library installed. You can install it using pip:

    ```bash
    pip install openpyxl
    ```

3. Download the input Excel files (available in the Learning Suite project description).
4. Run the Python script on the command line:

    ```bash
    python grade_summary.py
    ```

5. The output will be saved as `formatted_grades.xlsx`.

