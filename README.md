# Report Card Generator
This project is a Python-based script that generates personalized PDF report cards for students based on their academic performance data stored in an Excel file. It processes the data, calculates total and average scores, and creates a PDF report for each student with detailed subject-wise scores.

## Features
Read Student Data from Excel: Reads an Excel file containing student information, subjects, and scores.
Data Processing: Calculates total and average scores for each student and organizes subject-wise scores.
PDF Report Generation: Creates a PDF report card for each student with a summary of their performance and subject-wise scores.
Custom Styling: Enhances readability using custom table styles and color themes.

## Prerequisites
Before running the script, ensure you have the following:

- Python 3.x: Download Python
- Required Libraries:
- pandas: For reading and processing Excel data.
- reportlab: For generating PDF files.
- Install the dependencies using pip:

bash
Copy
Edit
pip install pandas reportlab
Excel File: An Excel file named student_scores.xlsx (or any specified name) containing the following columns:
Student ID
Name
Subject
Score
Usage
## Prepare the Excel File:

Ensure the Excel file has the required columns: Student ID, Name, Subject, Score.
## Run the Script:

Save the script in the same directory as the Excel file.
## Execute the script:
bash
Copy
Edit
python report_card_generator.py
##Generated PDFs:

For each student, a PDF file named report_card_<Student ID>.pdf will be created in the same directory.
## Example Input and Output
Input (Excel File: student_scores.xlsx)
Student ID |	Name	 |Subject    	|Score
--------------------------------------
101	       | John Doe|	Mathematics|	85
101	       |John Doe |Science	     |  90
102	       |Jane Roe |	Mathematics|	95
102        |Jane Roe |	Science    |	88
# Output (PDF Report Card)
For John Doe:

Total Score: 175
Average Score: 87.5
Subject-wise Scores:
Mathematics: 85
Science: 90
For Jane Roe:

Total Score: 183
Average Score: 91.5
Subject-wise Scores:
Mathematics: 95
Science: 88

## File Structure

bash
Copy
Edit
``` project_directory/
│
├── student_scores.xlsx  # Input Excel file
├── report_card_generator.py  # Python script
├── report_card_<StudentID>.pdf  # Generated PDFs

## Error Handling
Missing Columns: If the Excel file is missing required columns, the script will notify you and terminate.
File Not Found: If the Excel file path is incorrect, an error message will be displayed.
Invalid Data: Rows with missing or invalid values are skipped during processing.
