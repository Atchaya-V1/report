'''
This Python script reads data from an Excel file that contains student scores across different subjects. Here's a breakdown of how the script works:

1.Reading Data: The script starts by reading the Excel file using pandas. It checks if the necessary columns like "Student ID", "Name", "Subject", and "Score" are present. If anything’s missing, it raises an error.

2.Processing Data: After loading the data, it groups the scores by student ID and name, and calculates two things:
-Total Score: The sum of scores across all subjects for each student.
-Average Score: The mean score for each student.
Additionally, it collects the subject-wise scores for each student.

3.Generating PDF: For each student, the script generates a PDF report card:
It first prints the student’s name at the top.
Then it displays the total and average scores in a simple table.
Below that, a table lists the subject-wise scores.
The PDF is saved with the name report_card_<StudentID>.pdf.

4.Styling: The styling of the tables is simple yet effective. The total and average score table doesn't have any background color, while the subject score table has a light blue background for clarity.

Finally, the script automatically generates a PDF for each student in the Excel file, making it easy to generate multiple reports in one go.
'''



import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

# Step 1: Read the Excel File
def read_excel(file_path):
    try:
        # Load the Excel data into a pandas DataFrame
        data = pd.read_excel(student_records.xlsx)
        
        # Check if the required columns are present
        required_columns = {"Student ID", "Name", "Subject", "Score"}
        if not required_columns.issubset(data.columns):
            raise ValueError(f"The Excel file is missing one or more of the following columns: {', '.join(required_columns)}")
        
        return data
    except Exception as e:
        print(f"Error occurred while reading the Excel file: {e}")
        return None

# Step 2: Process the Data
def process_data(data):
    # Group the data by student ID and name
    grouped_data = data.groupby(["Student ID", "Name"])
    
    # Calculate the total and average scores for each student
    summary = grouped_data["Score"].agg(Total="sum", Average="mean").reset_index()
    
    # Get subject-wise scores for each student
    subject_details = data.groupby(["Student ID", "Name", "Subject"])["Score"].sum().reset_index()
    
    return summary, subject_details

# Step 3: Generate the Report Card PDF
def generate_pdf(student_id, name, total, average, subject_scores):
    file_name = f"report_card_{student_id}.pdf"
    pdf = SimpleDocTemplate(file_name, pagesize=letter)
    
    # Prepare the elements for the PDF document
    elements = []
    
    # Title Section: Student's Name
    elements.append(Table([[f"Report Card for {name}"]], style=[("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    elements.append(Table([[""]]))  # Adding space
    
    # Summary Table: Total and Average Scores (No Grey Background)
    summary_table = Table([[f"Total Score: {total}", f"Average Score: {average:.2f}"]])
    summary_table.setStyle(TableStyle([
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
    ]))
    elements.append(summary_table)
    elements.append(Table([[""]]))  # Adding space

    # Subject Scores Table
    subject_table_data = [["Subject", "Score"]] + subject_scores
    subject_table = Table(subject_table_data)
    subject_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(subject_table)
    
    # Build the PDF
    pdf.build(elements)
    print(f"Report card PDF generated: {file_name}")

# Main Function to Run the Script
if __name__ == "__main__":
    # Define the Excel file path (ensure it is in the same directory as the script)
    file_path = "student_scores.xlsx"  # Ensure the filename is correct
    
    # Read the Excel data
    data = read_excel(file_path)
    
    if data is not None:
        # Process the data to get total, average, and subject scores
        summaries, details = process_data(data)
        
        # For each student, generate their report card
        for _, student in summaries.iterrows():
            student_id = student["Student ID"]
            name = student["Name"]
            total = student["Total"]
            average = student["Average"]
            
            # Get subject-wise scores for the student
            student_scores = details[details["Student ID"] == student_id][["Subject", "Score"]].values.tolist()
            
            # Generate the report card PDF
            generate_pdf(student_id, name, total, average, student_scores)
