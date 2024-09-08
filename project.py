import easyocr
import pandas as pd
import re
import math
from PIL import Image, ImageEnhance, ImageFilter
import pdf2image
import numpy as np
from io import BytesIO

# Initialize EasyOCR Reader
reader = easyocr.Reader(['en'], gpu=False)  # Use GPU=True if you have a GPU and want to speed up the process

# Function to preprocess images
def preprocess_image(image):
    # Convert to grayscale
    image = image.convert('L')
    
    # Enhance contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)
    
    # Apply a blur filter to reduce noise
    image = image.filter(ImageFilter.MedianFilter(size=3))
    
    return image

# Function to convert PIL Image to numpy array
def pil_image_to_numpy(image):
    return np.array(image)

# Function to extract text from image using EasyOCR
def extract_text_using_easyocr(image):
    # Preprocess the image
    preprocessed_image = preprocess_image(image)
    
    # Convert PIL Image to numpy array
    image_array = pil_image_to_numpy(preprocessed_image)
    
    # Use EasyOCR to extract text
    results = reader.readtext(image_array)
    
    # Combine the text results
    full_text = " ".join([result[1] for result in results])
    return full_text

# Function to convert PDF to images and use EasyOCR
def extract_text_from_pdf_using_easyocr(pdf_path):
    images = pdf2image.convert_from_path(pdf_path)
    full_text = ""
    for image in images:
        text = extract_text_using_easyocr(image)
        full_text += text + "\n"
    return full_text

# General function to extract data from text using regex
def extract_data_from_text(text):
    data = []
    
    # Regex pattern to handle roll numbers starting with '0801', names, and decimal marks/status
    pattern = re.compile(r"(0801[A-Z\d]*[A-Z]?)\s+([A-Za-z\s]+?)\s+(\d+(\.\d+)?|A|None|Absent)", re.IGNORECASE)
    matches = pattern.findall(text)
    
    for match in matches:
        enrollment_no = match[0].strip()
        name = match[1].strip()
        marks_or_status = match[2].strip() if match[2] else "None"  # Handle missing marks
        
        if 'D' in enrollment_no.upper():  # Special handling for cases with 'D'
            print(f"Enrollment Number with D: {enrollment_no}, Name: {name}, Marks/Status: {marks_or_status}")

        if marks_or_status.replace('.', '', 1).isdigit():
            marks = math.ceil(float(marks_or_status))  # Use math.ceil() to round up
            status = "Present"
        elif marks_or_status.lower() in ["a", "absent", "none"]:
            marks = None
            status = "Absent"
        else:
            marks = None
            status = "Unknown"  # Handle unknown statuses
        
        data.append((enrollment_no, name, marks, status))
    
    return data

# Function to process the data
def process_data(data):
    df = pd.DataFrame(data, columns=['Enrollment No', 'Name', 'Marks', 'Status'])
    
    # Drop rows where 'Enrollment No' or 'Name' is missing
    df.dropna(subset=['Enrollment No', 'Name'], inplace=True)
    
    # Debugging: Print out DataFrame for inspection
    print("DataFrame:\n", df.head(10))  # Print first 10 rows for inspection
    
    # Handling 'Present' status
    df.loc[(df['Marks'].notnull()) & (df['Marks'] >= 7), 'Status'] = 'Pass'
    df.loc[(df['Marks'].notnull()) & (df['Marks'] < 7), 'Status'] = 'Fail'
    
    # Update status for 'Absent'
    df['Status'] = df['Status'].fillna('Absent')
    
    passed = df[df['Status'] == 'Pass']
    failed = df[df['Status'] == 'Fail']
    absent = df[df['Status'] == 'Absent']
    
    return passed, failed, absent

# Function to generate Excel sheets
def generate_excel(passed, failed, absent, output_path):
    with pd.ExcelWriter(output_path) as writer:
        if not passed.empty:
            passed.to_excel(writer, sheet_name="Passed Students", index=False)
        if not failed.empty:
            failed.to_excel(writer, sheet_name="Failed Students", index=False)
        if not absent.empty:
            absent.to_excel(writer, sheet_name="Absent Students", index=False)

# Main function
def main():
    pdf_path = r'C:\Users\abhis\Desktop\WhatsApp Image 2024-09-03 at 17.58.39_87dde7ec.pdf'  # Replace with your actual PDF path
    output_path = r'C:\Users\abhis\Desktop\python\student-marks.xlsx'

    # Attempt to extract text from the PDF using EasyOCR
    text = extract_text_from_pdf_using_easyocr(pdf_path)
    
    if not text.strip():
        print("No data extracted. Please check the PDF format.")
        return
    
    # Extract data using regex
    data = extract_data_from_text(text)
    
    # Print extracted data for debugging
    print("Extracted Data:\n", data)

    passed, failed, absent = process_data(data)
    
    # Print DataFrames for debugging
    print("\nPassed Students:\n", passed)
    print("\nFailed Students:\n", failed)
    print("\nAbsent Students:\n", absent)

    # Count and print the number of students in each category
    total_students = len(pd.concat([passed, failed, absent], ignore_index=True))
    print(f"\nTotal number of students: {total_students}")
    print(f"Number of students who passed: {len(passed)}")
    print(f"Number of students who failed: {len(failed)}")
    print(f"Number of students who were absent: {len(absent)}")
    
    try:
        generate_excel(passed, failed, absent, output_path)
        print("Excel file created successfully.")
    except PermissionError:
        print(f"Permission denied: Unable to write to '{output_path}'. Ensure the file is not open and try again.")
    except Exception as e:
        print(f"Error creating Excel file: {e}")

if __name__ == "__main__":
    main()
