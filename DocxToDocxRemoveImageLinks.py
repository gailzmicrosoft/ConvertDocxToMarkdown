################################################################################################################################
# Title: DocxToDocxRemoveImageLinks.py
# Author(s): Dr. Gail Zhou & GitHub CoPiLot
# Created: April 2024
# Description: This script cleans all .docx files in a directory by removing the lines that contain specific patterns. 
################################################################################################################################
import os  
import re # Regular Expression
from docx import Document


def lineContainsPatterns(line):
    patterns = ["![](media", "width=", "height="]
    for pattern in patterns:
        if pattern in line:
            return True
    return False

def clean_docx(input_file_path, output_dir): 
    # Load the document
    doc = Document(input_file_path)

    # Iterate over the paragraphs
    for paragraph in doc.paragraphs:
        # Split the paragraph into lines
        lines = paragraph.text.split('\n')
        # Remove the backslashes
        lines = [line.replace("\\","") for line in lines]
        # Remove lines that contain the patterns
        lines = [line for line in lines if not lineContainsPatterns(line)]
        # Join the remaining lines back together
        paragraph.text = '\n'.join(lines)

    # Save the cleaned document in the output directory
    output_file_path = os.path.join(output_dir, os.path.splitext(os.path.basename(input_file_path))[0] + "_image_removed.docx")
    doc.save(output_file_path)
    print(f"Cleaned file saved as: {output_file_path}")

def main(): 
    # Ask for the input directory path
    input_dir_path = input("Enter the input directory path: ")
    print(f"Input Directory: {input_dir_path}")

    # Ask for the output directory path
    output_dir_path = input("Enter the output directory path: ")
    print(f"Output Directory: {output_dir_path}")

    # Iterate over all .docx files in the input directory
    for filename in os.listdir(input_dir_path):
        if filename.endswith(".docx"):
            clean_docx(os.path.join(input_dir_path, filename), output_dir_path)

if __name__ == "__main__":
    main()