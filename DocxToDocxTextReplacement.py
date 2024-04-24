####################################################################################################################### 
# File: DocxToDocxTextReplacement.py
# Description: This script replaces the old text patterns with the new text patterns in all .docx files in a directory.
# See env_sample.txt for the required environment variables.
# Created: April 2024
# Author(s): Dr. Gail Zhou & GitHub CoPiLot
#######################################################################################################################
import os
import re # Regular Expression
from docx import Document
from dotenv import load_dotenv

# Load .env file
load_dotenv()

# Get old and new text patterns from .env file
OLD_TEXT_1 = os.getenv('OLD_TEXT_1')
NEW_TEXT_1 = os.getenv('NEW_TEXT_1')
OLD_TEXT_2 = os.getenv('OLD_TEXT_2')
NEW_TEXT_2 = os.getenv('NEW_TEXT_2')


def clean_docx(input_file_path, output_dir): 
    # Load the document
    doc = Document(input_file_path)

    # Iterate over the paragraphs
    for paragraph in doc.paragraphs:
        # Iterate over the runs in the paragraph
        for run in paragraph.runs:
            # Replace old text pattern with new one
            run.text = re.sub(OLD_TEXT_1, NEW_TEXT_1, run.text)
            run.text = re.sub(OLD_TEXT_2, NEW_TEXT_2, run.text)

    # Save the cleaned document in the output directory
    output_file_path = os.path.join(output_dir, os.path.splitext(os.path.basename(input_file_path))[0] + "_names_replaced.docx")
    doc.save(output_file_path)
    print(f"New file saved as: {output_file_path}")

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