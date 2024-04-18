
########################################################################################################################################################
# Title: MarkdownToDocx.py
# Author(s): Dr. Gail Zhou & Github Copilot
# Created: April 2024
# Description: This script converts all .docx files in a directory to .md files.
#######################################################################################################################################################
import os  
import shutil
import pandas as pd
from pathlib import Path

from dotenv import load_dotenv

#https://pypi.org/project/pypandoc/
import pypandoc

from colorama import init as colorama_init
from colorama import Fore

# provide a main function to run the script
def main(): 
    colorama_init()
        
    try:
        # Ask for the input directory path
        input_directory_path = input(f"{Fore.CYAN} Enter the input file path: {Fore.RESET}")
        print(f"Input directory: {input_directory_path}")
        # Ask for the output directory path
        output_directory_path = input(f"{Fore.YELLOW} Enter the output file path: {Fore.RESET}")
        print(f"Output directory: {output_directory_path}")
        if os.path.exists(output_directory_path):
            print(f"{Fore.YELLOW}The {output_directory_path} already exists. I will remove it. New directory with same name will be created. {Fore.RESET}")
            shutil.rmtree(output_directory_path)
        os.makedirs(output_directory_path)

        print(f"{Fore.CYAN}\n\n ******************************* Converting files into Docx format ******************************\n\n {Fore.RESET}")       
        # Recursively walk through all files in the input directory
        for root, dirs, files in os.walk(input_directory_path):
            for file_name in files:
                file_extention = Path(file_name).suffix
                if file_name.endswith(('.md', '.json')):
                    input_file_full_path = os.path.join(root, file_name) 

                    file_name_without_extension = Path(file_name).stem 
                    output_file_name = f"{file_name_without_extension}.docx"
                    output_file_full_path = os.path.join(output_directory_path, output_file_name)

                    if not os.path.exists(input_file_full_path):
                        print(f"{Fore.RED}The {input_file_full_path} does not exist.{Fore.RESET}")
                        continue
                              
                    # Expect one of these as input file format:
                    # biblatex, bibtex, commonmark, commonmark_x, creole, csljson, csv, docbook, docx, dokuwiki, endnotexml, epub, fb2, gfm, haddock, 
                    # html, ipynb, jats, jira, json, latex, man, markdown, markdown_github, markdown_mmd, markdown_phpextra, markdown_strict, 
                    # mediawiki, muse, native, odt, opml, org, ris, rst, rtf, t2t, textile, tikiwiki, tsv, twiki, vimwiki
                    
                    if (file_extention in ['.md']):
                        print(f"{Fore.YELLOW} Processing {input_file_full_path}{Fore.RESET}")
                        output = pypandoc.convert_file(input_file_full_path, 'docx', outputfile=output_file_full_path)
                        print(f"{Fore.CYAN}    output: {output_file_full_path} {Fore.RESET}")

                
                    if (file_extention in ['.json', '.txt']):
                        print(f"{Fore.YELLOW} File format not supported. Unable to process {input_file_full_path}{Fore.RESET}")

    
    except Exception as ex:
        print(ex)


if __name__ == '__main__': 
    main()
 


 
