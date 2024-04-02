
########################################################################################################################################################
# Title: ConvertDocxToMD.py
# Author(s): Dr. Gail Zhou
# Contributors: CoPiLot
#
# April 2024
# Introduction: This script converts all .docx files in a directory to .md files.
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

def main(): 
    colorama_init()
        
    try:
        # Get configuration settings 
        load_dotenv()

        # Normally the clean_md flag is set to True. If it is set to False, the script will not clean the markdown files.
        clean_md = os.getenv("CLEAN_MD")
        if (clean_md in ["True", "true", "TRUE", "T", "t", "Y", "y", "Yes", "yes", "YES"]):
            clean_md_flag = True
        else:
            clean_md_flag = False

        # Ask for the input directory path
        input_directory_path = input(f"{Fore.CYAN} Enter the input directory path for .docx files: {Fore.RESET}")
        print(f"Input directory: {input_directory_path}")
        # Ask for the output directory path
        raw_md_directory_path = input(f"{Fore.YELLOW} Enter the output directory path for RAW .md files: {Fore.RESET}")
        print(f"Output directory for raw MD files: {raw_md_directory_path}")
        if os.path.exists(raw_md_directory_path):
            print(f"{Fore.RED}The {raw_md_directory_path} already exists. It is removed. New directory with same name is now created. {Fore.RESET}")
            shutil.rmtree(raw_md_directory_path)
        os.makedirs(raw_md_directory_path)


        # Ask for the output directory path
        if clean_md_flag:
            clean_output_directory_path = input(f"{Fore.CYAN} Enter the output directory path for CLEAN .md files: {Fore.RESET}")
            print(f"Output directory for CLEAN .md files: {clean_output_directory_path}")
            if os.path.exists(clean_output_directory_path):
                print(f"{Fore.RED}The {clean_output_directory_path} already exists. It is removed. New directory with same name is now created.{Fore.RESET}")
                shutil.rmtree(clean_output_directory_path)
            os.makedirs(clean_output_directory_path)

        print(f"{Fore.CYAN}\n\n ******************************* Converting .docx files into Markdown Files in RAW format ******************************\n\n {Fore.RESET}")       
        # Recursively walk through all files in the input directory
        for root, dirs, files in os.walk(input_directory_path):
            for file_name in files:
                file_extention = Path(file_name).suffix
                if file_name.endswith(('.docx', '.txt')):
                    input_file_full_path = os.path.join(root, file_name) 

                    file_name_without_extension = Path(file_name).stem 
                    output_file_name = f"{file_name_without_extension}_raw.md"
                    output_file_full_path = os.path.join(raw_md_directory_path, output_file_name)

                    if not os.path.exists(input_file_full_path):
                        print(f"{Fore.RED}The {input_file_full_path} does not exist.{Fore.RESET}")
                        continue
                              
                    # Expect one of these as input file format:
                    # biblatex, bibtex, commonmark, commonmark_x, creole, csljson, csv, docbook, docx, dokuwiki, endnotexml, epub, fb2, gfm, haddock, 
                    # html, ipynb, jats, jira, json, latex, man, markdown, markdown_github, markdown_mmd, markdown_phpextra, markdown_strict, 
                    # mediawiki, muse, native, odt, opml, org, ris, rst, rtf, t2t, textile, tikiwiki, tsv, twiki, vimwiki
                    
                    if file_extention == '.docx':
                        print(f"{Fore.YELLOW}Processing {input_file_full_path}{Fore.RESET}")
                        output = pypandoc.convert_file(input_file_full_path, 'md', outputfile=output_file_full_path)
                        print(f"{Fore.CYAN}    output: {output_file_full_path} {Fore.RESET}")

                    if file_extention == '.txt':
                        print("")
                        #print(f"{Fore.RED} Unble to take .txt files yet. {file_name} ignored.  {Fore.RESET}")

        # Clean the markdown files if the clean_md flag is set to True
                        
        print(f"{Fore.CYAN}\n\n ******************************* Cleaning up Markdown Files *******************************\n\n{Fore.RESET}")       
        if clean_md_flag:
            # Recursively walk through all files in the input directory
            for root, dirs, files in os.walk( raw_md_directory_path):
                for file_name in files:
                    if file_name.endswith('.md'):
                        file_name_without_extension = Path(file_name).stem 
                        clean_md_file_name = f"{file_name_without_extension}_clean.md"
                        
                        input_file_full_path = os.path.join(root, file_name) 
                        output_file_full_path = os.path.join(clean_output_directory_path, clean_md_file_name)
                        
                        print(f"{Fore.CYAN} input: {input_file_full_path} {Fore.RESET}")
                        
                        # Process the file
                        with open(input_file_full_path, "r") as f:
                            lines = f.readlines()
                        with open(output_file_full_path, "w") as f:
                            for line in lines:
                                if not ("![](media" in line or "width=" in line or "height=" in line):
                                    f.write(line)
                        
                        print(f"{Fore.GREEN}      output: {output_file_full_path} {Fore.RESET}") 
                        print("")
        
        
        print(f"{Fore.GREEN} ******************************* Done. *******************************\n {Fore.RESET}") 
    
    except Exception as ex:
        print(ex)

if __name__ == '__main__': 
    main()
 


 
