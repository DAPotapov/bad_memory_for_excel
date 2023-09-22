#!.venv/bin/python

import os
import sys
import xml.etree.ElementTree as ET
import zipfile

def main():

    # Check for command-line usage
    if len(sys.argv) != 2:
        sys.exit("Usage: python app.py <file to process>")

    # Check for filetype
    if not zipfile.is_zipfile(sys.argv[1]):
        sys.exit("Provide MS Excel 2007 (or newer) file")

    # Prepare folder structure and extract files from archive / dirname? / .isdir?
    folder = os.path.split(sys.argv[1])[0]
    # TODO what if given file in current directory (no folder)? / os.path.split / .isfile
    filename = sys.argv[1].rsplit("/",1)[1]
    extension = filename.rsplit(".",1)[1] # splitext
    temp_folder = os.path.join(folder, filename.rsplit(".", 1)[0])
    with zipfile.ZipFile(sys.argv[1], "r") as file:
        file.extractall(path=temp_folder)

    sheets_list = []
    # Get into extracted file-tree and proceed with sheet's files
    sheets_path = os.path.join(temp_folder, 'xl', 'worksheets')
    if os.path.exists(sheets_path):

        # Get content of sheet and search for needed tag, remove it if found             
        for sheet in os.listdir(sheets_path):
            if os.path.isfile(os.path.join(sheets_path, sheet)):       
                sheet_contents = ET.parse(os.path.join(sheets_path, sheet))
                root = sheet_contents.getroot()
                tag = root.find(".//{*}sheetProtection")
                if tag != None:
                    root.remove(tag)
                    sheet_contents.write(os.path.join(sheets_path, sheet))
                    sheets_list.append(sheet)

    if sheets_list:
        # Create new archive
        # TODO check if files were modified and inform user of made changes
        new_filename = filename.rsplit(".",1)[0] + '_updated.' + extension
        with zipfile.ZipFile(os.path.join(folder, new_filename), "w", compression=zipfile.ZIP_DEFLATED) as new_file:
            
            # Walk through subfolder and pack contents to archive relatively
            for folder_name, subfolders, filenames in os.walk(temp_folder):
                for filename in filenames:

                    # Path to file to pack
                    file_path = os.path.join(folder_name, filename)

                    # Relative file path to store in archive
                    relpath = os.path.join(os.path.relpath(folder_name, temp_folder), filename)
                    new_file.write(file_path, relpath)
        print(f"New file created: '{new_filename}' and includes changed sheets: {','.join(sheets_list)}.")
    else:
        print("No files changed.")
    

if __name__ == '__main__':
    main()