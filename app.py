#!.venv/bin/python

import logging
import os
import sys
import xml.etree.ElementTree as ET
import zipfile


# Configure logging
logging.basicConfig(filename=".data/logger.log", 
                    filemode='a', 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', 
                    level=logging.INFO)
logger = logging.getLogger(__name__)

def main():

    # Check for command-line usage
    if len(sys.argv) != 2:
        sys.exit("Usage: python app.py <file to process>")

    # Check for filetype
    if not zipfile.is_zipfile(sys.argv[1]):
        sys.exit("Provide MS Excel 2007 (or newer) file")

    # Prepare folder structure and extract files from archive
    folder = os.path.split(sys.argv[1])[0]
    # TODO what if given file in current directory (no folder)?
    filename = sys.argv[1].rsplit("/",1)[1]
    extension = filename.rsplit(".",1)[1]
    temp_folder = os.path.join(folder, filename.rsplit(".", 1)[0])
    with zipfile.ZipFile(sys.argv[1], "r") as file:
        file.extractall(path=temp_folder)

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

    # Create new archive
    # TODO check if files were modified and inform user of made changes
    new_filename = filename.rsplit(".",1)[0] + '_updated.' + extension
    with zipfile.ZipFile(os.path.join(folder, new_filename), "w") as new_file:
        # TODO look here https://thispointer.com/python-how-to-create-a-zip-archive-from-multiple-files-or-directory/
        for item in os.listdir(temp_folder):
            new_file.write(os.path.join(temp_folder, item))


if __name__ == '__main__':
    main()