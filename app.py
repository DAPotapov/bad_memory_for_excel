#!.venv/bin/python

import logging
import os
import sys
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

    if not zipfile.is_zipfile(sys.argv[1]):
        sys.exit("Provide MS Excel 2007 (or newer) file")

    print(sys.argv[1])
    folder = os.path.split(sys.argv[1])[0]
    filename = sys.argv[1].rsplit("/",1)[1]
    temp_folder = os.path.join(folder, filename.rsplit(".",1)[0])
    print(folder, filename, temp_folder)
    with zipfile.ZipFile(sys.argv[1], "r") as file:
        file.extractall(path=temp_folder)
        # print(result)
    #     output_file = input_file.rsplit('.',1)[0] + '.csv'


if __name__ == '__main__':
    main()