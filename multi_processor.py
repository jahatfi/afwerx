"""
Parse AFWERX Proposals for required portions:
1. Budget within constraints
2. DAF customer and end-user
3. TPOCs
"""
import argparse
import os
import pprint
import sys
import time

from pptx import Presentation

import fitz
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from utils import is_directory, is_filename

#from PyPDF2 import PdfFileReader

# If you don't have tesseract executable in your PATH, include the following:
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract'

# TODO Update with mandatory sections, then check for their presence
sections = [
    "Method",
    "Approach"
]

key_phrases = [
    "DAF CUSTOMER",
    "DAF End-User",
    "Digitally signed by",
    "TPOC"
]
# ==============================================================================
# Reference: https://stackoverflow.com/questions/15008758/parsing-boolean-values-with-argparse
def str2bool(v):
    """
    Validates that an argparse argument is a boolean value.
    """
    if isinstance(v, bool):
        return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')
# ==============================================================================
def lower_str(x):
    """
    Converts the provided string to a lowercase string.
    """
    return x.lower()

# ==============================================================================
def process_ppt(file_name):
    """
    Prints the title of every slide
    """
    prs = Presentation(file_name)
    for slide in prs.slides:
        try:
            title = slide.shapes.title.text
        except AttributeError:
            title = slide.shapes[0].text
        print(title)
# ==============================================================================
def process_pdf_page_titles(file_name):
    """
    Prints the title of every page (intended for slides in pdf format)
    """    
    text = ''
    with fitz.open(file_name ) as doc:
        for page_count, page in enumerate(doc):
            #print(f"{page_count}".center(80,"-"))            
            text = page.getText().split('\n')[0]
            print(text)    
# ==============================================================================
def get_total_budget(file_name,
                    max_value=1250000,
                    keyphrase="Total Dollar Amount for this Proposal"):
    threshold = "$"+str(max_value/1000000)+"M"
    with fitz.open(file_name ) as doc:
        for page in doc:
            #print(f"{page_count}".center(80,"-"))
            text_segs = page.getText().split('\n')
            for seg_i, single_text in enumerate(text_segs):
                #print(keyphrase, type(keyphrase))
                if keyphrase.lower() in single_text.lower():
                    #print(single_text)
                    budget_str = text_segs[seg_i+1]
                    print(budget_str)
                    budget_float = float(budget_str.lstrip('$').replace(",",""))
                    if budget_float > max_value:
                        print(f"WARNING!  Proposed budget exceeds ${threshold}!")
                    break
            #print(text)
# ==============================================================================
def process_pdf_sigs_fitz(file_name):
    """
    Print info about digital signatures, as well as text and surrounding
    text containing any of the key phrases defined below
    """
    got_text = False

    with fitz.open(file_name) as doc:
        # Iterate over every page in the doc
        for page in doc:
            text_segs = page.getText().split('\n')
            text_segs = [text.strip() for text in text_segs if text]
            if not text_segs:
                continue
            got_text = True
            print(text_segs)
            # Iterate over every text field
            for seg_i in range(len(text_segs)):
                single_text = text_segs[seg_i]
                for key_phrase in key_phrases:
                    if key_phrase in single_text:
                        for x in text_segs[seg_i-2:seg_i+5]:
                            print(x)
                        seg_i += 5
                        break

    if not got_text:
        print("Failed to get text with fitz parser")

    return got_text

# ==============================================================================
def ocr_cleanup(open_file_handle, open_file_name, files_to_remove):
    print("REmoving")
    for file_name_to_remove in files_to_remove:
        os.remove(file_name_to_remove)
    if open_file_handle:
        open_file_handle.close()
    #if open_file_name:
    #    os.remove(open_file_name)

# ==============================================================================
def ocr_pdf(file_name):
    """
    https://www.geeksforgeeks.org/python-reading-contents-of-pdf-using-ocr-optical-character-recognition/
    """
    print(f"OCR'ing {file_name}.  This could take a minute.")
    # Counter to store images of each page of PDF to image
    image_counter = 1
    files_to_remove = []
    pages = convert_from_path(file_name, 500)
    # Iterate through all the pages stored above
    for page in pages:
        filename = "page_"+str(image_counter)+".jpg"
        # Save the image of the page in system
        page.save(filename, 'JPEG')
        files_to_remove.append(filename)
        image_counter += 1
    '''
    Part #2 - Recognizing text from the images using OCR
    '''
    # Creating a text file to write the output
    outfile = "out_text.txt"

    # Open the file in append mode so that
    # All contents of all images are added to the same file
    f = open(outfile, "a")

    # Iterate from 1 to total number of pages
    for filename in files_to_remove:
        # Recognize the text as string in image using pytesserct
        text = str(((pytesseract.image_to_string(Image.open(filename)))))
        # Finally, write the processed text to the file.
        f.write(text)
        print(f"Page {filename}")

        # The recognized text is stored in variable text
        # text = text.replace('-\n', '')
        text_segs = text.split("\n")
        text_segs = [x.strip() for x in text_segs if x]
        # Iterate over every text field
        for seg_i in range(len(text_segs)):
            single_text = text_segs[seg_i]
            for key_phrase in key_phrases:
                if key_phrase in single_text:
                    for x in text_segs[seg_i-2:seg_i+5]:
                        print(x)
                    print(f"Segment {seg_i}".center(80, "*"))
                    print(single_text.strip())
                    #ocr_cleanup(f, outfile, files_to_remove)
                    #return[0]
            #print(text)

# ==============================================================================
def main(args):
    target_files = set()
    dirs = []
    ppt_extensions = ["ppt", "pptx"]
    valid_extensions = ppt_extensions + ["pdf"]
    if not args.directory:
        args.directory = set()
    if not args.file:
        args.file = set()
    for directory in args.directory:
        dirs.append(directory)

    for file_name in args.file:
        file_extension = file_name.split(".")[-1]
        if file_extension in valid_extensions:
            target_files.add(file_name)
        else:
            print(f"Skipping {file_name} with extension {file_extension}")

    # Count the total number of files to be parsed.
    total_files = len(target_files)
    for source_dir in dirs:
        for (root, _, files) in os.walk(source_dir):
            for file_name in files:
                file_extension = file_name.split(".")[-1]
                if ((args.keyword and all(x in file_name.lower() for x in args.keyword)) or \
                    not args.keyword) and \
                    file_extension.lower() in valid_extensions:
                    target_files.add(root+'/'+file_name)

    print(f"Parsing {len(target_files)} files. This could take a few seconds.")
    # Reset the counter.  It will be incremented as each file is parsed.
    total_files = 0

    # Here is the serial (non-parallel) approach.  Slow, but it works.
    start = time.time()

    for file_name in sorted(target_files):
        print("-"*80)
        print(file_name)
        file_extension = file_name.split(".")[-1]

        if file_extension in ppt_extensions:
            process_ppt(file_name)
            total_files +=1
        else:
            #process_pdf_page_titles(file_name)
            get_total_budget(file_name)
            if not process_pdf_sigs_fitz(file_name):
                if args.ocr:
                    ocr_pdf(file_name)
                else:
                    print(f"Can't parse {file_name}; Consider enabling OCR with -o True")
            total_files += 1

            # TODO: lowercase and compare to remaining list
    end = time.time()
    print(f"{total_files} files in {end-start} seconds")

# ==============================================================================
if __name__ == "__main__":

    # Create the parser and add arguments
    parser = argparse.ArgumentParser()
    failed = False
    # More libraries are loaded if invocation is correct

    # Add an optional argument for the output file,
    # open in 'write' mode and and specify encoding
    parser.add_argument('--file',
                        '-f',
                        action='append',
                        type=is_filename,
                        help="Files to parse.  Must be ppt/ppt/pptx")

    parser.add_argument('--directory',
                        '-d',
                        action='append',
                        type=is_directory,
                        help="Directories to search for files to parse")

    parser.add_argument('--keyword',
                        '-k',
                        action='append',
                        type=lower_str,
                        help="Parse only filenames containing ALL these keywords"
                        )

    parser.add_argument('--ocr',
                        '-o',
                        type=str2bool,
                        default=False,
                        help="Use OCR (Slower, but can parse scanned PDFs)"
                        )

    args = parser.parse_args()
    pprint.pprint(args)


    if not args.file and not args.directory:
        print("Must provide at least one pdf file or directory (will recurse over all directory files.)")
        sys.exit(1)

    original_dir = os.getcwd()
    slide = main(args)
