"""
Parse AFWERX Proposals for required portions:
1. Budget within constraints
2. Proposed duration
3. DAF customer and end-user (x2 signatures)
4. TPOCs
5. TODO White Paper
6. Subcontractors cost
7. TODO Total manpower (number of employees/workers)
8. Travel costs
"""
import argparse
import os
import sys

from utils import is_directory, is_filename

# TODO Update with mandatory sections, then check for their presence
sections = [
    "Method",
    "Approach"
]

key_phrases = [
    "DAF Customer",
    "DAF End-User",
    "Digitally signed by",
    "TPOC:",
    "TPOCs:",
    "TPOCS:",
    "Technical Point of Contact"
]
# ==============================================================================
# Reference: https://stackoverflow.com/questions/50644066

def pretty_print(df):
    print(display(HTML(df.to_html().replace("\\n","<br>"))))
# ==============================================================================
# Reference: https://stackoverflow.com/questions/15008758/parsing-boolean-values-with-argparse
def str2bool(this_string):
    """
    Validates that an argparse argument is a boolean value.
    """
    if isinstance(this_string, bool):
        return this_string
    if this_string.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif this_string.lower() in ('no', 'false', 'f', 'n', '0'):
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
            text = page.get_text().split('\n')[0]
            print(text)    
# ==============================================================================
def parse_budget(file_name,
                 max_value=1250000,
                 keyphrase="Total Dollar Amount for this Proposal"):
    """
    Parse all relevant fields from budget
    1. Budget within constraints
    2. Proposed duration
    3. Subcontractor Cost
    4. TODO Total manpower (number of employees/workers)
    5. Travel costs
    """
    print("*"*80)
    print(f"Parsing budget: {file_name}")
    headings = [
        "Total Dollar Amount for this Proposal",
        "Total Subcontractor Costs (TSC)",
        "Total Direct Travel Costs (TDT)", # This may appear multiple times, add    
        "Proposed Base Duration (in months)"    
    ]
    text_segs = []
    unique_ts_costs = set()
    ts_cost = 0
    total_travel_cost = 0
    total_proposal_cost = 0
    unique_travel_costs = set()
    result = {}
    threshold = "$"+str(max_value/1000000)+"M"
    with fitz.open(file_name ) as doc:
        for page in doc:
            #print(f"{page_count}".center(80,"-"))
            text_segs += page.get_text().split('\n')

        for seg_i, single_text in enumerate(text_segs):
            #print(single_text)
            #print(keyphrase, type(keyphrase))
            for heading in headings:
                if heading.lower() in single_text.lower():
                    if heading == "Total Dollar Amount for this Proposal":
                        #print(single_text)
                        budget_str = text_segs[seg_i+1]
                        #print(budget_str)
                        total_proposal_cost = float(budget_str.lstrip('$').replace(",",""))
                        result["Total"] = total_proposal_cost
                        if total_proposal_cost > max_value:
                            print(f"WARNING! Proposed budget exceeds ${threshold}!")
                        continue
                    # Sum the total travel costs, careful not to add restated costs
                    elif heading == "Total Direct Travel Costs (TDT)":
                        cost_str = text_segs[seg_i+1]
                        cost_float = float(cost_str.lstrip('$').replace(",",""))
                        if cost_float not in unique_travel_costs:
                            unique_travel_costs.add(cost_float)
                            total_travel_cost += cost_float
                    # Sum the TSC, careful not to add restated costs
                    elif heading == "Total Subcontractor Costs (TSC)":
                        cost_str = text_segs[seg_i+1]
                        cost_float = float(cost_str.lstrip('$').replace(",",""))
                        if cost_float not in unique_travel_costs:
                            unique_ts_costs.add(cost_float)
                            ts_cost += cost_float
                    
                    elif heading == "Proposed Base Duration (in months)":
                        result["Duration (Mo.)"] = single_text.split()[-1].strip()
                        print(single_text)
                    #print(text_segs[seg_i+1])
            #print(text)     
    print(f"Total subcontractor costs: ${ts_cost}")
    print(f"Total travel costs ({len(unique_travel_costs)} unique): ${total_travel_cost}")
    print(f"Total proposal cost: {budget_str}")

    result["TSC"] = ts_cost
    result["TTC"] = total_travel_cost
    return result
# ==============================================================================
def get_total_budget(file_name,
                     max_value=1250000,
                     keyphrase="Total Dollar Amount for this Proposal"):
    """
    Print total budget shown in budget document
    """
    threshold = "$"+str(max_value/1000000)+"M"
    with fitz.open(file_name ) as doc:
        for page in doc:
            #print(f"{page_count}".center(80,"-"))
            text_segs = page.get_text().split('\n')
            for seg_i, single_text in enumerate(text_segs):
                #print(keyphrase, type(keyphrase))
                if keyphrase.lower() in single_text.lower():
                    #print(single_text)
                    budget_str = text_segs[seg_i+1]
                    print(budget_str)
                    budget_float = float(budget_str.lstrip('$').replace(",",""))
                    if budget_float > max_value:
                        print(f"WARNING! Proposed budget exceeds ${threshold}!")
                    break
            #print(text)
# ==============================================================================
def process_pdf_sigs_fitz(file_name):
    """
    Print info about digital signatures, as well as text and surrounding
    text containing any of the key phrases defined below
    """
    print(file_name)
    got_text = False
    result = defaultdict(str)

    with fitz.open(file_name) as doc:
        # Iterate over every page in the doc
        for page in doc:
            text_segs = page.get_text().split('\n')
            text_segs = [text.strip() for text in text_segs if text]
            if not text_segs:
                continue
            got_text = True
            #print(text_segs)
            # Iterate over every text field
            for seg_i, single_text in enumerate(text_segs):
                for key_phrase in key_phrases:
                    if key_phrase in ["DAF End-User", "DAF Customer"]:
                        if single_text == key_phrase:
                            for x in text_segs[seg_i:seg_i+2]:
                                result[key_phrase] += x + '\n'
                                print(x)
                            continue
                    # Remove the .lower() below for more stringent checking
                    #if key_phrase.lower() in single_text.lower():
                    if key_phrase in single_text:

                        if "TPOC" in key_phrase:
                            result[key_phrase] += single_text + "\n"
                            print(single_text)

                        elif key_phrase == "Digitally signed by":
                            print(f"Found {key_phrase}------------------------")
                            #for x in text_segs[seg_i:seg_i+5][::2]:
                            for x in text_segs[seg_i+1:seg_i+3]:
                                if ':' in x:
                                    break
                                result[key_phrase] += x + '\n'
                                print(x)
                            print(f"------------------------------------------")
                            seg_i += 5
                            break       
                        
                        else:

                            #for x in text_segs[seg_i:seg_i+5][::2]:
                            for x in text_segs[seg_i:seg_i+5]:
                                result[key_phrase] += x + '\n'
                                print(x)
                            print(f"------------------------------------------")
                            seg_i += 5
                            break

    if not got_text:
        print("Failed to get text with fitz parser")

    return result

# ==============================================================================
def ocr_cleanup(open_file_handle, open_file_name, files_to_remove):
    """
    Clean up artifacts of OCR - temp files & open file handles
    """
    print("Removing")
    for file_name_to_remove in files_to_remove:
        os.remove(file_name_to_remove)
    if open_file_handle:
        open_file_handle.close()
    if open_file_name:
        os.remove(open_file_name)

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

    #Part #2 - Recognizing text from the images using OCR
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
        for seg_i,_  in enumerate(text_segs):
            single_text = text_segs[seg_i]
            for key_phrase in key_phrases:
                if key_phrase in single_text:
                    print(f"Segment {seg_i}".center(80, "*"))
                    for x in text_segs[seg_i-2:seg_i+5]:
                        print(x)
                    print("*"*80)
                    #print(single_text.strip())
                    #ocr_cleanup(f, outfile, files_to_remove)
                    #return[0]
            #print(text)
    ocr_cleanup(f, outfile, files_to_remove)            

# ==============================================================================
def main():
    target_files = set()
    dirs = []
    ppt_extensions = ["ppt", "pptx"]
    valid_extensions = ppt_extensions + ["pdf"]
    all_info = defaultdict(dict)

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
        #print("-"*80)
        #print(file_name)

        prop_number = re.search(r"(\d{4})", file_name)
        if prop_number:
            prop_number = prop_number.group(1)
            #print(f"Proposal: {prop_number}")
        file_extension = file_name.split(".")[-1]

        if file_extension in ppt_extensions:
            process_ppt(file_name)
            total_files +=1
        else:
            #process_pdf_page_titles(file_name)
            #if "all_forms" in file_name.lower():
            #    all_info[prop_number].update(parse_budget(file_name))
            get_total_budget(file_name)
            
            # Try to get signatures and TPOC data from this PDF
            sig_dict = process_pdf_sigs_fitz(file_name)

            if sig_dict:
                all_info[prop_number].update(sig_dict)
            elif args.ocr:
                ocr_pdf(file_name)
            else:
                print(f"Can't parse {file_name}; Consider enabling OCR with -o True")
            
            total_files += 1

    end = time.time()
    print(f"{total_files} files in {end-start} seconds")
    results = pd.DataFrame.from_dict(all_info, orient="index")
    results.index.name = "Proposal ID"
    pprint.pprint(results.T)
    results.to_csv("proposals.csv")
# ==============================================================================
if __name__ == "__main__":

    # Create the parser and add arguments
    parser = argparse.ArgumentParser()

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

    if not args.file and not args.directory:
        print("Must provide at least one pdf file or directory (will recurse over all directory files.)")
        sys.exit(1)

    # Invocation correct; now load modules
    # Otherwise you force the user to wait for them to load, 
    # then tell them the invocation is incorrect, what a waste of time.
    
    print("Invocation correct, loading modules")
    import pprint
    import time
    import re
    import pandas as pd
    from IPython.display import display, HTML

    from collections import defaultdict
    from pptx import Presentation

    from tabulate import tabulate
    import fitz
    from pdf2image import convert_from_path
    from PIL import Image
    import pytesseract

    pd.set_option('display.width', 1000)
    pd.set_option('display.max_colwidth', 1000)
    # If you don't have tesseract executable in your PATH, include the following:
    #pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract'
    pprint.pprint(args)



    original_dir = os.getcwd()
    # args is global so no need to pass it
    slide = main()
