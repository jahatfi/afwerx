'''
I reused alot of the boilerplate regarding argument parsing and 
iterating / counting pdf files from another project I'm developing.
'''

import os
import pprint
import re
import shutil
import sys
import time
from copy import deepcopy
from io import StringIO
from subprocess import Popen, run

import fitz
import numpy as np
from matplotlib import cm
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage

# Globals
too_short = 0
bad_format = 0
g_exit_now = False
engines = ['pdfminer', 'fitz']
colors = cm.tab20

sections = [
    "adverse_events",
    "baseline", # Also common after the conclusion
    "conclusions",
    "study_size",
    "study_type",
    "dosage",
    "intervention_duration",
    "study_duration",
    "exclusion_criteria",
    "inclusion_criteria",
    "methods",
    "objectives",
    "outcome_measures",
    "results",
    "study_groups"
]

#===============================================================================
def highlight_last(page, text:str, r, g, b):
    all_text = []
    text_instances = page.searchFor(text)
    all_text.append(page.get_text().strip())
    ### HIGHLIGHT

    inst = text_instances.pop()
    print(f"Highlighting '{inst}'")
    highlight = page.addHighlightAnnot(inst)
    highlight.setColors({"stroke":(r, g, b)})
    highlight.update()
#===============================================================================
def find_spanning_text(all_text:list, text:str):
    # Now find all instances of the desired text that spans pages
    giant_text = " ".join(all_text)
    findings = [[m.start(), m.end()] for m in re.finditer(text, giant_text)]

    if not findings:
        return None, None, None    

    total_length = 0
    cumulative_length = 0
    first_portion = []
    middle = []
    last_portion = []
    page_index = 0
    while page_index < len(all_text):
        total_length += len(all_text[page_index])    
        if findings[0][0] > total_length:
            page_index += 1 
            continue

        first = all_text[page_index][findings[0][0]:]
        if first:
            cumulative_length =  len(first)         
            first_portion = [page_index, first]

  

        try:
            # The first (remaining) finding is on this page
            if findings[0][1] < total_length:
                # The end of this text is on the next page
                last_portion = [page_index, text[findings[0][0]:]]
                # Pop this finding now that we've record where it's found
                findings.pop(0)
                cumulative_length = 0
            else:
                # This entire page contains desired text
                this_page = page_index + 1
                preslice = findings[0][0]+len(first_portion)
                this_text = text[preslice:preslice+len(all_text[page_index+1])+1].strip()
                middle = [this_page, this_text]
                # Now update the start of the current finding by updating it
                # to point to the next page                
                findings[0][0] += len(all_text[page_index+1]) + cumulative_length
                cumulative_length = findings[0][0]
                # Skip the next page
                total_length += len(all_text[page_index+1]) 
                page_index += 1

        except IndexError:
            pass

        # Exit loop if no more findings to record
        if not findings:
            break    
        # Continue to next iteration
        page_index += 1 

    return (first_portion, middle, last_portion)
#==============================================================================
# https://pymupdf.readthedocs.io/en/latest/tutorial.html
def highlight(in_filename, out_filename, text, r=1,g=1,b=0):
    ### READ IN PDF
    print(f"Setting highlighter color to ({r},{g},{b})")
    r = float(r)
    g = float(g)
    b = float(b)
    doc = fitz.open(in_filename)
    page = doc[0]
    made_highlight = False
    all_text = []
    for page_index, page in enumerate(doc):
        print(f"Working on page #{page_index}")
        ### SEARCH
        text_instances = page.searchFor(text)
        all_text.append(page.get_text().strip())
        ### HIGHLIGHT

        for inst in text_instances:
            made_highlight = True
            print(f"Highlighting '{inst}'")
            highlight = page.addHighlightAnnot(inst)
            highlight.setColors({"stroke":(r, g, b)})
            highlight.update()

    if not made_highlight:
        print("No highlights made - checking for spanning text...")
        first_portion, middle, last_portion = find_spanning_text(all_text, text)
        if first_portion:
            print(f"first_portion {first_portion}")
            print(f"middle: {middle}")
            print(f"last_portion {last_portion}")

            for portion_page_number, portion_text in [first_portion, middle, last_portion]:   
                portion_page = doc[portion_page_number]
                highlight_last(portion_page, portion_text, r,g,b)
        made_highlight = True

    ### OUTPUT
    if made_highlight:
        print(f"Saving highlighted pdf to {out_filename}")
        doc.save(out_filename, garbage=4, deflate=True, clean=True)
    else:
        print("No highlights made :(")
    print("Done.")
# ==============================================================================
def check_for_signal():
    """
    If SIG INT detected, set a global stop flag.
    """
    if g_exit_now:
        print("Exited early due to SIGINT signal.")
        sys.exit(1)
# ==============================================================================
# Catch the INTERRUPT signal.  Exit by setting a global g_main_done flag,
# and given processes time to close cleanly if necessary.
def sigint_handler(sig, frame):
    global g_exit_now
    g_exit_now = True
#===============================================================================
# This function condenses whitespace in text.  If it can't, or if 
# the data is too short (indicating an error in parsing or not a 
# complete paper), simply return NaN.
def clean_text(original_text):
    global too_short
    global bad_format 
    text = deepcopy(original_text)
    print("="*80)
    if text == np.NaN:
        return np.NaN
    try:
        newline_count = text.count("\n")
        space_count = text.count(" ")
        
        newline_ratio = float(newline_count)/len(text)
        space_ratio = float(space_count)/len(text)

        print(f"{newline_ratio},{space_ratio}")
        # If the ratios are out of these ranges,
        # paper likely wasn't parsed properly
        if newline_ratio > .3 or space_ratio < .005:
            text = np.NaN
            bad_format +=1 
            print("Dropping row")

        elif newline_ratio > .2  or space_ratio < .05 or space_ratio > .23:
            
            text = text.replace("\n\n","\t")
            text = text.replace("\n"," ")
            text.replace("\t"," ")
            print("-"*80)
            text = text.replace('\n',' ').replace("- ", "")
            
            text = ' '.join(text.split())            
            
            #print(text)
            newline_count = text.count("\n")
            space_count = text.count(" ")

            newline_ratio = float(newline_count)/len(text)
            space_ratio = float(space_count)/len(text)
            print(f"UPDATED RATIOS: {newline_ratio},{space_ratio}")
            
            if space_ratio > .25:
                text = np.NaN
                bad_format +=1 
            
        else:
            text = text.replace('\n',' ').replace("- ", "")
            text = text.replace('\\n', ' ')
            text = ' '.join(text.split())        
            
        # After removing extra whitespace, if the len is less than 5k,
        # this generally indicates bad parsing or not a full article
        if len(text) < 5000:
            print(f"text too short ({len(text)} characters):")
            #print(text)
            too_short += 1
            text = np.NaN
            
    except AttributeError as e:
        print(e, original_text)
    finally:
        if  isinstance(text, str):
        #    print(text)
        #    time.sleep(.1)
            try:
                text = text.replace("vs.", "versus")
            except Exception as e:
                print(e)
                print(text)
        return text
#===============================================================================
# This class was left behind by the Fall 2020 team
class PdfConverter:
    def __init__(self, file_path):
        self.file_path = file_path
    
    # convert pdf file to a string which has space among words 
    def convert_pdf_to_txt(self):
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        #codec = 'utf-8'  # 'utf16','utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, laparams=laparams)
        fp = open(self.file_path, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = b""
        maxpages = 0
        caching = False
        pagenos = set()
        for page in PDFPage.get_pages(
            fp, 
            pagenos, 
            maxpages=maxpages, 
            password=password, 
            caching=caching, 
            check_extractable=False):
            interpreter.process_page(page)
            
        fp.close()
        device.close()
        str = retstr.getvalue()
        retstr.close()
        return str
    
    # convert pdf file text to string and save as a text_pdf.txt file
    def save_convert_pdf_to_txt(self):
        content = self.convert_pdf_to_txt()
        txt_pdf = open('text_pdf.txt', 'wb')
        txt_pdf.write(content.encode('utf-8'))
        txt_pdf.close()

# This function was left behind by the Fall 2020 team
def extract_text_pdfminer(path):
    # extract the PDF text
    pdfConverter = PdfConverter(file_path=path)
    return pdfConverter.convert_pdf_to_txt()

def extract_text_fitz(filepath):
    text = ''
    with fitz.open(filepath ) as doc:
        for page in doc:
            text+= page.getText()
    text = text.replace("-\n", "")
    text = text.replace("\n", " ")
    return text
#===============================================================================
def is_filename(filename:str):
    try:
        f = open(filename)
    except FileNotFoundError:
        raise argparse.ArgumentTypeError(f"Cannot open file'{filename}'")
    finally:
        try:
            f.close()
        except Exception:
            pass
    return filename
#===============================================================================
def is_directory(dir:str):
    if not os.path.isdir(dir):
        raise argparse.ArgumentTypeError(f"'{dir}' is not a valid directory")
    return os.path.abspath(dir)
# ==============================================================================
def is_section(section):
    if section not in sections:
        raise argparse.ArgumentTypeError(f"'{section}' is not a valid section")
    return section
# ==============================================================================
def is_engine(engine):
    if engine not in engines:
        raise argparse.ArgumentTypeError(f"'{engine}' is not a valid extraction engine")
    return engine        
# ==============================================================================
# Reference: https://stackoverflow.com/questions/15008758/
def str2bool(v):
    if isinstance(v, bool):
       return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')    

#===============================================================================
def process_single_pdf(file_name:str):
    # Extract text with engine; dump to file
    if args.engine == 'fitz':
        text = extract_text_fitz(file_name)
    elif args.engine == "pdfminer":
        text = extract_text_pdfminer(file_name)
    #continue
    # Run regexes, in order, appending all to the same file
    pdf_text_file_name =  args.output_dir + '/all_text_'+file_name.strip(".pdf")+'.txt'
    extracted_sections_file_name = args.output_dir + '/extracted_sections_'+file_name.strip(".pdf")+'.txt'
    print(f"Creating {pdf_text_file_name}")
    with open(pdf_text_file_name, 'w') as extracted_text_file:
        extracted_text_file.write(text)

    print("Running regex extraction.  This could take several minutes.")

    p = run([regex_script, pdf_text_file_name, extracted_sections_file_name])
    if p.returncode != 0:
        print(f"Oops, error running regex script on {file_name_with_path}")
        print(p.stderr)
        print(p.stdout)
        print(p.returncode)
        #print(p.check_returncode())
        return

    if not args.summarize and not args.colorize:
        return

    # If summarizer or colorizer (see above):
    # Open the file with extracted sections:
    print(f"Attempting to colorize {file_name}...")
    with open(extracted_sections_file_name, 'r') as text_file:
        hl_title = args.output_dir + "/highlighted_" + file_name
        prev_line = ""
        highlights_made = False
        # Iterate over each extracted section in the .txt file
        for line in text_file:
            if line == prev_line:
                continue

            if line.startswith("#"):
                section = line.strip('#').strip().lower()
                prev_line = line
                continue 

            if args.colorize:
                # Highlight the findings
                r, g, b, _= colors(sections.index(section)/len(sections))
                if os.path.isfile(hl_title):
                    if(highlight(hl_title, hl_title, line, r, g ,b)):
                        highlights_made = True
                elif highlight(file_name, hl_title, line, r, g ,b):
                    highlights_made = True

            if args.summarize:
                # TODO: Write this line to a section-specific tfrecord 
                pass

            prev_line = line

        # Summarize section as commanded
        # Colorize PDF if colorizer selected

    end = time.time()
    if not highlights_made:
        print("No highlights made.")
        os.remove(hl_title)
# ==============================================================================
def main(args):
    pdf_files = []
    dirs = []
    if not args.directory:
        args.directory = set()
    if not args.file:
        args.file = set()
    for dir in args.directory:
        dirs.append(dir)

    for each_file in args.file:
        if each_file.endswith(".pdf") and not each_file.startswith("highlight"):
            pdf_files.append(each_file)     

    # Count the total number of files to be parsed.
    total_files = len(pdf_files)
    for source_dir in dirs:
        for (root, _, files) in os.walk(source_dir):
            for file_name in files:
                if file_name.endswith(".pdf") and not file_name.startswith("highlight"):
                    pdf_files.append(os.path.join(root, file_name))

    print(
        f"Parsing {len(pdf_files)} .pdf files total. This could take a few seconds."
    )
    # Reset the counter.  It will be incremented as each file is parsed.
    total_files = 0

    # Here is the serial (non-parallel) approach.  Slow, but it works.
    start = time.time()

    for file_name_with_path in sorted(pdf_files):
        dir = os.path.dirname(file_name_with_path)
        file_name = os.path.basename(file_name_with_path)
        print(f"cd {dir}")
        #shutil.copyfile(regex_script, dir+'\\'+regex_script)
        if dir:
            os.chdir(dir)
        print(file_name)
        total_files += 1
        process_single_pdf(file_name)
    print(f"{total_files} files in {end-start} seconds")
# ==============================================================================
if __name__ == "__main__":
    import argparse

    # Create the parser and add arguments
    parser = argparse.ArgumentParser()
    failed = False
    # More libraries are loaded if invocation is correct

    # Add an optional argument for the output file,
    # open in 'write' mode and and specify encoding
    parser.add_argument('--output_dir', 
                        '-o', 
                        type=is_directory, 
                        default="output", 
                        help="Name of directory to store results.")

    parser.add_argument('--engine', 
                        '-e', 
                        help='fitz (default) or pdfminer',
                        type=is_engine,
                        default='fitz')

    parser.add_argument('--summarize',
                        '-s',
                        action='append',
                        type=is_section,
                        help=f"Use pegasus to summarize any of the following sections: {sections}")

    parser.add_argument('--colorize',
                        '-c',
                        type=str2bool,
                        default="False",
                        help="(Boolean: False by default)" \
                        + "Attempt to highlight extracted text in source pdf")

    parser.add_argument('--file',
                        '-f',
                        action='append', 
                        type=is_filename)

    parser.add_argument('--directory','-d', action='append', type=is_directory)                           
    args = parser.parse_args()
    pprint.pprint(args)


    if not args.file and not args.directory:
        print("Must provide at least one pdf file or directory (will recurse over all directory files.)")
        sys.exit(1)
    original_dir = os.getcwd()
    regex_script = original_dir+"/regex_extraction.sh"
    main(args)
