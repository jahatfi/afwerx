# Tool to aid evaluation of AFWERX Proposals

So far it can:
1. Print title of every slides of pdf/ppt/pptx
    1. TODO Verify presence of required slides
2. Parse budget line item, e.g. "Total Dollar Amount for this Proposal" in a well-formatted PDF.   
    1. Print a warning if parsed value exceeds some threshold.

3. Extract Contact info (using OCR* if necessary):
    1. PDF digital Signatures
    2. DAF Customer/End-User info
    3. TPOC info

### Installation (I need to test this)
1. python modules: `pip install -r requirements.txt`  
2. Install tesseract **only if using Optical Character Recognition (OCR) (for scanned PDFS)**:  
    Here's instructions for Windows 10: https://medium.com/quantrium-tech/installing-and-using-tesseract-4-on-windows-10-4f7930313f82


### Combine keywords filters (-k foo -k bar) to process only files that contain "foo" AND "bar"

```bash
python  multi_processor.py -h
usage: multi_processor.py [-h] [--budget-file BUDGET_FILE] [--questions-file QUESTIONS_FILE] [--file FILE] [--directory DIRECTORY] [--keyword KEYWORD] [--ocr OCR] [--out OUT]

optional arguments:
  -h, --help            show this help message and exit
  --budget-file BUDGET_FILE, -b BUDGET_FILE
                        Will parse files with this keyword as budget documents (default: budget)
  --questions-file QUESTIONS_FILE, -q QUESTIONS_FILE
                        Will parse files with this keyword for duration and firm+proposal questions (default: all_forms)
  --file FILE, -f FILE  Files to parse. Must be pdf/ppt/pptx (default: None)
  --directory DIRECTORY, -d DIRECTORY
                        Directories to search for files to parse (default: None)
  --keyword KEYWORD, -k KEYWORD
                        Parse only filenames containing ALL these keywords (default: None)
  --ocr OCR, -o OCR     Use OCR (Slower, but can parse scanned PDFs) (default: False)
  --out OUT             Save results to this file (will be a csv.) (default: proposals.csv)
```
### Examples
Process only proposal #1234
```bash
python multi_processor.py -d proposals_directory -k 1234
```

Process only files with "budget" in the filename 
```bash
python multi_processor.py -d proposals_directory -k budget
```

Process only files with "budget" in the filename in proposal 1234
```bash
python multi_processor.py -d proposals_directory -k budget -k 1234
```



### References
https://www.geeksforgeeks.org/python-reading-contents-of-pdf-using-ocr-optical-character-recognition/
