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
2. Install tesseract: Here's instructions for Windows 10: https://medium.com/quantrium-tech/installing-and-using-tesseract-4-on-windows-10-4f7930313f82


### Combine keywords filters (-k foo -k bar) to process only files that contain "foo" AND "bar"

```bash
python multi_processor.py --help
usage: multi_processor.py [-h] [--file FILE] [--directory DIRECTORY] [--keyword KEYWORD]

optional arguments:
  -h, --help            show this help message and exit
  --file FILE, -f FILE  Files to parse. Must be ppt/ppt/pptx
  --directory DIRECTORY, -d DIRECTORY
                        Directories to search for files to parse
  --keyword KEYWORD, -k KEYWORD
                        Parse only filenames containing ALL of these keywords
```


### References
https://www.geeksforgeeks.org/python-reading-contents-of-pdf-using-ocr-optical-character-recognition/
