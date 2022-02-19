# Tool to aid evaluation of AFWERX Proposals

So far it can:
1. Print title of every slides of pdf/ppt/pptx
    1. TODO Verify presence of required slides
2. Parse budget line item, e.g. "Total Dollar Amount for this Proposal" in a well-formatted PDF.   
    1. Print a warning if parsed value exceeds some threshold.

3. Extract Contact info:
    1. PDF digital Signatures
    2. DAF Customer/End-User info
    3. TPOC info

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
