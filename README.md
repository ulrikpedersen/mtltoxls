mtltoxls
========

This is a simple script to parse multiple mtl files and export the data 
into a single spreadsheet.

Installation:
-------------

Clone from github and install dependencies from `requirements.txt`

```
pip install -r requirements.txt
```

Usage
-----

Two arguments are implemented for the CLI: 

 1) Path to directory containing the mtl files
 2) Name of spreadsheet file to write to (will be overwritten if it already exist)

Example:
```
python mtltoxls.py data/Materials_orig_PTC/ result.xlsx
```

