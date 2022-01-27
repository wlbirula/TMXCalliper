# TMX Calliper

A simple tool to compare texts from two translation memory (TMX) files.

![TMXCalliper - screenshot](https://github.com/wlbirula/TMXCalliper/blob/main/docs/tmx_calliper_screenshot.png)

# How it works?

This program compares two translation memory (TMX) files, finds segments that occur in both files, and compares them using Levenshtein's algorithm.

The user gets the summary of the segments and their similarity in XLSX spreadsheet format.

The program may be useful for finding texts processed using machine translation engines (Google Translate etc.).

# Todo

- [ ] Error handling

- [ ] HTML output

# Technologies:
The script uses 3 Python modules:

+ appJar - http://appjar.info/

+ openpyxl - https://pypi.org/project/openpyxl/

+ python-Levenshtein - https://pypi.org/project/python-Levenshtein/



