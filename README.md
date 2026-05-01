# text-based-classification-tool
Tool for analyzing and classifying files based on content using keyword matching

# Academic File Classifier

Simple Python tool for scanning and classifying files based on their content.
Originally designed for organizing university data (faculties, dormitories, services), but can be adapted to other use cases.

---

## Overview

The script goes through all files in a selected directory, extracts text where possible, and assigns each file to a category based on predefined keywords. It also generates a PDF report with a summary of the dataset.

---

## What it does

* scans directories recursively
* reads content from PDF and DOCX files
* normalizes text (removes diacritics)
* matches keywords using regular expressions
* assigns files to categories (e.g. faculties, dormitories, services)
* lists files that could not be classified
* generates a structured PDF report

---

## How to run

```bash
python main.py
```

You will be asked to enter the path to the folder you want to analyze.

---

## Requirements

Install required packages:

```bash
pip install pandas pdfplumber python-docx fpdf unidecode
```

---

## Output

The script creates a PDF report containing:

* total number of files
* total data size
* list of classified files (including matched keywords)
* list of unmatched files

---

## Configuration

Classification is based on a dictionary of keywords:

```python
self.fakulty_data = {
    "Lekarska fakulta": r"...",
    "Koleje a menzy": r"...",
    ...
}
```

You can edit or extend this dictionary to fit your own data.

---

## Notes

* classification is based on simple keyword matching
* results depend on how well the patterns are defined
* not all file types are fully supported

---

## Possible improvements

* better handling of Excel files
* more advanced text analysis
* exporting results to other formats
* adding a user interface

---

## License

MIT License
