# Text based classification tool

A small Python script for scanning files in a folder and sorting them based on their content.

It was originally created for organizing university-related files (faculties, dormitories, services), but the idea can be reused for other types of data.

## What it does

The script goes through all files in a selected directory, tries to read their content, and looks for keywords.
If a match is found, the file is assigned to a category. If not, it is listed as unmatched.

It also creates a simple PDF report with an overview of the dataset.

## Supported files

The script works best with PDF and DOCX files, where it can extract text content.  
Other file types (such as Excel) are still included in the scan, but their content is not analyzed.

## Running the script

```bash id="r8d3ka"
python main.py
```

After running, you will be asked to enter the path to the folder you want to analyze.

## Requirements

Install dependencies with:

```bash
pip install -r requirements.txt
```

Alternatively, you can install them manually:

```bash
pip install pandas pdfplumber python-docx fpdf2 unidecode
```

## Output

The script generates a PDF report that includes:

* total number of files
* total size of the dataset
* list of classified files (with matched keywords)
* list of files that were not classified

## Configuration

Categories are defined in a dictionary inside the script.
You can adjust the keywords or add your own depending on your data.

## Notes

The classification is based on simple keyword matching, so the results depend on how well the patterns are defined. Text content is extracted only from PDF and DOCX files, while other file types are included in the analysis but their content is not processed.

## License

MIT

