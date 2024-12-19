# CIS Benchmark PDF Reader

This Python script extracts data from CIS Benchmark PDF files and converts it into three formats: text (`.txt`), JSON (`.json`), and Excel (`.xlsx`). It processes multiple PDF files in the current directory, extracting relevant benchmark information such as descriptions, rationale, audit commands, remediation steps, and MITRE ATT&CK mappings.

## Prerequisites

Before running the script, ensure you have the following Python packages installed:

- **PyMuPDF (fitz)**: For extracting text from PDF files.
- **Pandas**: For converting JSON data into Excel format.
- **Regular expressions (re)**: Used for pattern matching within the text.

You can install the required packages using `pip`:

```bash
pip install PyMuPDF pandas
```

## Files Processed

The script will process all `.pdf` files in the current directory. For each PDF, it generates three output files:

- **Text file (`.txt`)**: A clean version of the raw extracted text.
- **JSON file (`.json`)**: The extracted data structured as JSON objects.
- **Excel file (`.xlsx`)**: A table format of the extracted data converted from JSON.

The output filenames are based on the name of the input PDF file, with corresponding extensions for each format.

## How the Script Works

### Step 1: PDF Text Extraction
The script reads each PDF file in the current directory and extracts all the text from each page using the PyMuPDF library. The extracted text is saved into a `.txt` file, removing any blank lines.

### Step 2: Text Parsing
The script parses the extracted text and identifies various sections based on the content of the CIS Benchmark. It processes specific details such as:

- **CIS Name**
- **Profile Applicability**
- **Description**
- **Rationale**
- **Impact**
- **Audit**
- **Remediation**
- **Default Value**
- **References**
- **CIS Controls**
- **MITRE ATT&CK Mappings**

The data is organized into dictionaries and saved as a JSON file.

### Step 3: JSON to Excel Conversion
Once the data is structured as JSON, it is converted into a tabular format using Pandas and saved as an Excel file.

### Step 4: Output
For each PDF, the script generates the following output files:

- `file_name.txt` - The extracted text.
- `file_name.json` - The parsed information structured as JSON.
- `file_name.xlsx` - The parsed data formatted into an Excel spreadsheet.

## Running the Script

1. Place your CIS Benchmark PDF files in the same directory as the script.
2. Run the script:

   ```bash
   python cis_benchmark_reader.py
   ```

3. After the script completes, you will see the output files (`.txt`, `.json`, `.xlsx`) for each processed PDF in the same directory.

## Customization

The script is designed to handle the basic structure of CIS Benchmark PDFs. However, if your PDFs have additional sections or different formatting, you may need to adjust the regex patterns or section handling logic to accommodate your specific needs.

### Modifying the Directory
By default, the script looks for PDFs in the current directory (`"."`). If your PDFs are in a different directory, you can change the `directory` variable in the script:

```python
directory = "/path/to/your/pdf/files"
```

## License

This project is licensed under the MIT License.
```

### Markdown Features Used:

1. **Headers** (`#` for main headers, `##` for sub-headers, etc.)
2. **Code blocks** (triple backticks for `bash` and Python code examples).
3. **Lists** (`-` or `*` for unordered lists, `1.` for ordered lists).
4. **Inline code** (using backticks for `filename`, `pip install`, etc.).

Now everything is properly formatted as Markdown. You can copy and paste this into a `README.md` file. Let me know if you'd like any additional tweaks!
