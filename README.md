## Introduction
DBF2XLSX is a simple and user-friendly tool for converting DBF files to Excel files (XLSX format). It supports single file conversion as well as batch conversion of all DBF files in a directory.

## Features

*   Converts DBF files to XLSX files.
*   Supports custom file encoding or automatic encoding detection.
*   Supports single file or batch directory conversion.
*   Provides detailed logging.

## Usage

### 1. Directly Calling the Python Script

#### Installing Dependencies
Python 3.10 is recommended.

```bash
pip install -r requirements.txt
```

#### Command-Line Arguments

```
usage: dbf_to_excel.py [-h] [-e ENCODING] [-o OUTPUT_DIR] input_path

Converts a specified DBF file or all DBF files in a directory to Excel files.

positional arguments:
  input_path            DBF file or directory containing DBF files

optional arguments:
  -h, --help            show this help message and exit
  -e ENCODING, --encoding ENCODING
                        Encoding of the DBF file (e.g., gbk, utf-8, latin1, etc.; if not specified, it will be automatically detected)
  -o OUTPUT_DIR, --output_dir OUTPUT_DIR
                        Directory for output Excel files, defaults to the directory of the input path
```

#### Examples

*   **Converting a single DBF file:**

    ```bash
    python dbf_to_excel.py input.dbf
    ```

    This will generate an Excel file named `input.xlsx` in the same directory as `input.dbf`.

*   **Converting all DBF files in a directory:**

    ```bash
    python dbf_to_excel.py my_dbf_folder
    ```

    This will generate a corresponding XLSX file for each DBF file in the `my_dbf_folder` directory.

*   **Specifying the output directory:**

    ```bash
    python dbf_to_excel.py input.dbf -o output_folder
    ```

    This will generate `input.xlsx` in the `output_folder` directory.

*   **Specifying file encoding:**

    ```bash
    python dbf_to_excel.py input.dbf -e gbk
    ```

    This will read `input.dbf` using the GBK encoding.

### 2. Using the EXE Version (Windows)

If you are using the packaged EXE version (e.g., `dbf2xlsx.exe`), you can use the command line or drag-and-drop to operate it.

#### Command Line

The command-line arguments are the same as the Python script.

*   **Example:**

    ```bash
    dbf2xlsx.exe input.dbf -o output_folder
    ```

#### Drag and Drop
1. **Single file:** Drag a DBF file and drop it onto the `dbf2xlsx.exe` icon. The program will run automatically and generate the corresponding XLSX file in the same directory as the DBF file.
2. **Multiple files or directory:** Select multiple DBF files or a directory containing DBF files, and drag and drop them onto the `dbf2xlsx.exe` icon. The program will batch convert all files and generate XLSX files in the corresponding directories.

In simple terms: Press `Win + R` to open the Run dialog, type `cmd` to enter the Command Prompt, and drag the downloaded exe file into the window. Enter a space, then drag the dbf file you want to process into the window, press Enter to run it.

## Logging

During the program's execution, log information, including success, warning, and error messages, will be recorded in the `dbf_to_excel.log` file. This helps in troubleshooting.

## Notes

*   If the DBF file uses a special encoding, please use the `-e` parameter to specify the correct encoding.
*   If you encounter a conversion error, please check the log file for detailed information.
*   The EXE version may be falsely reported as a threat by some antivirus software, so please choose to trust the program.

## FAQ

**Q: What should I do if there are garbled characters in the converted Excel file?**

A: This may be caused by an encoding issue. Please try using the `-e` parameter to specify the correct encoding, such as `-e gbk` or `-e utf-8`. If you are unsure of the file encoding, you can omit the `-e` parameter to let the program automatically detect it.

**Q: Why can't the program find the input DBF file?**

A: Please check if the input path is correct and if the file exists. If the path contains spaces, please enclose the path in quotation marks.

## Contact

If you encounter any issues during use, please raise an issus.

## License

This program is licensed under the MIT License.
