# Markdown Table to Excel Converter

A simple yet powerful Python script to convert the first Markdown table found in a file into a beautifully formatted Excel spreadsheet. It preserves common Markdown styling like **bold** and *italic* text and correctly handles multi-line content within cells.

This tool is perfect for anyone who drafts reports or documentation in Markdown and needs to share data with colleagues who work primarily in Microsoft Office.

## Features

*   **Markdown to Excel:** Converts the first Markdown table in a `.md` file to an `.xlsx` file.
*   **Formatting Preservation:** Translates `**bold**` and `*italic*` Markdown syntax into corresponding Excel cell formatting.
*   **Multi-line Cell Support:** Recognizes `<br>` tags to create multi-line cells in Excel with text wrapping enabled.
*   **CLI Interface:** Easy-to-use command-line interface for specifying input and output files.
*   **Smart Column Sizing:** Automatically adjusts column widths in Excel to fit the content.
*   **Cross-Platform:** Built with Python, it runs on Windows, macOS, and Linux (and of course, WSL!).

## Prerequisites

*   Python 3.6+
*   Windows Subsystem for Linux (WSL) if you are on Windows and want to use it in a Linux environment.

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/md-to-excel-converter.git
    cd md-to-excel-converter
    ```

2.  **Set up a virtual environment (recommended):**
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Install the required Python libraries:**
    ```bash
    pip install pandas openpyxl
    ```

## How to Use

The script is run from the command line, specifying the input Markdown file and the desired output Excel file path.

### Syntax

```bash
python -m md_to_excel.converter -i <input_markdown_file> -o <output_excel_file>
```

### Example

1.  Navigate to the project's root directory.

2.  Run the script on the provided sample file located in the `examples/` directory:
    ```bash
    python -m md_to_excel.converter -i examples/sample_table.md -o output/report.xlsx
    ```

3.  **Check the output:** A new `output` directory will be created with a `report.xlsx` file inside. When you open it, you will see a formatted table similar to this:

    | **Project Name**      | **Status**      | *Lead Developer*   | **Progress (%)** | **Notes**                                                                 |
    | :-------------------- | :-------------- | :----------------- | :--------------- | :------------------------------------------------------------------------ |
    | **Project Phoenix**   | In Progress     | @developer\_one    | 50%              | This is a high-priority item.<br>Requires final review by Q3.              |
    | *Project Gemini*      | Completed       | @developer\_two    | 100%             | The deliverable was signed off on 2024-07-15.                             |
    | Project Hydra         | **On Hold**     | @developer\_three  | 20%              | Blocked by dependency X.<br>The team has a workaround for C\|D.            |
    | Project Apollo        | Not Started     | @developer\_four   | 0%               | Scheduled to begin next sprint.                                           |

    *Note: The actual Excel file will have formatted text (bold/italic) and proper cell wrapping, not the raw Markdown.*

## How It Works

The script performs the following steps:
1.  **Parses Command-Line Arguments:** Uses `argparse` to get the input and output file paths.
2.  **Reads the Markdown File:** It reads the file line by line to find the first table.
3.  **Identifies Table Structure:** It looks for a header row followed by a separator line (`|---|---|`).
4.  **Parses Rows:** Each table row is split into cells, handling escaped pipe characters (`\|`).
5.  **Processes Cell Content:** It converts `<br>` tags to newlines (`\n`).
6.  **Creates a Pandas DataFrame:** The headers and data rows are loaded into a DataFrame for structured handling.
7.  **Writes to Excel:** Using the `openpyxl` library, it writes the DataFrame to an `.xlsx` file.
8.  **Applies Formatting:** It iterates through the cells, removes Markdown formatting characters (`**`, `*`), and applies the corresponding `Font` and `Alignment` styles in Excel.
9.  **Adjusts Column Widths:** Finally, it calculates the necessary width for each column to ensure readability.

## Contributing

Contributions are welcome! If you have ideas for improvements or find a bug, please open an issue or submit a pull request.