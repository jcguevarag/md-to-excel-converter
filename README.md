# Markdown Table to Excel Converter

A simple yet powerful Python tool that converts Markdown tables into beautifully formatted Excel spreadsheets. Perfect for anyone who drafts reports or documentation in Markdown but needs to share data with colleagues who work primarily in Microsoft Office.

## Why This Tool?

Working with data often means switching between different formats and tools. You might draft a report in Markdown for version control and collaboration, but your stakeholders need the data in Excel for further analysis or presentation. This converter bridges that gap seamlessly, preserving formatting and structure while making your data accessible to Excel users.

## Features

- **Markdown to Excel Conversion**: Transforms the first Markdown table in a `.md` file into a `.xlsx` spreadsheet
- **Formatting Preservation**: Converts `**bold**` and `*italic*` Markdown syntax into proper Excel cell formatting
- **Multi-line Cell Support**: Handles `<br>` tags to create properly wrapped multi-line cells in Excel
- **Smart Column Sizing**: Automatically adjusts column widths for optimal readability
- **CLI Interface**: Simple command-line tool for batch processing and automation
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Prerequisites

- Python 3.8 or higher
- [UV package manager](https://docs.astral.sh/uv/) (recommended) or pip

## Installation

### Using UV (Recommended)

1. **Clone the repository:**
   ```bash
   git clone https://github.com/jcguevarag/md-to-excel-converter.git
   cd md-to-excel-converter
   ```

2. **Install dependencies:**
   ```bash
   uv sync
   ```

### Using pip

1. **Clone the repository:**
   ```bash
   git clone https://github.com/jcguevarag/md-to-excel-converter.git
   cd md-to-excel-converter
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

Convert a Markdown table to Excel using the command-line interface:

```bash
# Using UV (without activating environment)
uv run python -m md_to_excel.converter -i input.md -o output.xlsx

# Using UV (with activated environment)
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
python -m md_to_excel.converter -i input.md -o output.xlsx

# Using pip (with activated virtual environment)
python -m md_to_excel.converter -i input.md -o output.xlsx
```

### Example

Try the converter with the provided sample file:

```bash
# Without activating environment
uv run python -m md_to_excel.converter -i examples/sample_table.md -o output/my_report.xlsx

# Or with activated environment
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
python -m md_to_excel.converter -i examples/sample_table.md -o output/my_report.xlsx
```

This will create a formatted Excel file from the sample Markdown table. The input table:

```markdown
| **Project Name**    | **Status**    | *Lead Developer*   | **Progress (%)** | **Notes** |
|---------------------|---------------|--------------------|------------------|-----------|
| **Project Phoenix** | In Progress   | @developer_one     | 50%              | High priority item.<br>Review by Q3. |
| *Project Gemini*    | Completed     | @developer_two     | 100%             | Signed off 2024-07-15. |
| Project Hydra       | **On Hold**   | @developer_three   | 20%              | Blocked by dependency.<br>Workaround available. |
```

Will become a properly formatted Excel spreadsheet with:
- Bold headers and text where specified
- Italic formatting preserved
- Multi-line cells with proper text wrapping
- Auto-sized columns for readability

### Command Options

- `-i, --inputfile`: Path to the input Markdown file (required)
- `-o, --outputfile`: Path for the output Excel file, must end with `.xlsx` (required)

## How It Works

The converter follows these steps:

1. **Parse Input**: Reads the Markdown file and locates the first table
2. **Extract Structure**: Identifies headers, separator row, and data rows
3. **Process Content**: Handles escaped characters, HTML breaks, and Markdown formatting
4. **Create DataFrame**: Structures the data using pandas for reliable processing  
5. **Generate Excel**: Uses openpyxl to create a formatted spreadsheet with:
   - Bold/italic text formatting
   - Text wrapping for multi-line content
   - Optimized column widths
   - Professional styling

## Project Structure

```
md-to-excel-converter/
├── md_to_excel/
│   ├── __init__.py          # Package initialization
│   └── converter.py         # Main conversion logic
├── examples/
│   └── sample_table.md      # Example Markdown table
├── output/                  # Generated Excel files
├── requirements.txt         # Python dependencies
├── README.md               # This file
└── LICENSE                 # MIT License
```

## Supported Markdown Features

| Feature | Markdown Syntax | Excel Output |
|---------|----------------|--------------|
| Bold text | `**text**` | Bold formatting |
| Italic text | `*text*` or `_text_` | Italic formatting |
| Line breaks | `<br>` or `<br/>` | Multi-line cells with wrapping |
| Escaped pipes | `\|` | Regular pipe characters |
| Table alignment | `:---:`, `---:`, etc. | Left-aligned (Excel default) |

## Limitations

- Converts only the **first** table found in the Markdown file
- Complex nested formatting is not supported
- Table alignment indicators are ignored (all content is left-aligned)
- Links and images within tables are converted to plain text

## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests. 

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

Built with:
- [pandas](https://pandas.pydata.org/) - Data manipulation and analysis
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file creation and formatting
- [UV](https://docs.astral.sh/uv/) - Modern Python package management