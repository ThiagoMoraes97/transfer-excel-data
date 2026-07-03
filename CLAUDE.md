# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python tool for transferring data between Excel spreadsheets while preserving formatting. The project has two interfaces:

- **Web interface** (`web_app.py`): Modern Flask application with visual column mapping UI
- **Desktop interface** (`main.py`): Tkinter GUI application

## Development Commands

### Setup and Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run web interface (recommended)
python web_app.py
# Opens at http://127.0.0.1:5000

# Run desktop interface
python main.py
```

### Testing

No test suite is currently set up. Manual testing is done through:
- Web interface: Upload Excel files and test column mappings
- Desktop interface: Launch the GUI and test file transfers

## Architecture

### Core Business Logic

Both `main.py` and `web_app.py` contain similar business logic for Excel manipulation:

- **Column mapping**: Simple (`A:B`) and combined (`A+C:D`) mappings with configurable separator
- **Filter system**: Advanced filtering with operators (empty, contains, numeric comparisons, list exclusion)
- **Negative values**: Option to convert numeric values to negative for expense documents
- **VBA preservation**: Maintains macros in `.xlsm` files via `keep_vba` parameter
- **Format preservation**: Only writes cell values, maintains existing formatting

### Web Interface Architecture

`web_app.py` is a Flask application with:

- **Upload cache**: In-memory storage with 30-minute TTL for uploaded files
- **Column detection**: Automatic scanning of sheets to identify used columns with headers and sample data
- **REST API**: Endpoints for upload (`/api/upload`), sheet column listing (`/api/sheet-columns`), and data transfer (`/api/transfer`)
- **Frontend**: Vanilla JavaScript app in `static/app.js` with custom styling

### Desktop Interface Architecture

`main.py` is a Tkinter application with:

- **File dialogs**: Native OS file pickers for source/destination/output files
- **Business logic functions**: Shared parsing and data manipulation functions
- **UI components**: Form-based interface with text inputs for mappings

### Key Functions and Patterns

**Column parsing and validation**:
- `normalizar_coluna()` / `para_coluna_excel()`: Convert between column letters/numbers and validate format
- `parsear_mapeamentos()` / `_parse_mappings()`: Parse mapping strings (e.g., `A:B,D:F`)

**Data copying logic**:
- `copiar_colunas()` / `_copy_data()`: Main transfer functions that iterate through source rows and apply mappings
- Row-by-row processing with optional empty row skipping
- Filter application during iteration
- Negative value conversion if enabled

**Filter operators**:
- Text: `empty`, `not_empty`, `equals`, `not_equals`, `contains`, `not_contains`, `starts_with`, `ends_with`
- Numeric: `equals_zero`, `not_zero`, `greater_than`, `less_than`
- List: `all_in_list` (excludes rows where all comma-separated values are in the exclusion list)

**File handling**:
- Uses `openpyxl` with `data_only=False` for source files to preserve formulas
- Separate handling for `.xlsx` vs `.xlsm` files to preserve VBA/macro content
- Output files are generated with `_atualizada` suffix

### Frontend Structure

The web interface uses vanilla JavaScript with:

- **State management**: Simple object tracking source/destination files, sheets, and columns
- **DOM manipulation**: Direct DOM updates for file uploads, dropdowns, and mapping rows
- **Template-based UI**: Uses HTML templates for dynamic mapping and filter rows
- **Custom styling**: CSS variables for theming with a modern, clean design

## Important Implementation Details

- **Language**: The UI and code comments are in Portuguese (pt-BR)
- **Number format handling**: Supports both `.` and `,` as decimal separators for numeric values
- **Empty value detection**: Uses both `None` check and string whitespace stripping
- **Row counting**: Determines last filled row by scanning from bottom up per source column
- **Combined mappings**: Limits to 2 source columns per target column
- **File security**: Output files are always new files, never overwrite source/destination files

## Excel File Support

- **Formats**: `.xlsx` and `.xlsm` (with macro/VBA preservation)
- **Sheet detection**: Automatically lists available sheets and default to first sheet
- **Column scanning**: Identifies columns with data by scanning entire sheet, showing headers and sample values
- **Row/column limits**: Respects `max_row` and `max_column` from openpyxl workbook metadata