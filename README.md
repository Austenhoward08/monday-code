# Monday Exporter

Export Monday.com boards into polished Excel workbooks from the command line.

This tool uses the Monday.com GraphQL API to download board metadata, items, groups, and column
values, then renders the data into an `.xlsx` file with sensible formatting (styled headers,
freeze panes, alternating row colours, automatic column widths, date handling, and a summary tab).

## Prerequisites

- Python **3.10+**
- A Monday.com API v2 token with access to the boards you want to export

## Installation

```bash
git clone <repo-url>
cd monday-code
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -e .
```

## Usage

```bash
export MONDAY_API_TOKEN="your-api-token"
monday-export --board-id 123456789 --output path/to/board.xlsx
```

### Options

- `--board-id / -b` (required): Numeric Monday.com board ID to export.
- `--output / -o`: Path to the Excel file to create. Defaults to `<board-name>.xlsx`.
- `--api-token`: Provide the Monday.com API token inline instead of using the environment variable.
- `--include-subitems`: Fetch and include subitems (if your board uses them).
- `--page-size`: Override the pagination size when pulling items (default 500, max 1000).
- `--verbose / -v`: Enable informational logging to help debug API issues.

### Excel Output

- `Items` sheet with:
  - Board rows (ID, name, group, creator, timestamps, and every visible column)
  - Styled header row (bold, coloured, frozen)
  - Alternating row fill and thin borders
  - Auto-sized columns with sensible caps
  - Date columns rendered with `yyyy-mm-dd hh:mm` format
- `Summary` sheet with board metadata (name, ID, counts, description, export timestamp).

## Development

Install optional tooling:

```bash
pip install -e ".[dev]"
```

Run static checks:

```bash
ruff check .
mypy src
```

## Troubleshooting

- **401 / 403 errors**: Verify the API token and that it has access to the board.
- **Empty export**: Ensure you have items on the board and the user associated with the token can view them.
- **Large boards**: Increase `--page-size` (up to 1000) to reduce round trips. For extremely large boards,
  consider splitting the export per group to keep workbook sizes manageable.
