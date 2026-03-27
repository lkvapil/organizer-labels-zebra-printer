# Zebra Organiser

A PyQt6 GUI application for printing labels on a Zebra printer from an Excel file.

## Description

The application loads data from `organiser.xlsx` and sends labels to a Zebra printer via the ZPL protocol. Each row in the Excel file corresponds to one printed label.

## Requirements

- Python 3.8+
- PyQt6
- openpyxl
- zebra-day

Install dependencies:

```bash
pip install PyQt6 openpyxl zebra
```

## Usage

```bash
python3 organiser.py
```

## organiser.xlsx

The `organiser.xlsx` file contains the data to print. Each row equals one label — each cell in the row is printed as a separate line of text on the label.

## Features

- Automatic printer discovery (lpstat / Zebra library)
- Label size selection (50x25mm, 4x6", 4x3", ...)
- DPI setting (203 / 300)
- Optional border rectangle around text
- Automatic diacritic replacement for ZPL compatibility
- Remembers the last used printer

## Third-Party Licenses

This project uses the following open-source libraries:

| Library | License | Notes |
|---------|---------|-------|
| [PyQt6](https://www.riverbankcomputing.com/software/pyqt/) | GPL v3 / Commercial | Free for open-source use under GPL v3; commercial use requires a license from Riverbank Computing |
| [openpyxl](https://openpyxl.readthedocs.io/) | MIT | Free to use in any project |
| [zebra](https://pypi.org/project/zebra/) | MIT | Free to use in any project |

This project itself is released under the MIT License — see [LICENSE](LICENSE) for details.
