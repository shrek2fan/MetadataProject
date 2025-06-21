# MetadataProject

This repository contains utilities for cleaning and verifying large Excel metadata spreadsheets. The tools were developed as part of a school project and demonstrate complex data validation workflows.

## Project Structure

- `src/` – Python source code for the validation utilities and GUI.
- `data/` – Sample datasets and spreadsheets used by the scripts.
- `dist/` – Compiled binaries and generated example output.
- `assets/` – Additional resources such as screenshots.

## Requirements

Install dependencies using `pip`:

```bash
pip install -r requirements.txt
```

## Usage

The main validation logic is in `FinalBigSpread.py` and can be executed on an Excel file from the command line:

```bash
python -m src.FinalBigSpread --file your_spreadsheet.xlsx
```

A simple GUI front end is provided in `src/UI.py`:

```bash
python -m src.UI
```

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
