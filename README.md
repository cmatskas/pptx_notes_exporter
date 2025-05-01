# pptx_notes_exporter

A Python tool to export notes from PowerPoint presentations into a text file. Each slide's notes are formatted with `<SLIDE>` tags in the output file.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

Run the script and enter the path to your PowerPoint file when prompted:

```python
python main.py
```

Or use the function in your own code:

```python
from main import export_powerpoint_notes

# Export to default location (Desktop/Slide_Notes.txt)
export_powerpoint_notes("presentation.pptx")

# Export to custom location
export_powerpoint_notes("presentation.pptx", "output.txt")
```

## Features

- Extracts notes from PowerPoint presentations
- Skips hidden slides automatically
- Formats output with `<SLIDE>` tags
- Default output to Desktop as "Slide_Notes.txt"
- Support for custom output file location
- Handles slides with no notes (marks as "No notes")
- UTF-8 encoding support