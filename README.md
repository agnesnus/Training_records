# Training Records PDF to Excel (Streamlit)

This project provides a Streamlit app that reads uploaded training-record PDFs and generates an Excel file in the required format.

## What the app produces

The generated Excel file contains:

- Main header in row 1: `trainer records`
- Column headers in row 2:
	- `Name`
	- `Employee ID`
	- `Title`
	- `Revision`

The app extracts:

- `Name` from trainee name fields.
- `Employee ID` from Trainee Employee ID fields.
- `Title` from document number/revision plus description/topic, joined by a comma.
- `Revision` as normalized version values like `01`, `02`, etc.

If a trainee record contains multiple documents, each document becomes a separate row with the same trainee name and employee ID.

## Parsing logic included

The app handles common patterns where documents are represented as:

- Table columns such as `Document Number & Revision` and `Description/Topic`
- Multiple entries split by line breaks inside cells
- Text blocks where document number, revision, and title appear continuously
- Fallback matching when document/revision appears on one line and title on the next

Note: if the PDF is scanned or has unusual layout/font rendering (for example underline cues only), extraction may need manual review.

## Files

- `app.py`: Streamlit application and PDF-to-Excel logic
- `requirements.txt`: Python dependencies

## Setup

1. Create and activate a virtual environment (recommended):

```bash
python -m venv .venv
source .venv/bin/activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the app:

```bash
streamlit run app.py
```

4. In the web UI:
- Upload one or more PDF training records.
- Review extracted rows in the preview table.
- Click `Download Excel`.

## Output example

| Name | Employee ID | Title | Revision |
|---|---|---|---|
| John Doe | T12345 | SOP-001, Lockout and Tagout | 01 |
| John Doe | T12345 | WI-009, Chemical Handling | 02 |

## Improvement ideas

- Add OCR support (for image-only/scanned PDFs).
- Add per-template configuration for exact field anchors.
- Add confidence score and a validation screen before export.
