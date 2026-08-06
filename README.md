# PDF-OCR — Voter Registration PDF Data Extractor

A Flask web app that extracts structured data (name, serial, EPIC number, address, age, gender, etc.) from voter registration PDFs using Tesseract OCR and exports results to Excel.

---

## Features

- Upload a single PDF or an entire folder of PDFs
- Extracts voter card data: Serial, EPIC, Name, Relation, Father/Husband Name, House Number, Age, Gender, Ward No, Part No, Area
- Supports Hindi + English OCR (`eng+hin`)
- Exports results as `.xlsx` per PDF
- Live progress updates via Server-Sent Events (SSE)
- Optional area/locality extraction

---

## Requirements

### System Dependencies

| Dependency | Purpose |
|---|---|
| Python 3.9+ | Runtime |
| Tesseract OCR | OCR engine |
| Tesseract Hindi language pack | Hindi text recognition |

#### Install Tesseract on macOS

```bash
brew install tesseract
brew install tesseract-lang   # installs all language packs including Hindi
```

#### Install Tesseract on Ubuntu/Debian

```bash
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-hin
```

#### Install Tesseract on Windows

Download and install from: https://github.com/UB-Mannheim/tesseract/wiki  
During install, select **Hindi** under additional language data.  
Add Tesseract to your system PATH (e.g. `C:\Program Files\Tesseract-OCR`).

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/mr-1109/PDF-OCR.git
cd PDF-OCR
```

### 2. Create a virtual environment

```bash
python -m venv venv
```

Activate it:

- **macOS / Linux:**
  ```bash
  source venv/bin/activate
  ```
- **Windows:**
  ```bash
  venv\Scripts\activate
  ```

### 3. Install Python dependencies

```bash
pip install -r voter_ocr_web/requirements.txt
```

---

## Running the App

```bash
cd voter_ocr_web
python app.py
```

The app starts on **http://127.0.0.1:8080**

Open your browser and go to: `http://127.0.0.1:8080`

---

## How to Use

### Single PDF

1. Click **"Choose File"** and select a PDF
2. (Optional) Set page range, DPI, and whether to include area extraction
3. Click **"Run OCR"**
4. Wait for progress — results appear on screen
5. Click **"Download Excel"** to save the `.xlsx`

### Folder of PDFs

1. Click **"Upload Folder"** and select a folder containing PDFs
2. Set options (DPI, include area, etc.)
3. Click **"Run Folder OCR"**
4. Each PDF is processed one by one with live progress
5. Excel files are saved automatically to a `processed excel` folder next to your source folder

---

## Output

| Scenario | Output Location |
|---|---|
| Folder upload (local path detected) | `<parent of your folder>/processed excel/` |
| Browser file upload (no local path) | `voter_ocr_web/excel_outputs/<batch_id>/` |

Each output Excel has columns:
`WARD_NO, PART_NO, Page, List Type, Action, Serial, EPIC, Name, Relation, F_NAME, House Number, Age, Gender`

---

## Configuration

Key settings at the top of `app.py`:

| Setting | Default | Description |
|---|---|---|
| `dpi` | `200` | Render resolution — higher = slower but more accurate |
| `HEADER_HEIGHT` | `114` | PDF layout constant (pixels) |
| `CARD_HEIGHT` | `330` | Voter card height (pixels) |
| `MAX_CELLS` | `30` | Max voter cards per page |

---

## Troubleshooting

**`tesseract is not installed or it's not in your PATH`**  
→ Install Tesseract and ensure it's in your system PATH.

**Hindi text garbled / missing**  
→ Ensure the Hindi language pack (`hin.traineddata`) is installed. Check with:
```bash
tesseract --list-langs
```

**App is slow**  
→ Reduce DPI (try 150). OCR is CPU-intensive — lower DPI speeds up processing significantly.

**Port already in use**  
→ Kill existing process:
```bash
pkill -f "python.*app.py"
```
Then restart.

---

## Project Structure

```
PDF-OCR/
├── voter_ocr_web/
│   ├── app.py              # Main Flask app + OCR logic
│   ├── requirements.txt    # Python dependencies
│   └── templates/
│       └── index.html      # Web UI
└── README.md
```
