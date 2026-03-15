# PDF Tool — Watermark Removal + OCR to Word

A local, privacy-first command-line tool that removes watermarks from PDF files and/or converts them to editable Word documents using a locally-running AI model. No cloud services, no API keys — everything runs on your own machine.

---

## Features

- **Three processing modes** — pick what you need at runtime
- **Four-strategy watermark engine** — handles A-PDF QQAP markers, gray diagonal text, opacity-based form XObjects, and a pixel-level image fallback for stubborn cases
- **AI-powered OCR** via [glm-ocr](https://ollama.com/library/glm-ocr) running locally through [Ollama](https://ollama.com) — preserves headings, bold/italic, tables, lists, and inline code
- **Multithreaded OCR** — configurable thread count per PDF
- **Auto-pull model** — if `glm-ocr` isn't downloaded yet, the tool fetches it automatically before starting
- **Batch processing** — processes every `.pdf` in the current working directory in one run
- **Zero external binaries** — PDF rendering uses PyMuPDF (pure Python wheel, no Poppler needed)

---

## Modes

| # | Name | What it does | Output |
|---|------|-------------|--------|
| 1 | **Remove watermark only** | Strips watermarks from each PDF | `cleaned/<name>_cleaned.pdf` |
| 2 | **Remove watermark + OCR to Word** | Strips watermarks, then OCRs the clean PDF | `cleaned/<name>_cleaned.pdf` + `<name>.docx` |
| 3 | **OCR to Word only** | OCRs source PDFs directly, nothing is removed | `<name>.docx` |

The `.docx` files are always saved next to the original PDFs. The `cleaned/` folder is created automatically in the working directory.

---

## Requirements

### System requirements

| Requirement | Notes |
|-------------|-------|
| Python 3.10+ | 3.11 or 3.12 recommended |
| [Ollama](https://ollama.com/download) | Must be installed and running (`ollama serve`) |
| `glm-ocr` model | Auto-pulled on first run if missing |

> **Windows users:** No additional system binaries are needed. PyMuPDF ships as a self-contained wheel.

### Python packages

See [`requirements.txt`](requirements.txt). Install with:

```bash
pip install -r requirements.txt
```

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/your-username/pdf-tool.git
cd pdf-tool

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. Install Ollama (if not already installed)
#    https://ollama.com/download
#    Then start the Ollama server:
ollama serve
```

The `glm-ocr` model (~2.2 GB) will be pulled automatically the first time you run in a mode that needs OCR. You can also pull it manually:

```bash
ollama pull glm-ocr
```

---

## Usage

Place the script in the same folder as your PDF files, or `cd` into that folder, then run:

```bash
python pdf_tool.py
```

You will be greeted with an interactive menu:

```
╔══════════════════════════════════════════════════════════╗
║              PDF Tool  —  Watermark + OCR                ║
╠══════════════════════════════════════════════════════════╣
║  1  Remove watermark only          (PDF → clean PDF)     ║
║  2  Remove watermark + OCR to Word (PDF → clean PDF      ║
║                                         → .docx)         ║
║  3  OCR to Word only               (PDF → .docx)         ║
╚══════════════════════════════════════════════════════════╝
Enter mode [1 / 2 / 3]:
```

For modes 2 and 3 you will also be prompted for optional OCR settings:

```
OCR settings (press Enter to keep defaults):
  DPI (100-400) [200]:
  Threads (1-16) [4]:
```

Higher DPI improves OCR accuracy at the cost of speed and memory. More threads speed up multi-page documents — keep this at or below your physical CPU core count.

---

## Watermark Removal — How It Works

The engine tries four strategies in order, stopping as soon as one succeeds:

### Strategy 1 — QQAP marker
Targets the proprietary binary marker embedded by **A-PDF Watermark** software. Matching content streams and their immediate neighbors (bare `q`/`Q` wrappers) are zeroed out.

### Strategy 2 — Gray diagonal text blocks
Scans every `q…Q` content block in each page stream and removes blocks that match *all* of:
- Contain a `BT…ET` text-drawing section
- Use only a mid-gray color (`R ≈ G ≈ B`, value between 0.35 and 0.90), **or** carry a low fill-opacity (`ca < 0.7`)

Mixed streams (watermark blocks interleaved with real content) are handled correctly — only the watermark blocks are excised.

### Strategy 3 — Opacity-based Form XObjects
Finds Form XObjects that declare a `/Transparency` group or carry explicit `/ca` / `/CA` opacity keys and zeros their streams.

### Strategy 4 — Image fallback (rasterise + pixel erase)
If strategies 1–3 detect nothing, each page is rasterised at 2× resolution, and a pixel-level mask removes semi-transparent gray regions (the classic diagonal-text pattern). The PDF is rebuilt from the cleaned images. This is the slowest strategy but handles cases where the watermark is baked directly into the page raster.

---

## OCR — How It Works

1. Each page is rendered to a PNG using **PyMuPDF** at the configured DPI.
2. Pages are dispatched concurrently to **Ollama `glm-ocr`** — a 0.9B-parameter multimodal model optimised for document understanding.
3. The model returns structured **Markdown** preserving headings, bold/italic, tables, lists, code blocks, and special elements like figures and stamps.
4. The Markdown is converted to a properly formatted **`.docx`** file using `python-docx`, with page breaks between pages.

---

## Project Structure

```
pdf-tool/
├── pdf_tool.py               # Main script (all-in-one)
├── requirements.txt
└── README.md
```

---

## Troubleshooting

**`Cannot connect to Ollama`**
Make sure the Ollama server is running in a separate terminal:
```bash
ollama serve
```

**`"ollama" command not found`**
Ollama is not installed or not in your PATH. Download it from [ollama.com/download](https://ollama.com/download).

**OCR quality is poor**
Try increasing the DPI (e.g. `300` or `400`). For very small or blurry text this makes a significant difference.

**Watermark not fully removed**
The image fallback (Strategy 4) handles most remaining cases. If the watermark is still visible, the colors or opacity values may fall outside the current detection thresholds — open an issue with a sample.

**`ModuleNotFoundError: No module named 'cv2'`**
```bash
pip install opencv-python
```

---

## License

MIT — see `LICENSE` for details.
