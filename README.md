# Facebook Post Classifier

Classifies scraped Facebook motorcycle posts with Google Gemini and writes an
analytics-ready Excel workbook for competitor-content analysis.

## Project structure

```text
fb_post_classifier/
├── classify_posts.py
├── input/
└── output/
```

Place one or more Facebook scraper CSV files in `input/`. The script reads the
post caption from `text`, detects every OCR column case-insensitively, and
combines the available text before classification. Source files are never
overwritten.

Each input CSV produces an `.xlsx` workbook in `output/` with separate `Posts`,
`Offers`, `PricePoints`, and `Addons` sheets. See `data_dictionary.md` for the
complete output schema and classification logic.

## Documentation

- `guide.md`: practical instructions for reading the workbook, with scenario-based
  “what if?” examples and expected outputs.
- `data_dictionary.md`: formal definitions, types, allowed values, and logic for
  every output column.
- `original.md`: generic classification requirements.
- `codex_ver.md`: the implemented analytics design and classifier behavior.

## Setup

Python 3.10 or newer is recommended.

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
export GEMINI_API_KEY="your-google-api-key"
python classify_posts.py
```

The default model is `gemini-3.5-flash-lite`. You can override it with
`GEMINI_MODEL`.

Run the built-in deterministic checks without making API calls:

```bash
python classify_posts.py --self-test
```

Scraped CSV files, generated workbooks, environment files, and API keys are
excluded from Git.

## Optional OCR tools

- `Brightdata/rapidocr/`: local image OCR preprocessing.
- `Brightdata/gemini_ocr/`: separate Gemini image transcription and classification stages.
- `Brightdata/comparison/`: paired OCR comparison using your own local exports.

See each directory's README for setup and offline checks. These tools require local input files, which are not bundled. Hosted OCR and classification send source text or images to the configured provider and can incur charges.

## Public distribution

This repository starts with fresh history. Private source datasets, generated workbooks, cached media, executed evaluation notebooks, internal planning documents and presentation assets are excluded. Keep credentials in the process environment and inspect files before committing. Model defaults and embedded pricing are historical configuration; verify availability and rates before a live run.
