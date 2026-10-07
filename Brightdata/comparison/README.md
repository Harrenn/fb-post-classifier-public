# Paired OCR comparison

`compare_ocr.py` pairs local Apify and Bright Data exports by post URL and media ID, compares OCR readings, and can generate local audit sheets. No private exports or prior results are bundled.

Install Pillow in an environment with the classifier dependencies. Run:

```bash
python Brightdata/comparison/compare_ocr.py --help
python Brightdata/comparison/compare_ocr.py --self-test
```

Set input paths through the CLI for your own local exports and transcripts. Generated media and reports remain under the ignored output directory. Keep raw post URLs, media and comparison evidence private.
