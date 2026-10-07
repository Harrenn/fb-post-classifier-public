# Local OCR preprocessing

`process_brightdata.py` reads a local Bright Data JSON export, selects image attachments and produces OCR-enriched JSON and classifier-compatible CSV. Source files are not modified. No datasets or benchmark results are bundled.

Use a separate environment for the local OCR dependencies:

```bash
python -m venv Brightdata/rapidocr/.venv
Brightdata/rapidocr/.venv/bin/python -m pip install -r Brightdata/rapidocr/requirements.txt
Brightdata/rapidocr/.venv/bin/python Brightdata/rapidocr/process_brightdata.py --help
Brightdata/rapidocr/.venv/bin/python Brightdata/rapidocr/process_brightdata.py --self-test
```

On Windows use the environment's Scripts/python.exe. The default input is input/bright_data_dataset.json; use CLI options to choose your own input. Outputs and caches stay local and ignored. A real run can download source images and OCR model weights; the self-test does not run a live scrape or hosted API request.
