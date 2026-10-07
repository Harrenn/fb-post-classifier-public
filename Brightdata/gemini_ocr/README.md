# Gemini image OCR and classification

This optional pipeline transcribes images with Gemini, then invokes the standalone text classifier as a separate stage. It requires your own local Bright Data export and GEMINI_API_KEY in the process environment. No private inputs, manual baselines or experiment results are bundled.

Install the root requirements and this directory's requirements in a compatible environment. Inspect options before running:

```bash
python Brightdata/gemini_ocr/process_brightdata.py --help
python Brightdata/gemini_ocr/classify_compare_to_baseline.py --help
python Brightdata/gemini_ocr/process_brightdata.py --self-test
python Brightdata/gemini_ocr/classify_compare_to_baseline.py --self-test
```

Configure input/output paths and model settings through available CLI/environment options. Cached OCR and generated CSV/workbook files stay local and ignored. The PowerShell run_experiment.ps1 launcher prompts for a key with hidden input and clears its process variable afterward. Live runs send images/text to the provider and can incur charges; self-tests make no API calls. Do not publish real exports, cached transcripts, credentials or downloaded media.
