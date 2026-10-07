# Contributor context

This public source-only distribution contains the standalone classifier and optional OCR preprocessing tools. Internal management plans, costing presentations, experiment evidence and private datasets are excluded, and history starts fresh.

Read classify_posts.py, data_dictionary.md and guide.md before changing classification logic. Preserve the four-sheet contract, source grounding, unknown-value handling, model-price relationships, reversal of unsupported claims and token/error accounting. Do not refactor the single-file classifier without an explicit request.

Use only synthetic fixtures in tests. Never commit API keys, browser state, input datasets, OCR caches, downloaded media, generated outputs or notebook results. Credentials come from the process environment; preserve error redaction.

Run `python classify_posts.py --self-test` before and after classifier changes. Optional tools each expose `--self-test`. These checks make no API calls. Do not run paid API jobs or download private datasets for routine validation.

Update relevant documentation when behavior changes. Model availability, rates and package versions are configurable and can change; verify current provider documentation before operational use.
