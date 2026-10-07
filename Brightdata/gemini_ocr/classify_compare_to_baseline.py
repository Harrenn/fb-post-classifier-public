#!/usr/bin/env python3
"""Classify the Gemini-OCR 25-post CSV into compare_to_baseline.xlsx.

This is intentionally a second process/action after image OCR. It delegates the
actual classification, normalization, relationship checks, and workbook writing
to the repository's unchanged ``classify_posts.py`` implementation.
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[1]
DEFAULT_INPUT = REPO_ROOT / "input" / "compare_to_baseline.csv"
DEFAULT_OUTPUT = REPO_ROOT / "output" / "compare_to_baseline.xlsx"
DEFAULT_MODEL = "gemini-3.5-flash-lite"
DEFAULT_MIN_REQUEST_INTERVAL_SECONDS = 6.5


class RunnerError(RuntimeError):
    """Expected configuration or output failure."""


def load_classifier():
    path = REPO_ROOT / "classify_posts.py"
    spec = importlib.util.spec_from_file_location("fb_post_classifier_runtime", path)
    if spec is None or spec.loader is None:
        raise RunnerError("Could not load classify_posts.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--model", default=DEFAULT_MODEL)
    parser.add_argument(
        "--min-request-interval",
        type=float,
        default=DEFAULT_MIN_REQUEST_INTERVAL_SECONDS,
    )
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def run_self_test() -> None:
    checks = 0

    def require(condition: bool, message: str) -> None:
        nonlocal checks
        checks += 1
        if not condition:
            raise AssertionError(message)

    require(DEFAULT_INPUT.name == "compare_to_baseline.csv", "Input name failed")
    require(DEFAULT_OUTPUT.name == "compare_to_baseline.xlsx", "Output name failed")
    require(DEFAULT_MIN_REQUEST_INTERVAL_SECONDS >= 6.0, "Pacing is too aggressive")
    classifier = load_classifier()
    require(classifier.DEFAULT_MODEL == DEFAULT_MODEL, "Classifier model default drifted")
    print(f"Self-test passed: {checks} deterministic checks.")


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.self_test:
        run_self_test()
        return 0
    if args.min_request_interval < 0:
        print("Error: --min-request-interval must be non-negative.", file=sys.stderr)
        return 2

    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        print(
            "GEMINI_API_KEY is not set in the current process environment.",
            file=sys.stderr,
        )
        return 2

    classifier = load_classifier()
    input_path = args.input.resolve()
    output_path = args.output.resolve()
    generated_path = classifier.OUTPUT_DIR / f"{input_path.stem}{classifier.OUTPUT_SUFFIX}"
    if not input_path.exists():
        print(f"Input CSV does not exist: {input_path}", file=sys.stderr)
        return 2
    if output_path.exists() and not args.overwrite:
        print(
            f"Output already exists: {output_path}. Use --overwrite to replace it.",
            file=sys.stderr,
        )
        return 2

    model = args.model.strip() or DEFAULT_MODEL
    try:
        cost_rates = classifier.resolve_cost_rates(model)
        review_commercial = classifier.environment_flag(
            classifier.REVIEW_COMMERCIAL_ENV, True
        )
        pacer = classifier.RequestPacer(args.min_request_interval)
        print(f"Classification model: {model}")
        print(
            "Commercial completeness review: "
            + ("enabled" if review_commercial else "disabled")
        )
        print(f"Minimum request interval: {args.min_request_interval:g}s")
        with classifier.genai.Client(api_key=api_key) as client:
            (
                actual_generated_path,
                post_count,
                offer_count,
                price_point_count,
                addon_count,
                row_errors,
            ) = classifier.process_file(
                input_path,
                client,
                model,
                api_key,
                cost_rates,
                review_commercial,
                pacer,
            )
        if actual_generated_path.resolve() != generated_path.resolve():
            raise RunnerError("Classifier returned an unexpected generated path")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        os.replace(generated_path, output_path)
        print(
            f"Saved {output_path}: {post_count} posts, {offer_count} offers, "
            f"{price_point_count} price points, {addon_count} add-ons, "
            f"{row_errors} row errors."
        )
        return 0
    except classifier.FatalGeminiError as exc:
        print(
            "Gemini classification authentication, access, or quota failed: "
            + classifier.sanitize_error(exc, api_key),
            file=sys.stderr,
        )
        return 3
    except Exception as exc:
        print(
            "Classification failed: " + classifier.sanitize_error(exc, api_key),
            file=sys.stderr,
        )
        return 1
    finally:
        api_key = ""


if __name__ == "__main__":
    raise SystemExit(main())
