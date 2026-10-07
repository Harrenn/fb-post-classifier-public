#!/usr/bin/env python3
"""Pair Apify and Bright Data media OCR, score agreement, and build audit sheets.

The comparison is media-ID based. Private source URLs, downloaded images, and
comparison outputs are written only below ``Brightdata/comparison/output/``,
which is ignored by Git.
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import io
import json
import re
import statistics
import sys
import unicodedata
from dataclasses import asdict, dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

from PIL import Image, ImageDraw, ImageFont, ImageOps


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[1]
DEFAULT_APIFY = REPO_ROOT / "input" / "apify_25_post.json"
DEFAULT_BRIGHT = REPO_ROOT / "input" / "brighdata_25_post.json"
DEFAULT_SMALL = (
    REPO_ROOT
    / "Brightdata"
    / "rapidocr"
    / "output"
    / "brighdata_25_post_rapidocr_ppocrv6_small.json"
)
DEFAULT_MEDIUM = (
    REPO_ROOT
    / "Brightdata"
    / "rapidocr"
    / "output"
    / "brighdata_25_post_rapidocr_ppocrv6_medium.json"
)
DEFAULT_GEMINI = (
    REPO_ROOT
    / "Brightdata"
    / "gemini_ocr"
    / "output"
    / "compare_to_baseline_ocr.json"
)
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "output"
TOKEN_RE = re.compile(r"[A-Z0-9]+(?:[.,][0-9]+)*")
NUMBER_RE = re.compile(r"\d+(?:[.,]\d+)*")
COMMERCIAL_TERMS = {
    "ADDON",
    "CASH",
    "DAILY",
    "DISCOUNT",
    "DOWNPAYMENT",
    "FREE",
    "HELMET",
    "INSTALLMENT",
    "LTO",
    "PISO",
    "PROMO",
    "REBATE",
    "REGISTRATION",
    "SUBSIDY",
    "ZERO",
}


class ComparisonError(RuntimeError):
    """Expected input, pairing, download, or output failure."""


@dataclass
class ComparisonRow:
    case_id: str
    post_index: int
    post_url: str
    media_id: str
    media_type: str
    apify_text: str
    rapidocr_small_text: str
    ppocrv6_medium_text: str
    gemini_text: str
    gemini_available: bool
    apify_chars: int
    rapidocr_small_chars: int
    ppocrv6_medium_chars: int
    gemini_chars: int
    apify_tokens: int
    rapidocr_small_tokens: int
    ppocrv6_medium_tokens: int
    gemini_tokens: int
    apify_numbers: str
    rapidocr_small_numbers: str
    ppocrv6_medium_numbers: str
    gemini_numbers: str
    apify_terms: str
    rapidocr_small_terms: str
    ppocrv6_medium_terms: str
    gemini_terms: str
    apify_small_similarity: float
    apify_medium_similarity: float
    small_medium_similarity: float
    apify_gemini_similarity: float
    small_gemini_similarity: float
    image_source_url: str
    image_path: str = ""


def load_json_array(path: Path) -> list[dict[str, Any]]:
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError as exc:
        raise ComparisonError(f"Input file does not exist: {path}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise ComparisonError(f"Could not read {path.name}: {type(exc).__name__}") from exc
    if not isinstance(value, list) or not all(isinstance(item, dict) for item in value):
        raise ComparisonError(f"{path.name} must contain a JSON array of objects")
    return value


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).upper()
    return " ".join(text.replace("\u00a0", " ").split())


def tokens(value: object) -> list[str]:
    return TOKEN_RE.findall(normalize_text(value))


def numeric_tokens(value: object) -> list[str]:
    normalized: list[str] = []
    for raw in NUMBER_RE.findall(normalize_text(value)):
        cleaned = raw.replace(",", "")
        normalized.append(cleaned)
    return sorted(set(normalized))


def commercial_terms(value: object) -> list[str]:
    token_set = set(tokens(value))
    return sorted(token_set & COMMERCIAL_TERMS)


def similarity(left: object, right: object) -> float:
    return round(SequenceMatcher(None, normalize_text(left), normalize_text(right)).ratio(), 4)


def index_by_url(rows: Iterable[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    indexed: dict[str, dict[str, Any]] = {}
    for row in rows:
        url = str(row.get("url") or "").strip()
        if not url:
            raise ComparisonError("Every post must have a non-empty URL")
        if url in indexed:
            raise ComparisonError(f"Duplicate post URL in paired input: {url}")
        indexed[url] = row
    return indexed


def media_by_id(row: dict[str, Any], field: str) -> dict[str, dict[str, Any]]:
    indexed: dict[str, dict[str, Any]] = {}
    values = row.get(field)
    if not isinstance(values, list):
        return indexed
    for item in values:
        if not isinstance(item, dict):
            continue
        media_id = str(item.get("id") or "").strip()
        if media_id:
            indexed[media_id] = item
    return indexed


def choose_bright_image_url(attachment: dict[str, Any]) -> str:
    media_type = str(attachment.get("type") or "").strip().casefold()
    fields = ("thumbnail_url", "url") if media_type == "video" else ("url", "thumbnail_url")
    for field in fields:
        value = str(attachment.get(field) or "").strip()
        if value and not (field == "url" and "facebook.com" in value.casefold()):
            return value
    return ""


def build_rows(
    apify: list[dict[str, Any]],
    bright: list[dict[str, Any]],
    small: list[dict[str, Any]],
    medium: list[dict[str, Any]],
    gemini: list[dict[str, Any]] | None = None,
) -> list[ComparisonRow]:
    bright_by_url = index_by_url(bright)
    small_by_url = index_by_url(small)
    medium_by_url = index_by_url(medium)
    gemini_by_url = index_by_url(gemini) if gemini is not None else None
    apify_urls = set(index_by_url(apify))
    if apify_urls != set(bright_by_url) or apify_urls != set(small_by_url) or apify_urls != set(medium_by_url):
        raise ComparisonError("The four inputs do not contain the same post URL set")
    if gemini_by_url is not None and apify_urls != set(gemini_by_url):
        raise ComparisonError("The Gemini OCR input does not contain the same post URL set")

    output: list[ComparisonRow] = []
    for post_index, apify_post in enumerate(apify, start=1):
        post_url = str(apify_post["url"]).strip()
        apify_media = media_by_id(apify_post, "media")
        bright_media = media_by_id(bright_by_url[post_url], "attachments")
        small_media = media_by_id(small_by_url[post_url], "attachments")
        medium_media = media_by_id(medium_by_url[post_url], "attachments")
        gemini_media = (
            media_by_id(gemini_by_url[post_url], "attachments")
            if gemini_by_url is not None
            else {}
        )
        shared_ids = set(apify_media) & set(bright_media) & set(small_media) & set(medium_media)
        if gemini_by_url is not None:
            shared_ids &= set(gemini_media)
        shared_ids = sorted(shared_ids)
        for media_index, media_id in enumerate(shared_ids, start=1):
            apify_item = apify_media[media_id]
            bright_item = bright_media[media_id]
            small_item = small_media[media_id]
            medium_item = medium_media[media_id]
            gemini_item = gemini_media.get(media_id, {})
            apify_text = str(apify_item.get("ocrText") or "")
            small_text = str(small_item.get("ocrText") or "")
            medium_text = str(medium_item.get("ocrText") or "")
            gemini_text = str(gemini_item.get("ocrText") or "")
            output.append(
                ComparisonRow(
                    case_id=f"P{post_index:02d}-M{media_index:02d}",
                    post_index=post_index,
                    post_url=post_url,
                    media_id=media_id,
                    media_type=str(bright_item.get("type") or apify_item.get("__typename") or ""),
                    apify_text=apify_text,
                    rapidocr_small_text=small_text,
                    ppocrv6_medium_text=medium_text,
                    gemini_text=gemini_text,
                    gemini_available=gemini_by_url is not None,
                    apify_chars=len(apify_text),
                    rapidocr_small_chars=len(small_text),
                    ppocrv6_medium_chars=len(medium_text),
                    gemini_chars=len(gemini_text),
                    apify_tokens=len(tokens(apify_text)),
                    rapidocr_small_tokens=len(tokens(small_text)),
                    ppocrv6_medium_tokens=len(tokens(medium_text)),
                    gemini_tokens=len(tokens(gemini_text)),
                    apify_numbers="; ".join(numeric_tokens(apify_text)),
                    rapidocr_small_numbers="; ".join(numeric_tokens(small_text)),
                    ppocrv6_medium_numbers="; ".join(numeric_tokens(medium_text)),
                    gemini_numbers="; ".join(numeric_tokens(gemini_text)),
                    apify_terms="; ".join(commercial_terms(apify_text)),
                    rapidocr_small_terms="; ".join(commercial_terms(small_text)),
                    ppocrv6_medium_terms="; ".join(commercial_terms(medium_text)),
                    gemini_terms="; ".join(commercial_terms(gemini_text)),
                    apify_small_similarity=similarity(apify_text, small_text),
                    apify_medium_similarity=similarity(apify_text, medium_text),
                    small_medium_similarity=similarity(small_text, medium_text),
                    apify_gemini_similarity=similarity(apify_text, gemini_text),
                    small_gemini_similarity=similarity(small_text, gemini_text),
                    image_source_url=choose_bright_image_url(bright_item),
                )
            )
    return output


def download_images(rows: list[ComparisonRow], output_dir: Path) -> None:
    processor_path = REPO_ROOT / "Brightdata" / "rapidocr" / "process_brightdata.py"
    spec = importlib.util.spec_from_file_location("brightdata_processor", processor_path)
    if spec is None or spec.loader is None:
        raise ComparisonError("Could not load the Bright Data downloader")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)

    image_dir = output_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    for index, row in enumerate(rows, start=1):
        if not row.image_source_url:
            continue
        path = image_dir / f"{index:03d}_{row.case_id}.jpg"
        if not path.exists():
            payload = module.download_image(
                row.image_source_url,
                timeout_seconds=30,
                max_bytes=25 * 1024 * 1024,
                attempts=3,
            )
            with Image.open(io.BytesIO(payload)) as source:
                image = ImageOps.exif_transpose(source).convert("RGB")
                image.save(path, "JPEG", quality=94, optimize=True)
        row.image_path = str(path.resolve())


def make_contact_sheets(rows: list[ComparisonRow], output_dir: Path) -> list[Path]:
    eligible = [row for row in rows if row.image_path]
    sheet_dir = output_dir / "contact_sheets"
    sheet_dir.mkdir(parents=True, exist_ok=True)
    font = ImageFont.truetype("arial.ttf", 24)
    paths: list[Path] = []
    cell_width, cell_height, label_height = 760, 760, 42
    margin = 24
    for sheet_index in range(0, len(eligible), 4):
        page_rows = eligible[sheet_index : sheet_index + 4]
        canvas = Image.new("RGB", (2 * cell_width + 3 * margin, 2 * (cell_height + label_height) + 3 * margin), "white")
        draw = ImageDraw.Draw(canvas)
        for position, row in enumerate(page_rows):
            column, line = position % 2, position // 2
            x = margin + column * (cell_width + margin)
            y = margin + line * (cell_height + label_height + margin)
            with Image.open(row.image_path) as source:
                image = ImageOps.contain(source.convert("RGB"), (cell_width, cell_height))
            image_x = x + (cell_width - image.width) // 2
            image_y = y + label_height + (cell_height - image.height) // 2
            canvas.paste(image, (image_x, image_y))
            draw.text((x, y), f"{row.case_id}  media {row.media_id}", fill="black", font=font)
        path = sheet_dir / f"sheet_{sheet_index // 4 + 1:02d}.jpg"
        canvas.save(path, "JPEG", quality=94, optimize=True)
        paths.append(path)
    return paths


def write_outputs(rows: list[ComparisonRow], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    dictionaries = [asdict(row) for row in rows]
    (output_dir / "ocr_comparison.json").write_text(
        json.dumps(dictionaries, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    with (output_dir / "ocr_comparison.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(dictionaries[0]) if dictionaries else [])
        if dictionaries:
            writer.writeheader()
            writer.writerows(dictionaries)


def quality_summary(rows: list[ComparisonRow], adjudication_path: Path) -> dict[str, Any]:
    comparable = [row for row in rows if row.apify_text.strip()]
    summary: dict[str, Any] = {
        "matched_media": len(rows),
        "apify_ocr_nonempty": len(comparable),
        "text_volume": {
            "apify_chars": sum(row.apify_chars for row in comparable),
            "rapidocr_small_chars": sum(row.rapidocr_small_chars for row in comparable),
            "ppocrv6_medium_chars": sum(row.ppocrv6_medium_chars for row in comparable),
            "gemini_chars": sum(row.gemini_chars for row in comparable),
            "apify_tokens": sum(row.apify_tokens for row in comparable),
            "rapidocr_small_tokens": sum(row.rapidocr_small_tokens for row in comparable),
            "ppocrv6_medium_tokens": sum(row.ppocrv6_medium_tokens for row in comparable),
            "gemini_tokens": sum(row.gemini_tokens for row in comparable),
        },
        "mean_pairwise_similarity": {
            "apify_vs_rapidocr_small": round(
                statistics.mean(row.apify_small_similarity for row in comparable), 4
            ),
            "apify_vs_ppocrv6_medium": round(
                statistics.mean(row.apify_medium_similarity for row in comparable), 4
            ),
            "rapidocr_small_vs_ppocrv6_medium": round(
                statistics.mean(row.small_medium_similarity for row in comparable), 4
            ),
            "apify_vs_gemini": round(
                statistics.mean(row.apify_gemini_similarity for row in comparable), 4
            ),
            "rapidocr_small_vs_gemini": round(
                statistics.mean(row.small_gemini_similarity for row in comparable), 4
            ),
        },
    }
    if not any(row.gemini_available for row in rows):
        summary["text_volume"].pop("gemini_chars", None)
        summary["text_volume"].pop("gemini_tokens", None)
        summary["mean_pairwise_similarity"].pop("apify_vs_gemini", None)
        summary["mean_pairwise_similarity"].pop("rapidocr_small_vs_gemini", None)
    if adjudication_path.exists():
        adjudication = json.loads(adjudication_path.read_text(encoding="utf-8"))
        cases = adjudication.get("cases")
        if not isinstance(cases, list):
            raise ComparisonError("Visual adjudication must contain a cases array")
        known_ids = {row.case_id for row in rows}
        if any(not isinstance(case, dict) or case.get("case_id") not in known_ids for case in cases):
            raise ComparisonError("Visual adjudication contains an unknown case ID")
        fact_total = sum(int(case["facts"]) for case in cases)
        scores: dict[str, dict[str, float | int]] = {}
        models = ["apify", "rapidocr_small", "ppocrv6_medium"]
        if all("gemini" in case for case in cases):
            models.append("gemini")
        for model in models:
            correct = sum(int(case[model]) for case in cases)
            scores[model] = {
                "correct_facts": correct,
                "total_facts": fact_total,
                "fact_recall_percent": round(correct / fact_total * 100, 1),
            }
        summary["visual_fact_adjudication"] = {
            "reviewed_cases": len(cases),
            "scores": scores,
            "rubric": adjudication.get("rubric", {}),
        }
        numeric_scores = adjudication.get("numeric_fact_scores")
        if isinstance(numeric_scores, dict):
            total_numeric = int(numeric_scores["total_facts"])
            summary["numeric_fact_adjudication"] = {
                "scope": numeric_scores.get("scope", ""),
                "note": numeric_scores.get("note", ""),
                "scores": {
                    model: {
                        "correct_facts": int(numeric_scores[model]),
                        "total_facts": total_numeric,
                        "fact_recall_percent": round(
                            int(numeric_scores[model]) / total_numeric * 100, 1
                        ),
                    }
                    for model in models
                    if model in numeric_scores
                },
            }
    return summary


def run_self_test() -> None:
    assert normalize_text("  Piso\u00a0 deal ") == "PISO DEAL"
    assert tokens("P54,000 ZERO") == ["P54,000", "ZERO"]
    assert numeric_tokens("P54,000 and 200 daily") == ["200", "54000"]
    assert commercial_terms("Free helmet, zero downpayment") == [
        "DOWNPAYMENT",
        "FREE",
        "HELMET",
        "ZERO",
    ]
    assert similarity("PISO DEAL", "piso  deal") == 1.0
    print("Self-test passed: 5 deterministic checks.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--apify", type=Path, default=DEFAULT_APIFY)
    parser.add_argument("--bright", type=Path, default=DEFAULT_BRIGHT)
    parser.add_argument("--small", type=Path, default=DEFAULT_SMALL)
    parser.add_argument("--medium", type=Path, default=DEFAULT_MEDIUM)
    parser.add_argument("--gemini", type=Path, default=DEFAULT_GEMINI)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--download-images", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.self_test:
        run_self_test()
        return 0
    try:
        gemini_rows = (
            load_json_array(args.gemini.resolve()) if args.gemini.exists() else None
        )
        rows = build_rows(
            load_json_array(args.apify.resolve()),
            load_json_array(args.bright.resolve()),
            load_json_array(args.small.resolve()),
            load_json_array(args.medium.resolve()),
            gemini_rows,
        )
        if args.download_images:
            download_images(rows, args.output_dir.resolve())
            sheets = make_contact_sheets(rows, args.output_dir.resolve())
        else:
            sheets = []
        write_outputs(rows, args.output_dir.resolve())
        summary = quality_summary(
            rows, args.output_dir.resolve() / "visual_adjudication.json"
        )
        (args.output_dir.resolve() / "quality_summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        nonempty_apify = sum(bool(row.apify_text.strip()) for row in rows)
        print(
            f"Matched media={len(rows)}, Apify OCR non-empty={nonempty_apify}, "
            f"contact_sheets={len(sheets)}"
        )
        if "visual_fact_adjudication" in summary:
            scores = summary["visual_fact_adjudication"]["scores"]
            print(
                "Visual fact recall: "
                f"Apify={scores['apify']['fact_recall_percent']}%, "
                f"RapidOCR Small={scores['rapidocr_small']['fact_recall_percent']}%, "
                f"PP-OCRv6 Medium={scores['ppocrv6_medium']['fact_recall_percent']}%"
                + (
                    f", Gemini={scores['gemini']['fact_recall_percent']}%"
                    if "gemini" in scores
                    else ""
                )
            )
        print(f"Comparison output: {args.output_dir.resolve()}")
        return 0
    except (ComparisonError, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
