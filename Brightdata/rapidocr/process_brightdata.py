#!/usr/bin/env python3
"""Download Bright Data post images, OCR them locally, and enrich the export.

The source JSON is never modified. Each run writes model-specific artifacts:

* ``Brightdata/rapidocr/output/<input>_rapidocr_ppocrv6_<tier>.json`` with
  per-attachment ``ocrText`` and ``ocrStatus`` fields plus a post-level
  ``ocr_texts`` list.
* ``Brightdata/rapidocr/output/<input>_rapidocr_ppocrv6_<tier>.csv`` in the input
  shape accepted by ``classify_posts.py``.

OCR uses RapidOCR with PP-OCRv6 and ONNX Runtime on the local machine. The
default ``small`` tier is RapidOCR's standard profile; ``medium`` selects the
higher-accuracy PP-OCRv6 Medium detector and recognizer. No hosted OCR or LLM
API is called.
"""

from __future__ import annotations

import argparse
import copy
import csv
import hashlib
import io
import ipaddress
import json
import os
import socket
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any, Callable, Iterable


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parents[1]
DEFAULT_INPUT = REPO_ROOT / "input" / "bright_data_dataset.json"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) BrightDataLocalOCR/1.0"
MODEL_TYPES = ("tiny", "small", "medium")
CSV_BASE_COLUMNS = [
    "url",
    "text",
    "likes",
    "comments",
    "shares",
    "date_posted",
    "post_id",
    "post_type",
    "page_name",
    "page_url",
    "page_followers",
    "page_category",
    "is_sponsored",
]


class ProcessorError(RuntimeError):
    """Expected input, download, image, or configuration failure."""


@dataclass(frozen=True)
class ImageSource:
    field: str
    url: str


@dataclass(frozen=True)
class OCRResult:
    text: str
    status: str
    error: str = ""


def nonempty(value: object) -> bool:
    return value is not None and bool(str(value).strip())


def clean_ocr_lines(lines: Iterable[object]) -> str:
    cleaned: list[str] = []
    for value in lines:
        line = " ".join(str(value).replace("\u00a0", " ").split()).strip()
        if line and (not cleaned or line != cleaned[-1]):
            cleaned.append(line)
    return "\n".join(cleaned)


def package_version(name: str) -> str:
    try:
        return version(name)
    except PackageNotFoundError:
        return "unknown"


def validate_remote_url(raw_url: object) -> str:
    if not nonempty(raw_url):
        raise ProcessorError("image URL is empty")
    url = str(raw_url).strip()
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ProcessorError("image URL must use HTTP or HTTPS")

    try:
        literal_ip = ipaddress.ip_address(parsed.hostname.strip("[]"))
    except ValueError:
        literal_ip = None
    if literal_ip and (
        literal_ip.is_private
        or literal_ip.is_loopback
        or literal_ip.is_link_local
        or literal_ip.is_multicast
        or literal_ip.is_reserved
    ):
        raise ProcessorError("local or reserved image hosts are not allowed")
    return url


def choose_attachment_image(attachment: dict[str, Any]) -> ImageSource | None:
    media_type = str(attachment.get("type") or "").strip().casefold()
    fields = ("thumbnail_url", "url") if media_type == "video" else ("url", "thumbnail_url")
    for field in fields:
        value = attachment.get(field)
        if not nonempty(value):
            continue
        try:
            url = validate_remote_url(value)
        except ProcessorError:
            continue
        hostname = (urllib.parse.urlparse(url).hostname or "").casefold()
        if hostname.endswith("facebook.com") and field == "url":
            continue
        return ImageSource(field=field, url=url)
    return None


def safe_error(exc: BaseException) -> str:
    if isinstance(exc, urllib.error.HTTPError):
        return f"HTTP {exc.code}"
    if isinstance(exc, urllib.error.URLError):
        return f"network error: {type(exc.reason).__name__}"
    if isinstance(exc, (TimeoutError, socket.timeout)):
        return "network timeout"
    message = " ".join(str(exc).split())
    for prefix in ("http://", "https://"):
        if prefix in message.casefold():
            return type(exc).__name__
    return message[:180] or type(exc).__name__


def download_image(
    url: str,
    *,
    timeout_seconds: float,
    max_bytes: int,
    attempts: int,
) -> bytes:
    validated_url = validate_remote_url(url)
    last_error: BaseException | None = None
    for attempt in range(1, attempts + 1):
        request = urllib.request.Request(
            validated_url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
                content_type = response.headers.get_content_type().casefold()
                content_length = response.headers.get("Content-Length")
                if content_length and int(content_length) > max_bytes:
                    raise ProcessorError("image exceeds the configured byte limit")
                payload = response.read(max_bytes + 1)
                if len(payload) > max_bytes:
                    raise ProcessorError("image exceeds the configured byte limit")
                if not payload:
                    raise ProcessorError("image response was empty")
                if content_type in {"text/html", "application/json", "text/plain"}:
                    raise ProcessorError(f"URL returned {content_type}, not an image")
                return payload
        except (ProcessorError, urllib.error.URLError, TimeoutError, socket.timeout) as exc:
            last_error = exc
            if isinstance(exc, ProcessorError) or attempt == attempts:
                break
            time.sleep(min(2 ** (attempt - 1), 4))
    raise ProcessorError(safe_error(last_error or ProcessorError("download failed")))


class RapidOCREngine:
    def __init__(self, model_type: str = "small") -> None:
        try:
            import numpy as np
            from PIL import Image, ImageOps, UnidentifiedImageError
            from rapidocr import EngineType, ModelType, OCRVersion, RapidOCR
        except ImportError as exc:
            raise ProcessorError(
                "OCR dependencies are missing. Install Brightdata/rapidocr/requirements.txt."
            ) from exc

        if model_type not in MODEL_TYPES:
            raise ProcessorError(f"Unsupported PP-OCRv6 model tier: {model_type}")

        self._np = np
        self._image_class = Image
        self._image_ops = ImageOps
        self._unidentified_image_error = UnidentifiedImageError
        model = ModelType(model_type)
        self._engine = RapidOCR(
            params={
                "Det.engine_type": EngineType.ONNXRUNTIME,
                "Det.model_type": model,
                "Det.ocr_version": OCRVersion.PPOCRV6,
                "Rec.engine_type": EngineType.ONNXRUNTIME,
                "Rec.model_type": model,
                "Rec.ocr_version": OCRVersion.PPOCRV6,
            }
        )
        self.model_type = model_type
        self.name = (
            f"RapidOCR {package_version('rapidocr')} / PP-OCRv6 {model_type} / "
            f"ONNX Runtime {package_version('onnxruntime')}"
        )

    def recognize(self, payload: bytes) -> str:
        try:
            with self._image_class.open(io.BytesIO(payload)) as source:
                source.load()
                image = self._image_ops.exif_transpose(source).convert("RGB")
                pixels = image.width * image.height
                if pixels > 80_000_000:
                    raise ProcessorError("decoded image exceeds the 80-megapixel limit")
                image_array = self._np.asarray(image)
        except self._unidentified_image_error as exc:
            raise ProcessorError("downloaded content is not a supported image") from exc
        except (OSError, ValueError) as exc:
            raise ProcessorError(f"image decode failed: {type(exc).__name__}") from exc

        result = self._engine(image_array)
        texts = getattr(result, "txts", ()) or ()
        return clean_ocr_lines(texts)


def cache_key(url: str, engine_name: str) -> str:
    material = f"{engine_name}\0{url}".encode("utf-8")
    return hashlib.sha256(material).hexdigest()


def default_artifact_paths(input_path: Path, model_type: str) -> tuple[Path, Path, Path]:
    suffix = f"rapidocr_ppocrv6_{model_type}"
    stem = input_path.stem
    return (
        SCRIPT_DIR / "output" / f"{stem}_{suffix}.json",
        SCRIPT_DIR / "output" / f"{stem}_{suffix}.csv",
        SCRIPT_DIR / "cache" / f"{stem}_{suffix}_cache.json",
    )


def load_cache(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProcessorError(f"Could not read OCR cache: {safe_error(exc)}") from exc
    if not isinstance(value, dict):
        raise ProcessorError("OCR cache must be a JSON object")
    return {
        str(key): item
        for key, item in value.items()
        if isinstance(item, dict) and isinstance(item.get("status"), str)
    }


def atomic_json_write(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        newline="\n",
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=path.parent,
        delete=False,
    )
    temporary_path = Path(handle.name)
    try:
        with handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temporary_path, path)
    except BaseException:
        temporary_path.unlink(missing_ok=True)
        raise


def recognize_url(
    source: ImageSource,
    *,
    engine: RapidOCREngine,
    cache: dict[str, dict[str, str]],
    use_cache: bool,
    timeout_seconds: float,
    max_bytes: int,
    attempts: int,
    downloader: Callable[..., bytes] = download_image,
) -> OCRResult:
    key = cache_key(source.url, engine.name)
    if use_cache and key in cache:
        item = cache[key]
        return OCRResult(
            text=str(item.get("text") or ""),
            status=str(item.get("status") or "error"),
            error=str(item.get("error") or ""),
        )

    try:
        payload = downloader(
            source.url,
            timeout_seconds=timeout_seconds,
            max_bytes=max_bytes,
            attempts=attempts,
        )
        text = engine.recognize(payload)
        result = OCRResult(text=text, status="success" if text else "no_text")
    except Exception as exc:  # Continue after isolated media failures.
        result = OCRResult(text="", status="error", error=safe_error(exc))

    cache[key] = {"text": result.text, "status": result.status, "error": result.error}
    return result


def append_result(
    attachment: dict[str, Any],
    source: ImageSource | None,
    result: OCRResult,
    engine_name: str,
) -> None:
    attachment["ocrText"] = result.text
    attachment["ocrStatus"] = result.status
    attachment["ocrEngine"] = engine_name
    if source:
        attachment["ocrSourceField"] = source.field
    else:
        attachment.pop("ocrSourceField", None)
    if result.error:
        attachment["ocrError"] = result.error
    else:
        attachment.pop("ocrError", None)


def enrich_rows(
    rows: list[dict[str, Any]],
    *,
    engine: RapidOCREngine,
    cache: dict[str, dict[str, str]],
    use_cache: bool,
    timeout_seconds: float,
    max_bytes: int,
    attempts: int,
    progress: bool = True,
    cache_checkpoint: Callable[[], None] | None = None,
) -> tuple[list[dict[str, Any]], Counter[str]]:
    enriched = copy.deepcopy(rows)
    totals: Counter[str] = Counter()
    processed_at = datetime.now(timezone.utc).isoformat(timespec="seconds")

    for row_index, row in enumerate(enriched, start=1):
        attachments = row.get("attachments")
        if not isinstance(attachments, list):
            attachments = []
            row["attachments"] = attachments

        sources_seen: set[str] = set()
        row_texts: list[str] = []
        candidates = 0
        for attachment in attachments:
            if not isinstance(attachment, dict):
                totals["invalid_attachment"] += 1
                continue
            source = choose_attachment_image(attachment)
            if source is None:
                append_result(
                    attachment,
                    None,
                    OCRResult(text="", status="no_image_url"),
                    engine.name,
                )
                totals["no_image_url"] += 1
                continue

            candidates += 1
            sources_seen.add(source.url)
            result = recognize_url(
                source,
                engine=engine,
                cache=cache,
                use_cache=use_cache,
                timeout_seconds=timeout_seconds,
                max_bytes=max_bytes,
                attempts=attempts,
            )
            append_result(attachment, source, result, engine.name)
            totals[result.status] += 1
            if result.text:
                row_texts.append(result.text)

        fallback_value = row.get("post_image")
        if nonempty(fallback_value):
            try:
                fallback_url = validate_remote_url(fallback_value)
            except ProcessorError:
                fallback_url = ""
            if fallback_url and fallback_url not in sources_seen:
                candidates += 1
                source = ImageSource(field="post_image", url=fallback_url)
                result = recognize_url(
                    source,
                    engine=engine,
                    cache=cache,
                    use_cache=use_cache,
                    timeout_seconds=timeout_seconds,
                    max_bytes=max_bytes,
                    attempts=attempts,
                )
                row["post_image_ocrText"] = result.text
                row["post_image_ocrStatus"] = result.status
                if result.error:
                    row["post_image_ocrError"] = result.error
                else:
                    row.pop("post_image_ocrError", None)
                totals[result.status] += 1
                if result.text:
                    row_texts.append(result.text)

        row["ocr_texts"] = row_texts
        row["ocr_processing"] = {
            "engine": engine.name,
            "processed_at_utc": processed_at,
            "candidate_images": candidates,
            "images_with_text": len(row_texts),
        }
        totals["rows"] += 1
        if row_texts:
            totals["rows_with_text"] += 1
        if progress:
            print(
                f"[{row_index:>3}/{len(enriched)}] candidates={candidates} "
                f"with_text={len(row_texts)}",
                flush=True,
            )
        if cache_checkpoint:
            cache_checkpoint()

    return enriched, totals


def scalar_or_blank(value: object) -> object:
    if value is None or isinstance(value, (dict, list)):
        return ""
    return value


def classifier_csv_rows(rows: list[dict[str, Any]]) -> tuple[list[str], list[dict[str, object]]]:
    max_ocr = max((len(row.get("ocr_texts") or []) for row in rows), default=0)
    ocr_columns = [f"media/{index}/ocrText" for index in range(max_ocr)]
    output_rows: list[dict[str, object]] = []
    for row in rows:
        likes = row.get("likes")
        if not nonempty(likes):
            likes = row.get("num_likes_type")
        mapped: dict[str, object] = {
            "url": scalar_or_blank(row.get("url")),
            "text": scalar_or_blank(row.get("content")),
            "likes": scalar_or_blank(likes),
            "comments": scalar_or_blank(row.get("num_comments")),
            "shares": scalar_or_blank(row.get("num_shares")),
            "date_posted": scalar_or_blank(row.get("date_posted")),
            "post_id": scalar_or_blank(row.get("post_id")),
            "post_type": scalar_or_blank(row.get("post_type")),
            "page_name": scalar_or_blank(row.get("page_name")),
            "page_url": scalar_or_blank(row.get("page_url")),
            "page_followers": scalar_or_blank(row.get("page_followers")),
            "page_category": scalar_or_blank(row.get("page_category")),
            "is_sponsored": scalar_or_blank(row.get("is_sponsored")),
        }
        for index, text in enumerate(row.get("ocr_texts") or []):
            mapped[f"media/{index}/ocrText"] = text
        output_rows.append(mapped)
    return CSV_BASE_COLUMNS + ocr_columns, output_rows


def write_classifier_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    columns, mapped_rows = classifier_csv_rows(rows)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(mapped_rows)


def load_rows(path: Path) -> list[dict[str, Any]]:
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError as exc:
        raise ProcessorError(f"Input file does not exist: {path}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise ProcessorError(f"Could not read input JSON: {safe_error(exc)}") from exc
    if not isinstance(value, list) or not all(isinstance(item, dict) for item in value):
        raise ProcessorError("Bright Data input must be a JSON array of objects")
    return value


def run_self_test() -> None:
    assert choose_attachment_image(
        {
            "type": "Photo",
            "url": "https://scontent.example.test/image.jpg",
            "attachment_url": "https://www.facebook.com/photo/1",
        }
    ) == ImageSource("url", "https://scontent.example.test/image.jpg")
    assert choose_attachment_image(
        {
            "type": "Video",
            "url": "https://video.example.test/video.mp4",
            "thumbnail_url": "https://scontent.example.test/poster.jpg",
        }
    ) == ImageSource("thumbnail_url", "https://scontent.example.test/poster.jpg")
    assert choose_attachment_image(
        {"type": "Photo", "url": "https://www.facebook.com/photo/1"}
    ) is None
    assert clean_ocr_lines(["  ZERO   DOWNPAYMENT ", "P1,500", "P1,500"]) == (
        "ZERO DOWNPAYMENT\nP1,500"
    )
    assert cache_key("https://example.test/image.jpg", "small") != cache_key(
        "https://example.test/image.jpg", "medium"
    )
    output_path, csv_path, cache_path = default_artifact_paths(
        Path("sample.json"), "medium"
    )
    assert output_path.name == "sample_rapidocr_ppocrv6_medium.json"
    assert csv_path.name == "sample_rapidocr_ppocrv6_medium.csv"
    assert cache_path.name == "sample_rapidocr_ppocrv6_medium_cache.json"

    sample = [
        {
            "url": "https://www.facebook.com/example/posts/1",
            "content": "Caption",
            "num_comments": 2,
            "num_shares": 3,
            "ocr_texts": ["MODEL A\nP100 DAILY", "FREE HELMET"],
        }
    ]
    columns, mapped = classifier_csv_rows(sample)
    assert columns[-2:] == ["media/0/ocrText", "media/1/ocrText"]
    assert mapped[0]["text"] == "Caption"
    assert mapped[0]["comments"] == 2
    assert mapped[0]["media/1/ocrText"] == "FREE HELMET"
    print("Self-test passed: 12 deterministic checks.")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--csv-output", type=Path)
    parser.add_argument("--cache", type=Path)
    parser.add_argument(
        "--model-type",
        choices=MODEL_TYPES,
        default="small",
        help="PP-OCRv6 tier for both text detection and recognition (default: small).",
    )
    parser.add_argument("--limit", type=int, default=0, help="Process only the first N posts.")
    parser.add_argument("--timeout", type=float, default=30.0)
    parser.add_argument("--attempts", type=int, default=3)
    parser.add_argument("--max-image-mb", type=float, default=25.0)
    parser.add_argument("--no-cache", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.self_test:
        run_self_test()
        return 0
    if args.limit < 0 or args.timeout <= 0 or args.attempts < 1 or args.max_image_mb <= 0:
        print("Error: numeric options must be positive.", file=sys.stderr)
        return 2

    try:
        input_path = args.input.resolve()
        default_output, default_csv, default_cache = default_artifact_paths(
            input_path, args.model_type
        )
        output_path = (args.output or default_output).resolve()
        csv_output_path = (args.csv_output or default_csv).resolve()
        cache_path = (args.cache or default_cache).resolve()
        if input_path in {output_path, csv_output_path}:
            raise ProcessorError("Output paths must not overwrite the source JSON")

        rows = load_rows(input_path)
        if args.limit:
            rows = rows[: args.limit]
        cache = {} if args.no_cache else load_cache(cache_path)
        engine = RapidOCREngine(args.model_type)
        print(f"Loaded {len(rows)} Bright Data posts.")
        print(f"OCR engine: {engine.name}")
        enriched, totals = enrich_rows(
            rows,
            engine=engine,
            cache=cache,
            use_cache=not args.no_cache,
            timeout_seconds=args.timeout,
            max_bytes=int(args.max_image_mb * 1024 * 1024),
            attempts=args.attempts,
            cache_checkpoint=(
                None
                if args.no_cache
                else lambda: atomic_json_write(cache_path, cache)
            ),
        )
        atomic_json_write(output_path, enriched)
        write_classifier_csv(csv_output_path, enriched)
        if not args.no_cache:
            atomic_json_write(cache_path, cache)

        print(f"Enriched JSON: {output_path}")
        print(f"Classifier CSV: {csv_output_path}")
        print(
            "Summary: "
            f"rows={totals['rows']}, rows_with_ocr={totals['rows_with_text']}, "
            f"images_with_ocr={totals['success']}, no_text={totals['no_text']}, "
            f"errors={totals['error']}"
        )
        return 0
    except ProcessorError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
