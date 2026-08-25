"""Validate and normalize the JSON contract consumed by the GSWM COM deck builder."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class DeckContentError(RuntimeError):
    pass


def _text(value: object, field: str, *, limit: int) -> str:
    if not isinstance(value, str) or not value.strip():
        raise DeckContentError(f"{field} is required.")
    result = value.strip()
    if len(result) > limit:
        raise DeckContentError(f"{field} is too long ({len(result)} characters; maximum {limit}).")
    return result


def _list(value: object, field: str) -> list[Any]:
    if not isinstance(value, list):
        raise DeckContentError(f"{field} must be a list.")
    return value


def _scalar(value: object, field: str) -> str:
    if value is None or isinstance(value, (dict, list)):
        raise DeckContentError(f"{field} must be a displayable value.")
    return str(value)


def _image_reference(value: object, field: str, images: dict[str, Path]) -> str:
    reference = _text(value, field, limit=260)
    basename = Path(reference).name.casefold()
    matches = [path for name, path in images.items() if name.casefold() == basename]
    if len(matches) != 1:
        raise DeckContentError(f"{field} must match exactly one uploaded chart/image filename.")
    return str(matches[0].resolve())


def _validate_table(table: object, slide_number: int) -> dict[str, object]:
    if not isinstance(table, dict):
        raise DeckContentError(f"Slide {slide_number}: table must be an object.")
    headers = [_text(item, f"Slide {slide_number} table header", limit=40) for item in _list(table.get("headers"), f"Slide {slide_number} table headers")]
    if not 2 <= len(headers) <= 8:
        raise DeckContentError(f"Slide {slide_number}: tables require 2–8 columns.")
    source_rows = _list(table.get("rows"), f"Slide {slide_number} table rows")
    if not 1 <= len(source_rows) <= 18:
        raise DeckContentError(f"Slide {slide_number}: tables require 1–18 rows.")
    rows: list[list[str]] = []
    for row_number, row in enumerate(source_rows, start=1):
        values = _list(row, f"Slide {slide_number} table row {row_number}")
        if len(values) != len(headers):
            raise DeckContentError(f"Slide {slide_number} table row {row_number} has the wrong column count.")
        rows.append([_scalar(value, f"Slide {slide_number} table row {row_number}") for value in values])
    result: dict[str, object] = {"headers": headers, "rows": rows}
    if "color_column" in table:
        color_column = table["color_column"]
        if not isinstance(color_column, int) or not 0 <= color_column < len(headers):
            raise DeckContentError(f"Slide {slide_number}: color_column is outside the table.")
        result["color_column"] = color_column
    if "col_widths" in table:
        widths = _list(table["col_widths"], f"Slide {slide_number} col_widths")
        if len(widths) != len(headers) or not all(isinstance(item, (int, float)) and item > 0 for item in widths):
            raise DeckContentError(f"Slide {slide_number}: col_widths must contain one positive number per column.")
        total = float(sum(widths))
        result["col_widths"] = [float(item) / total for item in widths]
    return result


def _validate_picks(value: object, slide_number: int) -> dict[str, object]:
    if not isinstance(value, dict):
        raise DeckContentError(f"Slide {slide_number}: picks_table must be an object.")
    source_rows = _list(value.get("rows"), f"Slide {slide_number} picks rows")
    if not 1 <= len(source_rows) <= 10:
        raise DeckContentError(f"Slide {slide_number}: picks tables require 1–10 rows.")
    rows: list[dict[str, str]] = []
    for row_number, row in enumerate(source_rows, start=1):
        if not isinstance(row, dict):
            raise DeckContentError(f"Slide {slide_number} picks row {row_number} must be an object.")
        if row.get("type") == "banner":
            rows.append({"type": "banner", "text": _text(row.get("text"), f"Slide {slide_number} banner", limit=70)})
            continue
        rating_color = str(row.get("rating_color", "neutral")).casefold()
        if rating_color not in {"green", "amber", "red", "neutral"}:
            raise DeckContentError(f"Slide {slide_number} picks row {row_number}: invalid rating_color.")
        rows.append({
            "type": "stock",
            "ticker": _text(row.get("ticker"), f"Slide {slide_number} ticker", limit=12),
            "company": _text(row.get("company"), f"Slide {slide_number} company", limit=55),
            "rating": _text(row.get("rating"), f"Slide {slide_number} rating", limit=20),
            "rating_color": rating_color,
            "return": _text(row.get("return"), f"Slide {slide_number} return", limit=15),
            "vs_sp": _text(row.get("vs_sp"), f"Slide {slide_number} vs_sp", limit=15),
        })
    return {"rows": rows}


def deck_content_schema() -> dict[str, object]:
    """Return the provider-facing low-density bullet-deck schema."""

    return {
        "type": "object",
        "required": ["slides"],
        "properties": {
            "slides": {
                "type": "array",
                "minItems": 1,
                "maxItems": 40,
                "items": {
                    "type": "object",
                    "required": ["title", "bullets", "footer_right"],
                    "properties": {
                        "title": {"type": "string", "maxLength": 90},
                        "bullets": {
                            "type": "array", "minItems": 1, "maxItems": 8,
                            "items": {"type": "string", "maxLength": 240},
                        },
                        "footer_right": {"type": "string", "maxLength": 120},
                    },
                },
            }
        },
    }


def normalize_deck_payload(payload: dict[str, object], *, report_title: str,
                           image_paths: tuple[Path, ...] = ()) -> dict[str, object]:
    if not isinstance(payload, dict):
        raise DeckContentError("The PowerPoint content JSON must contain one object.")
    title = _text(report_title, "Presentation title", limit=90)
    source_slides = _list(payload.get("slides"), "slides")
    if not 1 <= len(source_slides) <= 40:
        raise DeckContentError("A presentation requires 1–40 content slides.")
    images = {path.name: path for path in image_paths}
    if len(images) != len(image_paths):
        raise DeckContentError("Uploaded chart/image filenames must be unique.")

    slides: list[dict[str, object]] = []
    for number, raw in enumerate(source_slides, start=1):
        if not isinstance(raw, dict):
            raise DeckContentError(f"Slide {number} must be an object.")
        slide: dict[str, object] = {"title": _text(raw.get("title"), f"Slide {number} title", limit=90)}
        for field in ("footer_left", "footer_right"):
            if field in raw:
                slide[field] = _text(raw[field], f"Slide {number} {field}", limit=120)
        if "header_right" in raw:
            header = [_text(item, f"Slide {number} header line", limit=80) for item in _list(raw["header_right"], f"Slide {number} header_right")]
            if len(header) > 2:
                raise DeckContentError(f"Slide {number}: header_right allows at most two lines.")
            slide["header_right"] = header

        modes = sum(key in raw for key in ("table", "picks_table", "footnote"))
        if modes > 1:
            raise DeckContentError(f"Slide {number}: choose one primary content layout.")
        if "table" in raw:
            slide["table"] = _validate_table(raw["table"], number)
        elif "picks_table" in raw:
            slide["picks_table"] = _validate_picks(raw["picks_table"], number)
        elif "footnote" in raw:
            slide["footnote"] = _text(raw["footnote"], f"Slide {number} footnote", limit=300)
            slide["chart_image"] = _image_reference(raw.get("chart_image"), f"Slide {number} chart_image", images)
        else:
            bullets = [_text(item, f"Slide {number} bullet", limit=240) for item in _list(raw.get("bullets"), f"Slide {number} bullets")]
            if not 1 <= len(bullets) <= 8:
                raise DeckContentError(f"Slide {number}: bullet slides require 1–8 bullets.")
            slide["bullets"] = bullets
            if raw.get("chart"):
                slide["chart"] = True
                slide["chart_image"] = _image_reference(raw.get("chart_image"), f"Slide {number} chart_image", images)
        if "footer_right" not in slide:
            slide["footer_right"] = "Source: Uploaded content"
        slides.append(slide)
    return {"title": title, "slides": slides}


def validate_deck_content(source: Path, *, report_title: str, image_paths: tuple[Path, ...] = ()) -> dict[str, object]:
    if source.suffix.lower() != ".json":
        raise DeckContentError(
            "This build accepts a structured .json content package. General document synthesis requires an AI provider."
        )
    try:
        payload = json.loads(source.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise DeckContentError("The PowerPoint content JSON is invalid.") from exc
    return normalize_deck_payload(payload, report_title=report_title, image_paths=image_paths)
