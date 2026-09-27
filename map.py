#!/usr/bin/env python3
"""Validate the workbook and build courses.json for the EveryCourse map."""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import urllib.parse
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from openpyxl import load_workbook

SCHEMA_VERSION = 3
REQUIRED_COLUMNS = ("Course", "Address", "City", "Type", "Region", "Lat", "Long")
ALLOWED_TYPES = {
    "public": "Public", "private": "Private", "semi private": "Semi-Private",
    "semi-private": "Semi-Private", "semi": "Semi-Private", "resort": "Resort",
    "permanently closed": "Permanently Closed", "permanently-closed": "Permanently Closed",
    "closed": "Permanently Closed",
}
CO_LAT_RANGE = (36.8, 41.2)
CO_LNG_RANGE = (-109.2, -101.8)
CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
LEADING_NUMBER = re.compile(r"^[\s\u00a0]*([+-]?(?:\d+(?:\.\d*)?|\.\d+))")


class BuildError(Exception):
    pass


def clean_text(value: Any, *, limit: int = 5000) -> str:
    if value is None:
        return ""
    return CONTROL_CHARS.sub("", str(value)).strip()[:limit]


def parse_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        number = float(value)
        return int(number) if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def parse_date(value: Any) -> str | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = clean_text(value, limit=50)
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def parse_coordinate(value: Any, label: str, course: str, warnings: list[str]) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
    else:
        text = clean_text(value, limit=100)
        match = LEADING_NUMBER.match(text)
        if not match:
            raise BuildError(f"{course}: invalid {label} value {value!r}")
        number = float(match.group(1))
        if match.group(1) != text:
            warnings.append(f"{course}: cleaned malformed {label} value {value!r}")
    if not math.isfinite(number):
        raise BuildError(f"{course}: non-finite {label} value")
    return number


def normalize_type(value: Any, course: str) -> str:
    raw = clean_text(value, limit=80)
    normalized = ALLOWED_TYPES.get(raw.lower())
    if not normalized:
        raise BuildError(f"{course}: unsupported course type {raw!r}")
    return normalized


def safe_url(value: Any, *, allowed_hosts: tuple[str, ...] = ()) -> str | None:
    text = clean_text(value, limit=2048)
    if not text:
        return None
    try:
        parsed = urlparse(text)
    except ValueError:
        return None
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return None
    host = (parsed.hostname or "").lower()
    if allowed_hosts and not any(host == h or host.endswith("." + h) for h in allowed_hosts):
        return None
    return text


def map_links(address: str, name: str, city: str) -> tuple[str, str]:
    query = address or f"{name}, {city}, Colorado"
    encoded = urllib.parse.quote(query)
    return (f"https://maps.apple.com/?q={encoded}",
            f"https://www.google.com/maps/search/?api=1&query={encoded}")


def cell_link(cell: Any) -> str | None:
    if cell.hyperlink and cell.hyperlink.target:
        return safe_url(cell.hyperlink.target)
    return safe_url(cell.value)


def build(input_path: Path, output_path: Path, strict: bool = False) -> dict[str, Any]:
    workbook = load_workbook(input_path, data_only=True, read_only=False)
    sheet = workbook.active
    headers = {clean_text(sheet.cell(1, c).value, limit=100): c for c in range(1, sheet.max_column + 1)
               if clean_text(sheet.cell(1, c).value, limit=100)}
    missing = [column for column in REQUIRED_COLUMNS if column not in headers]
    if missing:
        raise BuildError(f"Missing required columns: {', '.join(missing)}")

    warnings: list[str] = []
    errors: list[str] = []
    courses: list[dict[str, Any]] = []
    names_seen: set[str] = set()

    for row_number in range(2, sheet.max_row + 1):
        def value(column: str) -> Any:
            index = headers.get(column)
            return sheet.cell(row_number, index).value if index else None

        name = clean_text(value("Course"), limit=250)
        if not name:
            continue
        try:
            if name.casefold() in names_seen:
                raise BuildError(f"duplicate course name {name!r}")
            names_seen.add(name.casefold())
            lat = parse_coordinate(value("Lat"), "latitude", name, warnings)
            lng = parse_coordinate(value("Long"), "longitude", name, warnings)
            if lng > 0 and CO_LNG_RANGE[0] <= -lng <= CO_LNG_RANGE[1]:
                lng = -lng
                warnings.append(f"{name}: corrected positive Colorado longitude")
            if not CO_LAT_RANGE[0] <= lat <= CO_LAT_RANGE[1]:
                raise BuildError(f"{name}: latitude {lat} is outside Colorado")
            if not CO_LNG_RANGE[0] <= lng <= CO_LNG_RANGE[1]:
                raise BuildError(f"{name}: longitude {lng} is outside Colorado")

            course_type = normalize_type(value("Type"), name)
            address = clean_text(value("Address"), limit=500)
            city = clean_text(value("City"), limit=150)
            region = clean_text(value("Region"), limit=150)
            played_date = parse_date(value("1st Played"))
            reel_cell = sheet.cell(row_number, headers["Reel"]) if "Reel" in headers else None
            reel_url = cell_link(reel_cell) if reel_cell else None
            if reel_url and not safe_url(reel_url, allowed_hosts=("instagram.com", "youtu.be", "youtube.com", "tiktok.com", "facebook.com")):
                warnings.append(f"{name}: unrecognized video host; URL omitted")
                reel_url = None
            apple_maps, google_maps = map_links(address, name, city)
            courses.append({
                "id": len(courses) + 1, "name": name, "city": city, "region": region,
                "type": course_type, "address": address, "lat": round(lat, 7),
                "lng": round(lng, 7), "played": played_date is not None,
                "order": parse_int(value("Order")), "first_played": played_date,
                "video_url": reel_url, "has_video": reel_url is not None,
                "cost_level": parse_int(value("cost_level")),
                "accessibility_level": parse_int(value("accessibility_level")),
                "cam_ranking": parse_int(value("cam_ranking")),
                "anecdote": clean_text(value("anecdote"), limit=5000),
                "apple_maps": apple_maps, "google_maps": google_maps,
            })
        except BuildError as exc:
            errors.append(f"Row {row_number}: {exc}")

    if errors:
        raise BuildError("Validation failed:\n" + "\n".join(f"- {item}" for item in errors))
    if strict and warnings:
        raise BuildError("Strict validation failed:\n" + "\n".join(f"- {item}" for item in warnings))
    if not courses:
        raise BuildError("No valid courses found")

    now = datetime.now(timezone.utc)
    payload = {"meta": {"schema_version": SCHEMA_VERSION,
                         "generated_at": now.isoformat().replace("+00:00", "Z"),
                         "generated_at_unix": int(now.timestamp()), "count": len(courses),
                         "source_file": input_path.name, "warnings": warnings},
               "courses": courses}
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return payload


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate the EveryCourse workbook and build courses.json")
    parser.add_argument("--input", default="co_courses.xlsx", type=Path)
    parser.add_argument("--output", default="courses.json", type=Path)
    parser.add_argument("--strict", action="store_true", help="Treat data-cleanup warnings as errors")
    args = parser.parse_args()
    try:
        payload = build(args.input, args.output, args.strict)
    except (BuildError, OSError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(f"Wrote {payload['meta']['count']} courses to {args.output}")
    for warning in payload["meta"]["warnings"]:
        print(f"WARNING: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
