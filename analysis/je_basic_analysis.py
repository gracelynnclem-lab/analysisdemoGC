#!/usr/bin/env python3
"""Basic analysis for je_samples.xlsx without third-party dependencies."""
from __future__ import annotations

import argparse
import datetime as dt
import math
import re
import statistics
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_to_index(col: str) -> int:
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c) - 64)
    return idx


def read_sheet(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as workbook:
        shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
        shared_strings = [
            "".join(t.text or "" for t in si.findall(".//s:t", NS))
            for si in shared_root.findall("s:si", NS)
        ]
        sheet_root = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))

    rows: list[dict[int, str]] = []
    max_col = 0
    for row in sheet_root.findall("s:sheetData/s:row", NS):
        row_dict: dict[int, str] = {}
        for cell in row.findall("s:c", NS):
            ref = cell.attrib.get("r")
            if not ref:
                continue
            match = re.match(r"([A-Z]+)(\d+)", ref)
            if not match:
                continue
            col_idx = col_to_index(match.group(1))
            max_col = max(max_col, col_idx)
            cell_type = cell.attrib.get("t")
            value = ""
            if cell_type == "s":
                v = cell.find("s:v", NS)
                value = shared_strings[int(v.text)] if v is not None else ""
            elif cell_type == "inlineStr":
                t = cell.find("s:is/s:t", NS)
                value = t.text if t is not None else ""
            else:
                v = cell.find("s:v", NS)
                value = v.text if v is not None else ""
            row_dict[col_idx] = value
        rows.append(row_dict)

    table = []
    for row_dict in rows:
        table.append([row_dict.get(i, "") for i in range(1, max_col + 1)])
    return table


def parse_date(value: str) -> dt.date | None:
    if value is None:
        return None
    cleaned = value.strip()
    if cleaned == "":
        return None
    if re.fullmatch(r"\d+(\.\d+)?", cleaned):
        serial = float(cleaned)
        base = dt.date(1899, 12, 30)
        try:
            return base + dt.timedelta(days=serial)
        except OverflowError:
            return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y"):
        try:
            return dt.datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    if re.fullmatch(r"\d{4}-\d{2}", cleaned):
        try:
            return dt.datetime.strptime(cleaned + "-01", "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def safe_float(value: str) -> float | None:
    if value is None:
        return None
    cleaned = value.strip()
    if cleaned == "":
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def describe_numeric(values: list[float]) -> dict[str, float]:
    if not values:
        return {}
    return {
        "min": min(values),
        "max": max(values),
        "mean": statistics.mean(values),
        "median": statistics.median(values),
        "sum": sum(values),
    }


def svg_bar_chart(labels: list[str], values: list[int], title: str, path: Path) -> None:
    width, height = 900, 450
    margin = 60
    chart_width = width - 2 * margin
    chart_height = height - 2 * margin
    max_val = max(values) if values else 1

    bars = []
    for i, (label, value) in enumerate(zip(labels, values)):
        bar_width = chart_width / len(values)
        x = margin + i * bar_width
        bar_height = (value / max_val) * chart_height
        y = margin + (chart_height - bar_height)
        bars.append(
            f"<rect x='{x:.1f}' y='{y:.1f}' width='{bar_width * 0.8:.1f}' height='{bar_height:.1f}' fill='#4c78a8' />"
        )
        bars.append(
            f"<text x='{x + bar_width * 0.4:.1f}' y='{height - margin + 20}' font-size='10' text-anchor='middle'>{label}</text>"
        )
        bars.append(
            f"<text x='{x + bar_width * 0.4:.1f}' y='{y - 5:.1f}' font-size='10' text-anchor='middle'>{value}</text>"
        )

    svg = [
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}'>",
        f"<text x='{width / 2}' y='30' text-anchor='middle' font-size='16' font-family='sans-serif'>{title}</text>",
        f"<line x1='{margin}' y1='{margin}' x2='{margin}' y2='{height - margin}' stroke='black' />",
        f"<line x1='{margin}' y1='{height - margin}' x2='{width - margin}' y2='{height - margin}' stroke='black' />",
        *bars,
        "</svg>",
    ]
    path.write_text("\n".join(svg), encoding="utf-8")


def svg_histogram(values: list[float], title: str, path: Path, bins: int = 20) -> None:
    if not values:
        path.write_text("<svg xmlns='http://www.w3.org/2000/svg' width='300' height='100'></svg>")
        return
    min_val, max_val = min(values), max(values)
    if math.isclose(min_val, max_val):
        min_val -= 1
        max_val += 1
    bin_width = (max_val - min_val) / bins
    counts = [0] * bins
    for value in values:
        idx = int((value - min_val) / bin_width)
        if idx == bins:
            idx -= 1
        counts[idx] += 1

    width, height = 900, 450
    margin = 60
    chart_width = width - 2 * margin
    chart_height = height - 2 * margin
    max_count = max(counts) if counts else 1

    bars = []
    for i, count in enumerate(counts):
        bar_width = chart_width / bins
        x = margin + i * bar_width
        bar_height = (count / max_count) * chart_height
        y = margin + (chart_height - bar_height)
        bars.append(
            f"<rect x='{x:.1f}' y='{y:.1f}' width='{bar_width * 0.9:.1f}' height='{bar_height:.1f}' fill='#f58518' />"
        )

    svg = [
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}'>",
        f"<text x='{width / 2}' y='30' text-anchor='middle' font-size='16' font-family='sans-serif'>{title}</text>",
        f"<line x1='{margin}' y1='{margin}' x2='{margin}' y2='{height - margin}' stroke='black' />",
        f"<line x1='{margin}' y1='{height - margin}' x2='{width - margin}' y2='{height - margin}' stroke='black' />",
        *bars,
        "</svg>",
    ]
    path.write_text("\n".join(svg), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run a basic analysis on a JE Excel export and generate artifacts."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=ROOT / "je_samples.xlsx",
        help="Path to the Excel workbook to analyze.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "analysis_outputs",
        help="Directory to write summary and chart artifacts.",
    )
    args = parser.parse_args()

    if not args.input.exists():
        raise SystemExit(f"Input file not found: {args.input}")

    output_dir: Path = args.output_dir
    output_dir.mkdir(exist_ok=True)
    summary_path = output_dir / "summary.txt"
    source_chart_path = output_dir / "source_counts.svg"
    amount_hist_path = output_dir / "amount_histogram.svg"

    table = read_sheet(args.input)
    header = table[0]
    data_rows = table[1:]

    columns = {name: [] for name in header}
    for row in data_rows:
        for idx, name in enumerate(header):
            columns[name].append(row[idx] if idx < len(row) else "")

    summary_lines = []
    summary_lines.append("Basic Journal Entry (JE) Analysis")
    summary_lines.append("================================")
    summary_lines.append(f"Total rows (excluding header): {len(data_rows)}")
    summary_lines.append(f"Total columns: {len(header)}")
    summary_lines.append("")

    summary_lines.append("Column overview:")
    for name in header:
        values = columns[name]
        non_empty = [v for v in values if str(v).strip() != ""]
        unique_count = len(set(non_empty))
        summary_lines.append(
            f"- {name}: non-empty={len(non_empty):,}, empty={len(values) - len(non_empty):,}, unique={unique_count:,}"
        )

    summary_lines.append("")
    summary_lines.append("Date ranges:")
    for name in header:
        if "date" in name.lower() or name.lower() == "period":
            parsed = [parse_date(v) for v in columns[name]]
            parsed = [d for d in parsed if d is not None]
            if parsed:
                summary_lines.append(
                    f"- {name}: {min(parsed)} to {max(parsed)} (count={len(parsed):,})"
                )
            else:
                summary_lines.append(f"- {name}: no parseable dates")

    summary_lines.append("")
    summary_lines.append("Numeric summaries (auto-detected):")
    for name in header:
        values = columns[name]
        numeric_values = [safe_float(v) for v in values]
        numeric_values = [v for v in numeric_values if v is not None]
        if numeric_values and len(numeric_values) / len(values) >= 0.8:
            stats = describe_numeric(numeric_values)
            summary_lines.append(
                f"- {name}: min={stats['min']:.2f}, max={stats['max']:.2f}, mean={stats['mean']:.2f}, median={stats['median']:.2f}, sum={stats['sum']:.2f}"
            )

    summary_lines.append("")
    summary_lines.append("Top categorical values (top 5):")
    for name in header:
        values = [v.strip() for v in columns[name] if str(v).strip() != ""]
        if not values:
            continue
        unique_count = len(set(values))
        if unique_count <= 1:
            continue
        if unique_count <= 100:
            top_counts = Counter(values).most_common(5)
            formatted = "; ".join([f"{val} ({count})" for val, count in top_counts])
            summary_lines.append(f"- {name}: {formatted}")

    summary_path.write_text("\n".join(summary_lines), encoding="utf-8")

    # Charts
    source_values = [v.strip() for v in columns.get("Source", []) if v.strip()]
    if source_values:
        top_sources = Counter(source_values).most_common(10)
        labels = [val.strip() for val, _ in top_sources]
        counts = [count for _, count in top_sources]
        svg_bar_chart(labels, counts, "Top 10 Sources by JE Count", source_chart_path)

    amount_values = [safe_float(v) for v in columns.get("Amount", [])]
    amount_values = [v for v in amount_values if v is not None]
    if amount_values:
        svg_histogram(amount_values, "Amount Distribution (Histogram)", amount_hist_path)


if __name__ == "__main__":
    main()
