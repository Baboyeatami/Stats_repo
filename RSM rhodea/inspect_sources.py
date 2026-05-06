from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd


ROOT = Path(__file__).resolve().parent
DOCX_PATH = ROOT / "reference.docx"
XLSX_PATH = ROOT / "Book-of-Analysis-2026-rhodea.xlsx"
OUT_DIR = ROOT / "outputs"
OUT_DIR.mkdir(exist_ok=True)

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def text_from_element(el: ET.Element) -> str:
    pieces: list[str] = []
    for node in el.iter():
        if node.tag == f"{{{NS['w']}}}t" and node.text:
            pieces.append(node.text)
        elif node.tag == f"{{{NS['w']}}}tab":
            pieces.append("\t")
        elif node.tag == f"{{{NS['w']}}}br":
            pieces.append("\n")
    return "".join(pieces).strip()


def extract_docx_text(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find("w:body", NS)
    if body is None:
        return []

    blocks: list[str] = []
    for child in body:
        if child.tag == f"{{{NS['w']}}}p":
            text = text_from_element(child)
            if text:
                blocks.append(text)
        elif child.tag == f"{{{NS['w']}}}tbl":
            rows: list[str] = []
            for row in child.findall(".//w:tr", NS):
                cells = [text_from_element(cell) for cell in row.findall("./w:tc", NS)]
                if any(cells):
                    rows.append(" | ".join(cells))
            if rows:
                blocks.append("\n".join(rows))
    return blocks


def objective_methods_extract(blocks: list[str]) -> dict[str, list[str]]:
    joined = "\n".join(blocks)
    headings = [
        "objective",
        "objectives",
        "method",
        "methods",
        "methodology",
        "data analysis",
        "statistical analysis",
    ]
    matches = []
    for i, block in enumerate(blocks):
        clean = re.sub(r"\s+", " ", block).strip()
        if any(re.search(rf"\b{h}\b", clean, flags=re.I) for h in headings):
            matches.append((i, clean))

    windows: list[str] = []
    seen: set[int] = set()
    for i, _ in matches:
        for j in range(max(0, i - 1), min(len(blocks), i + 7)):
            if j not in seen:
                windows.append(blocks[j])
                seen.add(j)

    return {"matches": [m[1] for m in matches], "context": windows, "all_text": [joined]}


def workbook_summary(path: Path) -> dict:
    xls = pd.ExcelFile(path)
    summary: dict = {"sheets": []}
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        cols = []
        for col in df.columns:
            series = df[col]
            cols.append(
                {
                    "name": str(col),
                    "dtype": str(series.dtype),
                    "non_null": int(series.notna().sum()),
                    "missing": int(series.isna().sum()),
                    "unique": int(series.nunique(dropna=True)),
                    "sample": [str(x) for x in series.dropna().head(5).tolist()],
                }
            )
        summary["sheets"].append(
            {
                "name": sheet,
                "shape": [int(df.shape[0]), int(df.shape[1])],
                "columns": cols,
                "head": df.head(8).fillna("").astype(str).to_dict(orient="records"),
            }
        )
    return summary


def main() -> None:
    blocks = extract_docx_text(DOCX_PATH)
    doc_extract = objective_methods_extract(blocks)
    wb = workbook_summary(XLSX_PATH)

    (OUT_DIR / "reference_text.txt").write_text("\n\n".join(blocks), encoding="utf-8")
    (OUT_DIR / "source_summary.json").write_text(
        json.dumps({"docx": doc_extract, "workbook": wb}, indent=2),
        encoding="utf-8",
    )

    print("DOCX objective/method context:")
    for block in doc_extract["context"][:40]:
        print("-", re.sub(r"\\s+", " ", block)[:500])
    print("\nWorkbook sheets:")
    for sheet in wb["sheets"]:
        print(f"- {sheet['name']}: {sheet['shape'][0]} rows x {sheet['shape'][1]} cols")
        print("  " + ", ".join(c["name"] for c in sheet["columns"]))


if __name__ == "__main__":
    main()
