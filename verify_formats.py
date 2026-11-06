"""Generate the fixture catalog, export it, and write a round-trip matrix."""

import argparse
import csv
import re
from pathlib import Path
from xml.etree import ElementTree as XML
from zipfile import ZIP_STORED, ZipFile

from openpyxl import load_workbook

from converter import Cell, ConversionError, Table, read_batch, read_table_file
from exporters import FORMATS, WORKBOOK_FORMATS, save_table, save_workbook
from fixtures import catalog, generate

MIME_TYPES = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    ".xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
ODS_MEDIA = "application/vnd.oasis.opendocument.spreadsheet"
NS = {
    name: f"urn:oasis:names:tc:opendocument:xmlns:{name}:1.0"
    for name in ("office", "table", "text", "manifest")
}


def tag(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def expected_grid(table: Table) -> list[list[Cell]]:
    return [list(table.headers), *[
        [row.get(column) for column in range(len(table.headers))]
        for row in table.rows
    ]]


def assert_cells(actual: list[list[Cell]], expected: list[list[Cell]], *, excel: bool = False) -> None:
    assert len(actual) == len(expected), (len(actual), len(expected))
    for number, (row, reference) in enumerate(zip(actual, expected), start=1):
        assert len(row) == len(reference), number
        for column, (value, wanted) in enumerate(zip(row, reference), start=1):
            if excel and isinstance(value, str):
                value = re.sub(r"_x([0-9a-fA-F]{4})_", lambda m: chr(int(m[1], 16)), value)
            if excel and wanted == "":
                assert value in ("", None), (number, column)
            else:
                assert value == wanted, (number, column, str(value)[:60], str(wanted)[:60])
                assert isinstance(value, bool) == isinstance(wanted, bool), (number, column)


def verify_excel(path: Path, expected: dict[str, list[list[Cell]]]) -> None:
    with ZipFile(path) as archive:
        assert archive.testzip() is None
        manifest = XML.fromstring(archive.read("[Content_Types].xml"))
        assert any(
            part.get("PartName") == "/xl/workbook.xml"
            and part.get("ContentType") == MIME_TYPES[path.suffix]
            for part in manifest
        )
        assert not any("vbaProject" in member or "externalLink" in member for member in archive.namelist())
    workbook = load_workbook(path)
    try:
        assert workbook.sheetnames == list(expected)
        assert workbook.template == (path.suffix in (".xltx", ".xltm"))
        for name, reference in expected.items():
            sheet = workbook[name]
            assert sheet.freeze_panes == "A2" and sheet.auto_filter.ref
            assert all(cell.data_type != "f" and cell.hyperlink is None for row in sheet.rows for cell in row)
            assert_cells([list(row) for row in sheet.values], reference, excel=True)
    finally:
        workbook.close()


def read_ods_cell(cell: XML.Element) -> Cell:
    kind = cell.get(tag("office", "value-type"))
    if kind == "boolean":
        return cell.attrib[tag("office", "boolean-value")] == "true"
    if kind == "float":
        return float(cell.attrib[tag("office", "value")])
    if kind == "string":
        paragraph = cell.find("text:p", NS)
        text = paragraph.text or ""
        for child in paragraph:
            if child.tag == tag("text", "s"):
                text += " " * int(child.get(tag("text", "c"), "1"))
            else:
                text += {tag("text", "tab"): "\t", tag("text", "line-break"): "\n"}[child.tag]
            text += child.tail or ""
        return text
    assert kind is None
    return None


def verify_ods(path: Path, expected: dict[str, list[list[Cell]]]) -> None:
    with ZipFile(path) as archive:
        assert archive.testzip() is None
        first = archive.infolist()[0]
        assert first.filename == "mimetype" and first.compress_type == ZIP_STORED
        assert archive.read("mimetype").decode() == ODS_MEDIA
        manifest = XML.fromstring(archive.read("META-INF/manifest.xml"))
        assert any(
            entry.get(tag("manifest", "full-path")) == "/"
            and entry.get(tag("manifest", "media-type")) == ODS_MEDIA
            for entry in manifest
        )
        content = XML.fromstring(archive.read("content.xml"))
    tables = content.findall("office:body/office:spreadsheet/table:table", NS)
    assert [table.get(tag("table", "name")) for table in tables] == list(expected)
    for table, reference in zip(tables, expected.values()):
        rows = []
        for row in table.findall("table:table-row", NS):
            cells = []
            for cell in row:
                assert tag("table", "formula") not in cell.attrib
                cells.extend([read_ods_cell(cell)] * int(cell.get(tag("table", "number-columns-repeated"), "1")))
            cells.extend([None] * (len(reference[0]) - len(cells)))
            rows.append(cells)
        assert_cells(rows, reference)


def verify_delimited(path: Path, expected: list[list[Cell]]) -> None:
    assert path.read_bytes().startswith(b"\xef\xbb\xbf")
    with path.open(encoding="utf-8-sig", newline="") as stream:
        actual = list(csv.reader(stream, delimiter="," if path.suffix == ".csv" else "\t"))
    text_rows = []
    for row in expected:
        cells = []
        for value in row:
            if value is None:
                cells.append("")
            elif isinstance(value, bool):
                cells.append(str(value).upper())
            elif isinstance(value, float):
                cells.append(format(value, ".15g"))
            elif re.match(r"(?:[\t\r\n]|\s*[=+\-@＝＋－＠])", value):
                cells.append("'" + value)
            else:
                cells.append(value)
        text_rows.append(cells)
    assert_cells(actual, text_rows)


def verify_workbook(path: Path, expected: dict[str, list[list[Cell]]]) -> None:
    if path.suffix == ".ods":
        verify_ods(path, expected)
    else:
        verify_excel(path, expected)


def run_batch(output: Path, paths: list[Path]) -> list[str]:
    name_tests = {
        "logs.json": "logs",
        "LOGS.jsonl": "LOGS (2)",
        "History.json": "History (2)",
        "'[서울]'.json": "_서울_",
        "😀" * 17 + "a.json": "😀" * 15,
        "😀" * 17 + "b.json": "😀" * 13 + " (2)",
        "a&b.json": "a&b",
    }
    folder = output / "batch-name-inputs"
    folder.mkdir()
    expected_names = [path.stem for path in paths] + list(name_tests.values())
    for index, filename in enumerate(name_tests):
        path = folder / filename
        path.write_text(f'{{"marker": {index}}}\n', encoding="utf-8")
        paths.append(path)
    batch = read_batch(paths)
    assert [item.name for item in batch] == expected_names
    tables = {item.name: item.table for item in batch}
    expected = {item.name: expected_grid(item.table) for item in batch}
    lines = ["# Tablyfern batch workbook matrix", "", "| Format | Tabs verified | Result |", "| --- | ---: | --- |"]
    failures = []
    for extension in WORKBOOK_FORMATS:
        path = output / ("combined" + extension)
        try:
            save_workbook(tables, path)
            verify_workbook(path, expected)
        except (AssertionError, ConversionError, OSError, ValueError) as exc:
            failures.append(f"{path.name}: {type(exc).__name__} {exc}")
            lines.append(f"| {extension} | — | FAIL |")
        else:
            lines.append(f"| {extension} | {len(batch)} | PASS |")
    lines.extend(["", "## Files and resulting tabs", "", "| Input | Tab | Rows × columns |", "| --- | --- | --- |"])
    for item in batch:
        lines.append(f"| {item.path.name} | {item.name} | {len(item.table.rows)} × {len(item.table.headers)} |")
    if failures:
        lines.extend(["", "## Failures", "", *["- " + error for error in failures]])
    (output / "batch-matrix.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{len(batch)} input files per workbook · {len(WORKBOOK_FORMATS) - len(failures)} batch formats passed")
    return failures


def run(output: Path) -> int:
    cases = catalog()
    generate(output / "inputs", cases)
    exports = output / "exports"
    exports.mkdir()
    headings = "| Input | Parse | " + " | ".join(FORMATS) + " |"
    lines = ["# Tablyfern format matrix", "", headings, "|" + "---|" * (2 + len(FORMATS))]
    failures = []
    rejected = exported = 0
    for case in cases:
        source = output / "inputs" / case.filename
        try:
            _, table = read_table_file(source)
        except ConversionError as exc:
            good = exc.code == case.expected
            rejected += good
            if not good:
                failures.append(f"{case.filename}: unexpected input error {exc.code}")
            cells = [f"{'REJECTED' if good else 'FAIL'} ({exc.code})", *["·"] * len(FORMATS)]
        else:
            shape = len(table.rows), len(table.headers)
            if shape != case.expected:
                failures.append(f"{case.filename}: expected {case.expected}, got {shape}")
            cells = [f"{shape[0]} × {shape[1]}"]
            expected = expected_grid(table)
            for extension in FORMATS:
                path = exports / (source.stem + extension)
                try:
                    save_table(table, path)
                    if extension in WORKBOOK_FORMATS:
                        verify_workbook(path, {"Data": expected})
                    else:
                        verify_delimited(path, expected)
                except (AssertionError, ConversionError, OSError, ValueError) as exc:
                    failures.append(f"{path.name}: {type(exc).__name__} {exc}")
                    cells.append("FAIL")
                else:
                    exported += 1
                    cells.append("PASS")
        lines.append("| " + " | ".join([case.filename, *cells]) + " |")
    lines.extend(["", f"{len(cases)} inputs · {exported} verified exports · {rejected} expected input rejections"])
    if failures:
        lines.extend(["", "## Failures", "", *["- " + error for error in failures]])
    (output / "matrix.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(lines[-1] if not failures else "\n".join(failures))
    failures.extend(run_batch(output, [output / "inputs" / case.filename for case in cases if isinstance(case.expected, tuple)]))
    return int(bool(failures))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path, help="a new evidence directory")
    raise SystemExit(run(parser.parse_args().output))
