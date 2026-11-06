"""Write real spreadsheet packages or delimited text from a Tablyfern table."""

import csv
import os
import re
from pathlib import Path
from shutil import copyfileobj
from tempfile import NamedTemporaryFile, TemporaryFile
from xml.etree import ElementTree
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

from xlsxwriter import Workbook
from xlsxwriter.exceptions import XlsxWriterException

from converter import Cell, ConversionError, Table

OPEN_XML_TYPES = {
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    ".xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
WORKBOOK_FORMATS = (".xlsx", *OPEN_XML_TYPES, ".ods")
FORMATS = (*WORKBOOK_FORMATS, ".csv", ".tsv")
MAX_DELIMITED_CELLS = 10_000_000
CONTENT_TYPES = "[Content_Types].xml"
ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet"
ODF_NS = "urn:oasis:names:tc:opendocument:xmlns:"
ODS_MANIFEST = f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="{ODF_NS}manifest:1.0" manifest:version="1.3">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="{ODS_MIME}"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"""
ODS_START = f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{ODF_NS}office:1.0"
  xmlns:table="{ODF_NS}table:1.0" xmlns:text="{ODF_NS}text:1.0" office:version="1.3">
  <office:body><office:spreadsheet>"""
ODS_END = "</office:spreadsheet></office:body></office:document-content>"
ODS_WHITESPACE = re.compile(r" +|[\t\n\r]")
INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def _write_xlsx(tables: dict[str, Table], output, scratch: Path) -> None:
    with Workbook(output, {"constant_memory": True, "tmpdir": str(scratch)}) as workbook:
        header_style = workbook.add_format({"bold": True, "bg_color": "#DCEFE6"})
        blank = workbook.add_format()
        for name, table in tables.items():
            sheet = workbook.add_worksheet(name)
            sheet.set_column(0, len(table.headers) - 1, 24)
            sheet.freeze_panes(1, 0)
            sheet.autofilter(0, 0, len(table.rows), len(table.headers) - 1)
            for column, header in enumerate(table.headers):
                sheet.write_string(0, column, header, header_style)
            for row_number, row in enumerate(table.rows, start=1):
                for column, value in row.items():
                    if isinstance(value, str):
                        sheet.write_string(row_number, column, value)
                    else:
                        sheet.write(row_number, column, value)
                if not row:
                    sheet.write_blank(row_number, 0, None, blank)


def _write_open_xml_variant(tables: dict[str, Table], output: Path, extension: str) -> None:
    with TemporaryFile(dir=output.parent) as base:
        _write_xlsx(tables, base, output.parent)
        with ZipFile(base) as original, ZipFile(output, "w", ZIP_DEFLATED) as variant:
            manifest = ElementTree.fromstring(original.read(CONTENT_TYPES))
            workbook_part = next(part for part in manifest if part.get("PartName") == "/xl/workbook.xml")
            workbook_part.set("ContentType", OPEN_XML_TYPES[extension])
            variant.writestr(CONTENT_TYPES, ElementTree.tostring(manifest, encoding="utf-8"))
            for entry in original.namelist():
                if entry != CONTENT_TYPES:
                    with original.open(entry) as source, variant.open(entry, "w") as destination:
                        copyfileobj(source, destination)


def _delimited_cell(value: Cell) -> str:
    if isinstance(value, str):
        if value.startswith(("\t", "\r", "\n")) or value.lstrip().startswith(
            ("=", "+", "-", "@", "＝", "＋", "－", "＠")
        ):
            return "'" + value
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if value is None:
        return ""
    return format(value, ".15g")


def _write_delimited(table: Table, output: Path, extension: str) -> None:
    with output.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream, delimiter="," if extension == ".csv" else "\t", quoting=csv.QUOTE_ALL)
        writer.writerow(map(_delimited_cell, table.headers))
        for row in table.rows:
            writer.writerow(_delimited_cell(row.get(column)) for column in range(len(table.headers)))


def _ods_whitespace(match: re.Match) -> str:
    token = match.group()
    if token[0] == " ":
        return f'<text:s text:c="{len(token)}"/>'
    return {"\t": "<text:tab/>", "\n": "<text:line-break/>", "\r": "&#13;"}[token]


def _ods_cell(value: Cell, sheet: str, row: int, column: int) -> str:
    if isinstance(value, str):
        if INVALID_XML.search(value):
            raise ConversionError("ods_character", sheet=sheet, row=row, column=column)
        text = ODS_WHITESPACE.sub(_ods_whitespace, escape(value))
        return f'<table:table-cell office:value-type="string"><text:p>{text}</text:p></table:table-cell>'
    if isinstance(value, bool):
        return f'<table:table-cell office:value-type="boolean" office:boolean-value="{str(value).lower()}"/>'
    return f'<table:table-cell office:value-type="float" office:value="{format(value, ".15g")}"/>'


def _write_ods(tables: dict[str, Table], output: Path) -> None:
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("mimetype", ODS_MIME, compress_type=ZIP_STORED)
        archive.writestr("META-INF/manifest.xml", ODS_MANIFEST)
        with archive.open("content.xml", "w") as stream:
            stream.write(ODS_START.encode())
            for name, table in tables.items():
                stream.write(f"<table:table table:name={quoteattr(name)}>".encode())
                headers = "".join(_ods_cell(header, name, 1, i) for i, header in enumerate(table.headers, start=1))
                stream.write(f"<table:table-row>{headers}</table:table-row>".encode())
                for number, row in enumerate(table.rows, start=2):
                    stream.write(b"<table:table-row>")
                    previous = 0
                    for column, value in sorted(row.items()):
                        if column > previous:
                            gap = column - previous
                            stream.write(f'<table:table-cell table:number-columns-repeated="{gap}"/>'.encode())
                        stream.write(_ods_cell(value, name, number, column + 1).encode())
                        previous = column + 1
                    if not row:
                        stream.write(b"<table:table-cell/>")
                    stream.write(b"</table:table-row>")
                stream.write(b"</table:table>")
            stream.write(ODS_END.encode())


def save_table(table: Table, destination: Path) -> None:
    extension = destination.suffix.lower()
    if extension not in FORMATS:
        raise ConversionError("extension")
    if extension in (".csv", ".tsv") and (1 + len(table.rows)) * len(table.headers) > MAX_DELIMITED_CELLS:
        raise ConversionError("delimited_limit")
    _save({"Data": table}, destination, extension)


def save_workbook(tables: dict[str, Table], destination: Path) -> None:
    extension = destination.suffix.lower()
    if extension not in WORKBOOK_FORMATS:
        raise ConversionError("batch_format")
    if not tables:
        raise ConversionError("batch_empty")
    _save(tables, destination, extension)


def _save(tables: dict[str, Table], destination: Path, extension: str) -> None:
    """Replace the destination only after a complete export."""
    with NamedTemporaryFile(dir=destination.parent, suffix=extension, delete=False) as file:
        temporary = Path(file.name)
    try:
        if extension == ".xlsx":
            _write_xlsx(tables, temporary, temporary.parent)
        elif extension in OPEN_XML_TYPES:
            _write_open_xml_variant(tables, temporary, extension)
        elif extension == ".ods":
            _write_ods(tables, temporary)
        else:
            _write_delimited(tables["Data"], temporary, extension)
        os.replace(temporary, destination)
    except XlsxWriterException as exc:
        raise ConversionError("save") from exc
    finally:
        temporary.unlink(missing_ok=True)
