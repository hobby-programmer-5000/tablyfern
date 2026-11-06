"""Turn JSON or JSON Lines into a sparse worksheet."""

import json
import re
import unicodedata
from dataclasses import dataclass
from decimal import Decimal, DecimalException
from pathlib import Path

MAX_BYTES = 10 * 1024 * 1024
MAX_BATCH_BYTES = 20 * 1024 * 1024
MAX_BATCH_FILES = 100
MAX_ROWS = 1_048_576
MAX_COLUMNS = 16_384
MAX_TEXT = 32_767
MIN_EXCEL_NUMBER = Decimal("2.2251e-308")
MAX_EXCEL_NUMBER = Decimal("9.99999999999999e307")

Cell = str | float | bool | None


class ConversionError(ValueError):
    def __init__(self, code: str, **details: int | str):
        super().__init__(code)
        self.code = code
        self.details = details


@dataclass(frozen=True)
class JsonNumber:
    literal: str

    def to_cell(self) -> Cell:
        try:
            value = Decimal(self.literal)
            if len(value.as_tuple().digits) <= 15 and (
                not value or MIN_EXCEL_NUMBER <= value.copy_abs() <= MAX_EXCEL_NUMBER
            ):
                return float(value)
        except DecimalException:
            pass
        return self.literal


@dataclass(frozen=True)
class Table:
    headers: tuple[str, ...]
    rows: list[dict[int, Cell]]


@dataclass(frozen=True)
class BatchSheet:
    path: Path
    name: str
    source: str
    table: Table


class BatchInputError(ValueError):
    def __init__(self, path: Path, issue: ConversionError):
        super().__init__(issue.code)
        self.path = path
        self.issue = issue


def input_mode_for(path: Path) -> str:
    return {".json": "json", ".jsonl": "jsonl", ".ndjson": "jsonl"}.get(path.suffix.lower(), "auto")


def _sheet_name(stem: str, used: set[str]) -> str:
    base = unicodedata.normalize("NFC", stem)
    base = re.sub(r"[\[\]:*?/\\\x00-\x1f\ud800-\udfff\ufffe\uffff]", "_", base).strip(" '") or "Data"
    suffix = ""
    number = 1
    while True:
        limit = 2 * (31 - len(suffix))
        prefix = base.encode("utf-16-le")[:limit].decode("utf-16-le", errors="ignore")
        name = prefix.rstrip(" '") + suffix
        if name.casefold() not in used:
            used.add(name.casefold())
            return name
        number += 1
        suffix = f" ({number})"


def read_batch(paths: list[Path]) -> list[BatchSheet]:
    if len(paths) > MAX_BATCH_FILES:
        raise ConversionError("batch_count")
    used = {"history"}
    sheets = []
    total = 0
    for path in paths:
        try:
            source, size = _read_source(path)
            total += size
            if total > MAX_BATCH_BYTES:
                raise ConversionError("batch_size")
            table = parse_table(source, input_mode_for(path))
        except ConversionError as exc:
            raise BatchInputError(path, exc) from exc
        except OSError as exc:
            raise BatchInputError(path, ConversionError("read")) from exc
        sheets.append(BatchSheet(path, _sheet_name(path.stem, used), source, table))
    return sheets


def _object_from_pairs(pairs: list[tuple[str, object]]) -> dict[str, object]:
    result = {}
    for key, value in pairs:
        if key in result:
            raise ConversionError("duplicate")
        result[key] = value
    return result


def _reject_constant(_constant: str) -> None:
    raise ConversionError("constant")


def read_json_file(path: Path) -> str:
    return _read_source(path)[0]


def _read_source(path: Path) -> tuple[str, int]:
    with path.open("rb") as stream:
        data = stream.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise ConversionError("size")
    try:
        return data.decode("utf-8-sig"), len(data)
    except UnicodeDecodeError as exc:
        raise ConversionError("encoding") from exc


def _json_text(value: object) -> str:
    if isinstance(value, JsonNumber):
        return value.literal
    if isinstance(value, dict):
        fields = (f"{_json_text(key)}:{_json_text(item)}" for key, item in value.items())
        return "{" + ",".join(fields) + "}"
    if isinstance(value, list):
        return "[" + ",".join(map(_json_text, value)) + "]"
    return json.dumps(value, ensure_ascii=False)


def _flatten(record: dict, prefix: tuple[str, ...] = ()):
    for key, value in record.items():
        path = (*prefix, key)
        if isinstance(value, dict) and value:
            yield from _flatten(value, path)
        else:
            header = ".".join(
                part.replace("\\", "\\\\").replace(".", "\\.") if part else r"\0"
                for part in path
            )
            yield header, value


def _check_text(value: str, row: int, column: int) -> None:
    try:
        length = len(value.encode("utf-16-le")) // 2
    except UnicodeEncodeError as exc:
        raise ConversionError("unicode", row=row, column=column) from exc
    if length > MAX_TEXT or value.count("\n") > 253:
        raise ConversionError("cell_limit", row=row, column=column)


def _cell_value(value: object) -> Cell:
    if isinstance(value, JsonNumber):
        return value.to_cell()
    if isinstance(value, (list, dict)):
        return _json_text(value)
    return value


def _load(source: str) -> object:
    return json.loads(
        source, parse_int=JsonNumber, parse_float=JsonNumber,
        parse_constant=_reject_constant, object_pairs_hook=_object_from_pairs,
    )


def _load_lines(source: str) -> list:
    records = []
    for number, line in enumerate(source.removesuffix("\n").split("\n"), start=1):
        if number >= MAX_ROWS:
            raise ConversionError("rows")
        try:
            records.append(_load(line))
        except json.JSONDecodeError as exc:
            raise ConversionError("invalid", line=number, column=exc.colno) from exc
        except ConversionError as exc:
            raise ConversionError(exc.code + "_line", line=number) from exc
    return records


def parse_table(source: str, mode: str = "json") -> Table:
    try:
        if len(source.encode("utf-8")) > MAX_BYTES:
            raise ConversionError("size")
        source = source.removeprefix("\ufeff")
        if mode == "jsonl":
            return _to_table(_load_lines(source))
        try:
            root = _load(source)
        except json.JSONDecodeError as exc:
            if mode == "auto" and exc.msg == "Extra data" and "\n" in source:
                return _to_table(_load_lines(source))
            raise
        return _to_table(root if isinstance(root, list) else [root])
    except json.JSONDecodeError as exc:
        raise ConversionError("invalid", line=exc.lineno, column=exc.colno) from exc
    except UnicodeEncodeError as exc:
        raise ConversionError("encoding") from exc
    except RecursionError as exc:
        raise ConversionError("depth") from exc


def _to_table(records: list) -> Table:
    if len(records) >= MAX_ROWS:
        raise ConversionError("rows")

    column_by_header = {}
    rows = []
    object_records = bool(records) and all(isinstance(record, dict) for record in records)
    if not object_records:
        column_by_header["value"] = 0

    for row_number, record in enumerate(records, start=2):
        fields = _flatten(record) if object_records else (("value", record),)
        row = {}
        for header, value in fields:
            if header not in column_by_header:
                column = len(column_by_header)
                if column >= MAX_COLUMNS:
                    raise ConversionError("columns")
                _check_text(header, 1, column + 1)
                column_by_header[header] = column
            column = column_by_header[header]
            cell = _cell_value(value)
            if isinstance(cell, str):
                _check_text(cell, row_number, column + 1)
            if cell is not None:
                row[column] = cell
        rows.append(row)
    if not column_by_header:
        column_by_header["value"] = 0
    return Table(tuple(column_by_header), rows)
