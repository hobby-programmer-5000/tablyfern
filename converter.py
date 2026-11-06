"""Turn JSON or JSON Lines into a sparse worksheet."""

import json
import os
import re
import unicodedata
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from decimal import Decimal, DecimalException
from io import StringIO
from itertools import count
from pathlib import Path
from tempfile import TemporaryFile
from typing import TextIO

EDITOR_THRESHOLD = 10 * 1024 * 1024
READ_CHUNK = 256 * 1024
MAX_ROWS = 1_048_576
MAX_COLUMNS = 16_384
MAX_TEXT = 32_767
MAX_LINE_FEEDS = 253
MIN_EXCEL_NUMBER = Decimal("2.2251e-308")
MAX_EXCEL_NUMBER = Decimal("9.99999999999999e307")

Cell = str | float | bool | None
Row = dict[int, Cell]


class DiskRows:
    """Reusable, sparse rows backed by an automatically removed temporary file."""

    def __init__(self, file: TextIO, count: int, columns: list[int]):
        self.file = file
        self.count = count
        self.columns = columns

    def __len__(self) -> int:
        return self.count

    def __iter__(self) -> Iterator[Row]:
        self.file.seek(0)
        for line in self.file:
            yield {self.columns[column]: value for column, value in json.loads(line)}

    def __del__(self):
        self.file.close()


class ConversionError(ValueError):
    def __init__(self, code: str, **details: int | str):
        super().__init__(code)
        self.code = code
        self.details = details


class ExtraJsonData(ConversionError):
    def __init__(self):
        super().__init__("invalid_record", record=1)


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
    rows: list[Row] | DiskRows


@dataclass(frozen=True)
class BatchSheet:
    path: Path
    name: str
    source: str | None
    table: Table


class BatchInputError(ValueError):
    def __init__(self, path: Path, issue: ConversionError):
        super().__init__(issue.code)
        self.path = path
        self.issue = issue


def input_mode_for(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".jsonl", ".ndjson", ".jsonlines", ".ldjson"}:
        return "jsonl"
    return "json" if suffix == ".json" else "auto"


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
    used = {"history"}
    sheets = []
    for path in paths:
        try:
            source, table = read_table_file(path, editor_limit=1024 * 1024)
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
    try:
        with path.open("r", encoding="utf-8-sig", newline="\n") as stream:
            return stream.read()
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
    if length > MAX_TEXT or value.count("\n") > MAX_LINE_FEEDS:
        raise ConversionError("cell_limit", row=row, column=column)


def _cell_value(value: object) -> Cell:
    if isinstance(value, JsonNumber):
        return value.to_cell()
    if isinstance(value, (list, dict)):
        return _json_text(value)
    return value


DECODER = json.JSONDecoder(
    parse_int=JsonNumber, parse_float=JsonNumber,
    parse_constant=_reject_constant, object_pairs_hook=_object_from_pairs,
)
JSON_SPACE = re.compile(r"[ \t\n\r]*")
NUMBER_TAIL = re.compile(r"(?:\.|[eE][+-]?)[0-9]*\Z")


def _load(source: str) -> object:
    return DECODER.decode(source)


def _load_line(line: str, number: int) -> object:
    try:
        return _load(line)
    except json.JSONDecodeError as exc:
        raise ConversionError("invalid", line=number, column=exc.colno) from exc
    except ConversionError as exc:
        raise ConversionError(exc.code + "_line", line=number) from exc


def _load_lines(source: str) -> list:
    records = []
    for number, line in enumerate(source.removesuffix("\n").split("\n"), start=1):
        if number >= MAX_ROWS:
            raise ConversionError("rows")
        records.append(_load_line(line, number))
    return records


def parse_table(source: str, mode: str = "json") -> Table:
    try:
        source = source.removeprefix("\ufeff")
        if len(source) > EDITOR_THRESHOLD:
            with StringIO(source, newline="\n") as stream:
                return _stream_table(stream, mode)
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


class JsonRecords:
    """Decode one top-level array item at a time using the standard decoder."""

    def __init__(self, stream: TextIO):
        self.stream = stream
        self.buffer = ""
        self.pos = 0
        self.eof = False

    def fill(self, size: int) -> bool:
        self.buffer = self.buffer[self.pos:]
        self.pos = 0
        chunk = self.stream.read(size)
        self.buffer += chunk
        self.eof = not chunk
        return bool(chunk)

    def peek(self) -> str:
        while True:
            self.pos = JSON_SPACE.match(self.buffer, self.pos).end()
            if self.pos < len(self.buffer):
                return self.buffer[self.pos]
            if self.eof or not self.fill(READ_CHUNK):
                return ""

    def value(self, number: int) -> object:
        while True:
            try:
                value, end = DECODER.raw_decode(self.buffer, self.pos)
            except json.JSONDecodeError as exc:
                if self.eof or not self.needs_more(exc):
                    raise ConversionError("invalid_record", record=number) from exc
            else:
                next_char = self.buffer[end:end + 1]
                # A number ending at a chunk boundary may have more digits or an exponent.
                incomplete = not self.eof and isinstance(value, JsonNumber) and (
                    not next_char or NUMBER_TAIL.match(self.buffer, end) is not None
                )
                if not incomplete:
                    self.pos = end
                    return value
            # Geometric growth keeps retries linear in the size of one large record.
            self.fill(max(READ_CHUNK, len(self.buffer) - self.pos))

    def needs_more(self, error: json.JSONDecodeError) -> bool:
        if error.pos == len(self.buffer) or error.msg.startswith("Unterminated string"):
            return True
        if error.msg.startswith("Invalid \\u"):
            return len(self.buffer) - error.pos < 6
        if error.msg == "Expecting value" and len(self.buffer) - error.pos < 9:
            tail = self.buffer[error.pos:]
            return any(literal.startswith(tail) for literal in ("true", "false", "null", "NaN", "Infinity", "-Infinity"))
        return error.msg == "Expecting ',' delimiter" and NUMBER_TAIL.match(self.buffer, error.pos) is not None

    def __iter__(self) -> Iterator[object]:
        if self.peek() != "[":
            value = self.value(1)
            if self.peek():
                raise ExtraJsonData()
            yield value
            return
        self.pos += 1
        if self.peek() == "]":
            self.pos += 1
        else:
            for number in count(1):
                self.peek()
                yield self.value(number)
                delimiter = self.peek()
                if delimiter not in (",", "]"):
                    raise ConversionError("invalid_record", record=number)
                self.pos += 1
                if delimiter == "]":
                    break
        if self.peek():
            raise ExtraJsonData()


def _stream_lines(stream: TextIO) -> Iterator[object]:
    for number, line in enumerate(stream, start=1):
        if number >= MAX_ROWS:
            raise ConversionError("rows")
        yield _load_line(line, number)


def _source_stamp(stream: TextIO) -> tuple[int, int]:
    stat = os.fstat(stream.fileno())
    return stat.st_size, stat.st_mtime_ns


def _read_stream(stream: TextIO, read_records, check_source, *, detect_trailing: bool = False) -> Table:
    stream.seek(0)
    object_records = True
    records = 0
    for records, record in enumerate(read_records(stream), start=1):
        object_records = object_records and isinstance(record, dict)
        if not detect_trailing:
            if records >= MAX_ROWS:
                raise ConversionError("rows")
            if not object_records:
                break
    if records >= MAX_ROWS:
        raise ConversionError("rows")
    if check_source is not None:
        check_source()
    stream.seek(0)
    return _spool_table(read_records(stream), object_records)


def _stream_table(stream: TextIO, mode: str, check_source=None) -> Table:
    if mode == "jsonl":
        return _read_stream(stream, _stream_lines, check_source)
    try:
        return _read_stream(stream, JsonRecords, check_source, detect_trailing=mode == "auto")
    except ExtraJsonData:
        if mode != "auto":
            raise
    return _read_stream(stream, _stream_lines, check_source)


def read_table_file(path: Path, editor_limit: int = EDITOR_THRESHOLD) -> tuple[str | None, Table]:
    mode = input_mode_for(path)
    try:
        with path.open("r", encoding="utf-8-sig", newline="\n") as stream:
            stamp = _source_stamp(stream)

            def check_source():
                if _source_stamp(stream) != stamp:
                    raise ConversionError("source_changed")

            if stamp[0] <= editor_limit:
                source = stream.read()
                table = parse_table(source, mode)
            else:
                source = None
                table = _stream_table(stream, mode, check_source)
            check_source()
            return source, table
    except UnicodeDecodeError as exc:
        raise ConversionError("encoding") from exc
    except RecursionError as exc:
        raise ConversionError("depth") from exc


def _text_parts(value: Cell, row: int, column: int) -> Iterator[Cell]:
    if not isinstance(value, str):
        yield value
        return
    try:
        encoded = value.encode("utf-16-le")
    except UnicodeEncodeError as exc:
        raise ConversionError("unicode", row=row, column=column) from exc
    if len(encoded) <= 2 * MAX_TEXT and value.count("\n") <= MAX_LINE_FEEDS:
        yield value
        return
    # Cut at both Excel limits while keeping every character, including newlines.
    start = 0
    while start < len(encoded):
        end = min(start + 2 * MAX_TEXT, len(encoded))
        if 0xD8 <= encoded[end - 1] <= 0xDB:
            end -= 2
        text = encoded[start:end].decode("utf-16-le")
        if text.count("\n") > MAX_LINE_FEEDS:
            text = "\n".join(text.split("\n", MAX_LINE_FEEDS + 1)[:MAX_LINE_FEEDS]) + "\n"
            end = start + len(text.encode("utf-16-le"))
        yield text
        start = end


def _rows(
    records: Iterable[object], columns: dict[str, list[int]], object_records: bool,
) -> Iterator[Row]:
    column_count = 0
    for row_number, record in enumerate(records, start=2):
        if row_number > MAX_ROWS:
            raise ConversionError("rows")
        if object_records and not isinstance(record, dict):
            raise ConversionError("source_changed")
        fields = _flatten(record) if object_records else (("value", record),)
        row = {}
        for header, value in fields:
            if header not in columns:
                _check_text(header, 1, column_count + 1)
                columns[header] = []
            indices = columns[header]
            column = indices[0] if indices else column_count
            for index, cell in enumerate(_text_parts(_cell_value(value), row_number, column + 1)):
                if index == len(indices):
                    if column_count >= MAX_COLUMNS:
                        raise ConversionError("columns")
                    indices.append(column_count)
                    column_count += 1
                if cell is not None:
                    row[indices[index]] = cell
        yield row


def _spool_table(records: Iterable[object], object_records: bool) -> Table:
    try:
        file = TemporaryFile(mode="w+t", encoding="utf-8", newline="\n")
    except OSError as exc:
        raise ConversionError("scratch") from exc
    columns: dict[str, list[int]] = {}
    count = 0
    try:
        for count, row in enumerate(_rows(records, columns, object_records), start=1):
            try:
                file.write(json.dumps(list(row.items()), ensure_ascii=False, separators=(",", ":")) + "\n")
            except OSError as exc:
                raise ConversionError("scratch") from exc
        try:
            file.flush()
        except OSError as exc:
            raise ConversionError("scratch") from exc
        headers, remap = _continuation_headers(columns)
        return Table(headers, DiskRows(file, count, remap))
    except BaseException:
        file.close()
        raise


def _continuation_headers(columns: dict[str, list[int]]) -> tuple[tuple[str, ...], list[int]]:
    used = set(columns)
    headers = []
    remap = [0] * sum(map(len, columns.values()))
    for name, indices in columns.items():
        for part, original in enumerate(indices, start=1):
            header = name
            if part > 1:
                suffix = f" [part {part}]"
                while True:
                    prefix = name.encode("utf-16-le")[:2 * (MAX_TEXT - len(suffix))].decode("utf-16-le", "ignore")
                    header = prefix + suffix
                    if header not in used:
                        break
                    suffix += "_"
                used.add(header)
            remap[original] = len(headers)
            headers.append(header)
    return tuple(headers) or ("value",), remap


def _to_table(records: list) -> Table:
    object_records = bool(records) and all(isinstance(record, dict) for record in records)
    columns: dict[str, list[int]] = {}
    rows = list(_rows(records, columns, object_records))
    headers, remap = _continuation_headers(columns)
    for index, row in enumerate(rows):
        rows[index] = {remap[column]: value for column, value in row.items()}
    return Table(headers, rows)
