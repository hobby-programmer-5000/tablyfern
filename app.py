"""Windows desktop interface for Tablyfern."""

import json
import sys
from itertools import islice
from pathlib import Path

from PySide6.QtCore import QThread, Qt, Slot
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from converter import (
    MAX_BATCH_FILES,
    MAX_BYTES,
    MAX_COLUMNS,
    MAX_LINE_FEEDS,
    MAX_RECORD_SIZE,
    MAX_ROWS,
    MAX_TEXT,
    BatchInputError,
    BatchSheet,
    ConversionError,
    Table,
    input_mode_for,
    parse_table,
    read_batch,
    read_table_file,
)
from exporters import FORMATS, MAX_DELIMITED_CELLS, WORKBOOK_FORMATS, save_table, save_workbook
from version import VERSION

PREVIEW_ROWS = 50
PREVIEW_COLUMNS = 40
EXAMPLE = """[
  {"name": "Mina", "active": true, "profile": {"city": "서울"}, "score": 92},
  {"name": "Jin", "active": false, "profile": {"city": "부산"}, "score": 87.5}
]"""

MESSAGES = {
    "tagline": ("Convert JSON and JSON Lines to spreadsheets on this device.", "이 기기에서 JSON과 JSONL을 스프레드시트로 변환합니다."),
    "language": ("Language", "언어"),
    "limits": ("Limits", "사용 한도"),
    "close": ("Close", "닫기"),
    "open": ("Open JSON…", "JSON 열기…"),
    "open_batch": ("Open multiple…", "여러 파일 열기…"),
    "sample": ("Load example", "예제 불러오기"),
    "preview": ("Preview table", "표 미리보기"),
    "export": ("Save file…", "파일 저장…"),
    "input": ("JSON input", "JSON 입력"),
    "table": ("Table preview", "표 미리보기"),
    "input_mode": ("Input format", "입력 형식"),
    "batch_picker": ("Choose a tab to preview", "미리 볼 탭 선택"),
    "auto": ("Detect input", "입력 자동 감지"),
    "output_format": ("Output format", "출력 형식"),
    "placeholder": ("Paste JSON or JSON Lines, or open a UTF-8 file.", "JSON 또는 JSONL을 붙여넣거나 UTF-8 파일을 여세요."),
    "ready": ("Open a JSON or JSONL file, or paste data to begin.", "JSON 또는 JSONL 파일을 열거나 데이터를 붙여넣으세요."),
    "edited": ("Preview the current JSON before saving.", "저장하기 전에 현재 JSON을 미리 확인하세요."),
    "summary": ("Rows {rows:,} · Columns {columns:,}", "{rows:,}행 · {columns:,}열"),
    "batch_summary": ("{count} tabs · {name} · Rows {rows:,} · Columns {columns:,}", "{count}개 탭 · {name} · {rows:,}행 · {columns:,}열"),
    "partial": (" · Preview shows {shown_rows:,} rows × {shown_columns:,} columns", " · 미리보기 {shown_rows:,}행 × {shown_columns:,}열 표시"),
    "saved": ("Saved {name}", "{name} 저장 완료"),
    "saved_batch": ("Saved {name} with {count} tabs", "{name} 저장 완료 · {count}개 탭"),
    "json_filter": ("JSON files (*.json *.jsonl *.ndjson);;All files (*)", "JSON 파일 (*.json *.jsonl *.ndjson);;모든 파일 (*)"),
    "export_filter": ("{kind} file (*{extension})", "{kind} 파일 (*{extension})"),
    "invalid": ("Invalid JSON at line {line}, column {column}. Check the syntax and retry.", "JSON 형식 오류 ({line}줄, {column}번째 문자). 문법을 확인하고 다시 시도하세요."),
    "duplicate": ("An object has a duplicate key. Give its fields unique names.", "객체에 중복 키가 있습니다. 각 필드에 서로 다른 이름을 사용하세요."),
    "duplicate_line": ("Duplicate object key on line {line}. Give its fields unique names.", "{line}줄에 중복 키가 있습니다. 각 필드에 서로 다른 이름을 사용하세요."),
    "constant": ("Use valid JSON numbers. NaN and Infinity are unsupported.", "올바른 JSON 숫자를 입력하세요. NaN과 Infinity는 지원되지 않습니다."),
    "constant_line": ("Invalid number on line {line}. Replace NaN or Infinity with a valid JSON value.", "{line}줄의 숫자가 잘못되었습니다. NaN 또는 Infinity를 올바른 JSON 값으로 바꾸세요."),
    "depth": ("The JSON is nested too deeply. Reduce nesting and retry.", "JSON 중첩이 너무 깊습니다. 중첩을 줄인 뒤 다시 시도하세요."),
    "size": ("Pasted input exceeds 10 MiB. Save it as a file and open it instead.", "붙여넣은 입력이 10 MiB를 초과했습니다. 파일로 저장한 뒤 열어 주세요."),
    "encoding": ("Use UTF-8 JSON. Save the input as UTF-8 and retry.", "UTF-8 JSON을 사용하세요. 입력을 UTF-8로 저장한 뒤 다시 시도하세요."),
    "unicode": ("Invalid Unicode in Excel row {row}, column {column}. Correct the source text.", "엑셀 {row}행 {column}열에 잘못된 유니코드가 있습니다. 원본 텍스트를 수정하세요."),
    "cell_limit": ("Text exceeds Excel's limit in row {row}, column {column}. Shorten the source value.", "{row}행 {column}열의 텍스트가 엑셀 셀 한도를 초과합니다. 원본 값을 줄이세요."),
    "rows": ("Too many records for one Excel sheet. Split the JSON and retry.", "엑셀 시트 한 장에 담을 수 있는 행 수를 초과했습니다. JSON을 나누어 다시 시도하세요."),
    "columns": ("Too many fields for one Excel sheet. Reduce the number of fields.", "엑셀 시트 한 장에 담을 수 있는 열 수를 초과했습니다. 필드 수를 줄이세요."),
    "read": ("Could not open the file. Check that it is available and readable.", "파일을 열 수 없습니다. 파일 위치와 읽기 권한을 확인하세요."),
    "save": ("Could not save the workbook. Check the destination or close the file in Excel and retry.", "통합 문서를 저장할 수 없습니다. 저장 위치를 확인하거나 엑셀에서 파일을 닫고 다시 시도하세요."),
    "extension": ("Use a filename ending in {extension}.", "파일 이름의 끝에 {extension}을 사용하세요."),
    "ods_character": ("ODS cannot store a character in {sheet}, row {row}, column {column}. Edit it or choose Excel format.", "ODS가 {sheet} 탭 {row}행 {column}열의 문자를 저장할 수 없습니다. 해당 문자를 편집하거나 Excel 형식을 선택하세요."),
    "delimited_limit": ("This table exceeds the CSV/TSV grid limit of 10 million cells. Choose a workbook format.", "CSV/TSV의 1,000만 셀 한도를 초과했습니다. 통합 문서 형식을 선택하세요."),
    "batch_failure": ("Could not load the selection. {name} · {reason}", "선택한 파일을 불러오지 못했습니다. {name} · {reason}"),
    "batch_count": ("Select at most 100 files at a time.", "한 번에 최대 100개 파일을 선택하세요."),
    "loading": ("Reading and preparing the file…", "파일을 읽고 표를 만드는 중…"),
    "saving": ("Writing the output file…", "출력 파일을 저장하는 중…"),
    "stream_notice": ("This file is processed on disk. Use the table preview below.", "이 파일은 디스크에서 처리합니다. 아래 표 미리보기에서 확인하세요."),
    "large_format": ("For large input, use a .json, .jsonl or .ndjson extension.", "큰 입력 파일에는 .json, .jsonl 또는 .ndjson 확장자를 사용하세요."),
    "record_size": ("Record {record} exceeds 16 Mi characters. Split that record before opening.", "{record}번째 레코드가 16 Mi자(16,777,216자)를 초과했습니다. 해당 레코드를 나눈 뒤 열어 주세요."),
    "invalid_record": ("Invalid JSON near record {record}. Check the source syntax.", "{record}번째 레코드 부근에 JSON 형식 오류가 있습니다. 원본 문법을 확인하세요."),
    "source_changed": ("The input changed while being read. Open it again.", "읽는 동안 입력 파일이 바뀌었습니다. 다시 열어 주세요."),
    "scratch": ("Could not write temporary data. Check available disk space.", "임시 데이터를 쓸 수 없습니다. 디스크 여유 공간을 확인하세요."),
    "workbook_size": ("The workbook exceeds the output package size limit. Export fewer records at a time.", "통합 문서가 출력 패키지 크기 한도를 초과했습니다. 레코드를 나눠 저장하세요."),
    "batch_empty": ("Select at least one input file.", "입력 파일을 하나 이상 선택하세요."),
    "batch_format": ("Choose a workbook format for multiple tabs.", "여러 탭을 저장할 수 있는 통합 문서 형식을 선택하세요."),
}

LIMITS = (
    """FILE INPUT
There is no fixed file-size cap for UTF-8 JSON, JSONL or NDJSON. Large files use temporary disk space. A 1 GiB file can fit if its records and output meet the limits below and both the temporary and destination drives have enough free space. A workbook package can hit a size limit even when its rows and columns fit.

PER RECORD AND WORKSHEET
Disk processing supports up to {record:,} characters per JSON array item or JSONL/NDJSON line. A single JSON object or scalar counts as one record.
Each input file maps to one worksheet with up to {rows:,} data rows plus the header and {columns:,} columns, including continuation columns.
Imported text beyond {text:,} UTF-16 units or {line_feeds} line feeds per cell continues in adjacent columns. Join the parts in order to recover the text.

OTHER LIMITS
Editor paste is limited to {paste} MiB. Preview shows up to {preview_rows} rows and {preview_columns} columns. A batch can contain up to {batch} files. CSV and TSV support one table and up to {delimited:,} cells, including the header.""",
    """파일 입력
UTF-8 JSON, JSONL, NDJSON에 고정 파일 크기 한도는 없습니다. 큰 파일은 디스크 임시 공간을 사용합니다. 1 GiB 파일도 아래 레코드·출력 한도에 맞고 임시 폴더와 저장 위치에 여유 공간이 있으면 처리할 수 있습니다. 행과 열이 한도 안이어도 통합 문서 패키지가 크기 한도에 도달할 수 있습니다.

레코드 및 시트
디스크 처리 시 JSON 배열 항목 하나 또는 JSONL·NDJSON 한 줄은 최대 {record:,}자입니다. 단일 JSON 객체나 값은 레코드 하나입니다.
파일 하나당 시트 한 장으로 저장합니다. 시트마다 머리글 외 데이터 {rows:,}행, 이어지는 열을 포함해 {columns:,}열까지 지원합니다.
가져온 텍스트가 셀당 UTF-16 {text:,}단위 또는 줄바꿈 {line_feeds}개를 넘으면 옆 열로 나눕니다. 열 순서대로 이어 붙이면 원래 텍스트를 복원할 수 있습니다.

기타 한도
붙여넣기는 {paste} MiB까지이며, 미리보기는 최대 {preview_rows}행과 {preview_columns}열을 보여줍니다. 한 번에 파일 {batch}개까지 열 수 있습니다. CSV와 TSV는 표 하나를 지원하며 머리글을 포함해 {delimited:,}셀까지 저장할 수 있습니다.""",
)


class FileJob(QThread):
    def __init__(self, action, on_success, fallback: str, parent):
        super().__init__(parent)
        self.action = action
        self.on_success = on_success
        self.fallback = fallback
        self.result = None
        self.error = None

    def run(self):
        try:
            self.result = self.action()
        except Exception as exc:
            self.error = exc


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tablyfern")
        self.setMinimumSize(640, 460)
        self.resize(960, 700)
        self.table: Table | None = None
        self.batch: list[BatchSheet] = []
        self.file_table: Table | None = None
        self.job: FileJob | None = None
        self.source_name = "tablyfern"
        self.status_key = "ready"
        self.status_values = {}
        self.status_error = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)
        heading = QHBoxLayout()
        title = QLabel("Tablyfern")
        font = title.font()
        font.setPointSize(24)
        font.setBold(True)
        title.setFont(font)
        heading.addWidget(title)
        heading.addStretch()
        self.limits_button = QPushButton()
        self.limits_button.clicked.connect(self.show_limits)
        heading.addWidget(self.limits_button)
        self.language = QComboBox()
        self.language.addItems(["English", "한국어"])
        heading.addWidget(self.language)
        layout.addLayout(heading)
        self.tagline = QLabel()
        layout.addWidget(self.tagline)

        actions = QHBoxLayout()
        self.buttons = {}
        self.output_format = QComboBox()
        self.set_output_formats(FORMATS)
        for key, action in (
            ("open", self.open_json),
            ("open_batch", self.open_batch),
            ("sample", self.load_example),
            ("preview", self.update_preview),
            ("export", self.export),
        ):
            if key == "export":
                actions.addStretch()
                actions.addWidget(self.output_format)
            button = QPushButton()
            button.clicked.connect(action)
            button.setMinimumHeight(32)
            self.buttons[key] = button
            actions.addWidget(button)
        layout.addLayout(actions)

        split = QSplitter(Qt.Orientation.Vertical)
        self.editor = QPlainTextEdit()
        self.editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.editor.textChanged.connect(self.invalidate_preview)
        self.input_label = QLabel()
        self.input_label.setBuddy(self.editor)
        self.input_mode = QComboBox()
        self.input_mode.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        for label, mode in (("", "auto"), ("JSON", "json"), ("JSONL", "jsonl")):
            self.input_mode.addItem(label, mode)
        self.input_mode.currentIndexChanged.connect(self.invalidate_preview)
        self.batch_picker = QComboBox()
        self.batch_picker.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.batch_picker.setMaximumWidth(340)
        self.batch_picker.currentIndexChanged.connect(self.update_preview)
        self.batch_picker.hide()
        split.addWidget(self.panel(self.input_label, self.editor, self.input_mode, self.batch_picker))

        self.grid = QTableWidget()
        self.grid.setAlternatingRowColors(True)
        self.grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.grid.horizontalHeader().setDefaultSectionSize(145)
        self.table_label = QLabel()
        self.table_label.setBuddy(self.grid)
        split.addWidget(self.panel(self.table_label, self.grid))
        split.setChildrenCollapsible(False)
        split.setSizes([220, 300])
        layout.addWidget(split, 1)

        self.status = QLabel()
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        self.language.currentIndexChanged.connect(self.translate)
        self.translate()

    @staticmethod
    def panel(label: QLabel, content: QWidget, *controls: QWidget) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        heading = QHBoxLayout()
        heading.addWidget(label)
        heading.addStretch()
        for control in controls:
            heading.addWidget(control)
        layout.addLayout(heading)
        layout.addWidget(content)
        return panel

    def text(self, key: str, **values) -> str:
        return MESSAGES[key][self.language.currentIndex()].format(**values)

    def translate(self) -> None:
        self.tagline.setText(self.text("tagline"))
        self.limits_button.setText(self.text("limits"))
        self.language.setAccessibleName(self.text("language"))
        self.input_label.setText(self.text("input"))
        self.input_mode.setItemText(0, self.text("auto"))
        self.input_mode.setAccessibleName(self.text("input_mode"))
        self.batch_picker.setAccessibleName(self.text("batch_picker"))
        self.output_format.setAccessibleName(self.text("output_format"))
        self.editor.setAccessibleName(self.text("input"))
        self.editor.setPlaceholderText(self.text("placeholder"))
        self.table_label.setText(self.text("table"))
        self.grid.setAccessibleName(self.text("table"))
        for key, button in self.buttons.items():
            button.setText(self.text(key))
        self.buttons["export"].setEnabled(self.table is not None)
        if self.file_table is not None or (
            self.batch and self.batch[self.batch_picker.currentIndex()].source is None
        ):
            self.editor.setPlainText(self.text("stream_notice"))
        self.show_status(self.status_key, error=self.status_error, **self.status_values)

    def show_limits(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(self.text("limits"))
        dialog.resize(550, 430)
        layout = QVBoxLayout(dialog)
        details = QPlainTextEdit()
        details.setReadOnly(True)
        details.setPlainText(LIMITS[self.language.currentIndex()].format(
            record=MAX_RECORD_SIZE, rows=MAX_ROWS - 1, columns=MAX_COLUMNS,
            text=MAX_TEXT, line_feeds=MAX_LINE_FEEDS, paste=MAX_BYTES // (1024 * 1024),
            preview_rows=PREVIEW_ROWS, preview_columns=PREVIEW_COLUMNS, batch=MAX_BATCH_FILES,
            delimited=MAX_DELIMITED_CELLS,
        ))
        details.setAccessibleName(self.text("limits"))
        layout.addWidget(details)
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.button(QDialogButtonBox.StandardButton.Close).setText(self.text("close"))
        close.rejected.connect(dialog.reject)
        layout.addWidget(close)
        dialog.exec()

    def show_status(self, key: str, *, error: bool = False, **values) -> None:
        self.status_key, self.status_values, self.status_error = key, values, error
        if key == "batch_failure":
            issue = values["issue"]
            message = self.text(key, name=values["name"], reason=self.text(issue.code, **issue.details))
        else:
            message = self.text(key, **values)
        if key in ("summary", "batch_summary"):
            if values["rows"] > PREVIEW_ROWS or values["columns"] > PREVIEW_COLUMNS:
                message += self.text("partial", shown_rows=self.grid.rowCount(), shown_columns=self.grid.columnCount())
        self.status.setText(message)
        font = self.status.font()
        font.setBold(error)
        self.status.setFont(font)

    def invalidate_preview(self) -> None:
        if self.batch or self.file_table is not None:
            return
        self.table = None
        self.grid.clear()
        self.grid.setRowCount(0)
        self.grid.setColumnCount(0)
        self.buttons["export"].setEnabled(False)
        self.show_status("edited")

    def open_json(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(self, self.text("open"), "", self.text("json_filter"))
        if not filename:
            return
        path = Path(filename)
        self.start_job(
            "loading", lambda: read_table_file(path),
            lambda result: self.show_open_file(path, *result), "read",
        )

    def show_open_file(self, path: Path, source: str | None, table: Table) -> None:
        self.set_batch([])
        self.source_name = path.stem
        self.input_mode.setCurrentIndex(self.input_mode.findData(input_mode_for(path)))
        if source is None:
            self.file_table = table
            self.editor.setReadOnly(True)
            self.input_mode.hide()
        self.editor.setPlainText(self.text("stream_notice") if source is None else source)
        self.show_table(table)

    def load_example(self) -> None:
        self.set_batch([])
        self.source_name = "tablyfern-example"
        self.input_mode.setCurrentIndex(0)
        self.editor.setPlainText(EXAMPLE)
        self.update_preview()

    def set_output_formats(self, formats: tuple[str, ...]) -> None:
        selected = self.output_format.currentData()
        self.output_format.clear()
        for extension in formats:
            self.output_format.addItem(extension.upper().removeprefix("."), extension)
        self.output_format.setCurrentIndex(max(0, self.output_format.findData(selected)))

    def set_batch(self, batch: list[BatchSheet]) -> None:
        self.file_table = None
        self.batch = batch
        self.batch_picker.blockSignals(True)
        self.batch_picker.clear()
        for item in batch:
            self.batch_picker.addItem(item.path.name)
        self.batch_picker.blockSignals(False)
        self.batch_picker.setVisible(bool(batch))
        self.input_mode.setVisible(not batch)
        self.editor.setReadOnly(bool(batch))
        self.set_output_formats(WORKBOOK_FORMATS if batch else FORMATS)

    def open_batch(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, self.text("open_batch"), "", self.text("json_filter"))
        if not files:
            return
        self.start_job(
            "loading", lambda: read_batch([Path(filename) for filename in files]),
            self.show_open_batch, "read",
        )

    def show_open_batch(self, batch: list[BatchSheet]) -> None:
        self.source_name = "tablyfern-workbook"
        self.set_batch(batch)
        self.update_preview()

    def update_preview(self) -> None:
        if self.batch:
            current = self.batch[self.batch_picker.currentIndex()]
            self.batch_picker.setToolTip(f"{current.path.name} → {current.name}")
            self.editor.setPlainText(self.text("stream_notice") if current.source is None else current.source)
            self.render_table(current.table)
            self.show_status(
                "batch_summary", count=len(self.batch), name=current.name,
                rows=len(current.table.rows), columns=len(current.table.headers),
            )
            return
        if self.file_table is not None:
            self.show_table(self.file_table)
            return
        self.invalidate_preview()
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            table = parse_table(self.editor.toPlainText(), self.input_mode.currentData())
        except ConversionError as exc:
            self.show_status(exc.code, error=True, **exc.details)
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.show_table(table)

    def show_table(self, table: Table) -> None:
        self.render_table(table)
        self.show_status("summary", rows=len(table.rows), columns=len(table.headers))

    def render_table(self, table: Table) -> None:
        self.grid.clear()
        headers = table.headers[:PREVIEW_COLUMNS]
        rows = list(islice(table.rows, PREVIEW_ROWS))
        self.grid.setColumnCount(len(headers))
        self.grid.setHorizontalHeaderLabels(headers)
        self.grid.setRowCount(len(rows))
        for row_number, row in enumerate(rows):
            for column, value in row.items():
                if column < PREVIEW_COLUMNS:
                    if isinstance(value, bool):
                        display = "TRUE" if value else "FALSE"
                    elif isinstance(value, float):
                        display = format(value, ".15g")
                    else:
                        display = str(value)
                    self.grid.setItem(row_number, column, QTableWidgetItem(display))
        self.table = table
        self.buttons["export"].setEnabled(True)

    def export(self) -> None:
        if self.table is None:
            return
        extension = self.output_format.currentData()
        kind = self.output_format.currentText()
        dialog = QFileDialog(self, self.text("export"), "", self.text("export_filter", kind=kind, extension=extension))
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        dialog.setDefaultSuffix(extension.removeprefix("."))
        dialog.selectFile(self.source_name + extension)
        if not dialog.exec():
            return
        destination = Path(dialog.selectedFiles()[0])
        if destination.suffix.lower() != extension:
            self.show_status("extension", error=True, extension=extension)
            return
        if self.batch:
            tables = {item.name: item.table for item in self.batch}
            self.start_job("saving", lambda: save_workbook(tables, destination),
                           lambda _: self.show_status("saved_batch", name=destination.name, count=len(tables)), "save")
        else:
            table = self.table
            self.start_job("saving", lambda: save_table(table, destination),
                           lambda _: self.show_status("saved", name=destination.name), "save")

    def start_job(self, status: str, action, on_success, fallback: str) -> None:
        job = FileJob(action, on_success, fallback, self)
        self.job = job
        self.show_status(status)
        self.setEnabled(False)
        job.finished.connect(self.finish_job)
        job.start()

    @Slot()
    def finish_job(self) -> None:
        job = self.job
        self.job = None
        self.setEnabled(True)
        if isinstance(job.error, BatchInputError):
            self.show_status("batch_failure", error=True, name=job.error.path.name, issue=job.error.issue)
        elif isinstance(job.error, ConversionError):
            self.show_status(job.error.code, error=True, **job.error.details)
        elif job.error is not None:
            self.show_status(job.fallback, error=True)
        else:
            job.on_success(job.result)
        job.deleteLater()

    def closeEvent(self, event) -> None:
        if self.job is not None:
            event.ignore()
        else:
            super().closeEvent(event)


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Tablyfern")
    app.setApplicationVersion(VERSION)
    window = Window()
    if len(sys.argv) == 3 and sys.argv[1] == "--smoke-check":
        return smoke_check(window, Path(sys.argv[2]))
    window.show()
    return app.exec()


def smoke_check(window: Window, output: Path) -> int:
    from xml.etree import ElementTree
    from zipfile import ZipFile

    if window.buttons["export"].isEnabled() or window.output_format.currentData() != ".xlsx":
        return 1
    root = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
    samples = root / "samples"
    batch = read_batch([samples / "orders.json", samples / "events.jsonl", samples / "empty.json"])
    window.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen)
    window.set_batch(batch)
    window.update_preview()
    window.show()
    QApplication.processEvents()
    if window.grid.item(0, 1).text() != "서울":
        return 1
    output.mkdir(parents=True, exist_ok=True)
    destination = output / "smoke.xlsx"
    save_workbook({item.name: item.table for item in batch}, destination)
    with ZipFile(destination) as archive:
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        sheets = [sheet.attrib["name"] for sheet in workbook.findall(".//{*}sheet")]
    if sheets != ["orders", "events", "empty"]:
        return 1
    (output / "smoke.json").write_text(
        json.dumps({"version": VERSION, "frozen": bool(getattr(sys, "frozen", False)), "tabs": sheets}) + "\n",
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
