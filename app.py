"""Windows desktop interface for Tablyfern."""

import json
import sys
import threading
from itertools import islice
from pathlib import Path

from PySide6.QtCore import QSize, QThread, QTimer, Qt, Slot
from PySide6.QtGui import QFontDatabase, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from converter import (
    MAX_COLUMNS,
    MAX_LINE_FEEDS,
    MAX_ROWS,
    MAX_TEXT,
    BatchInputError,
    BatchSheet,
    Cancelled,
    ConversionError,
    Table,
    cancellation,
    exceeds_editor_limit,
    input_mode_for,
    parse_table,
    read_batch,
    read_table_file,
)
from exporters import FORMATS, WORKBOOK_FORMATS, save_table, save_workbook
from ui_theme import ICON_DIR, LIGHT_STYLE, stock_icon, use_light_style
from version import VERSION

PREVIEW_ROWS = 50
PREVIEW_COLUMNS = 40
EXAMPLE = """[
  {"name": "Mina", "active": true, "profile": {"city": "서울"}, "score": 92},
  {"name": "Jin", "active": false, "profile": {"city": "부산"}, "score": 87.5}
]"""

MESSAGES = {
    "tagline": ("JSON & JSON Lines to spreadsheets", "JSON과 JSONL을 스프레드시트로 변환"),
    "language": ("Language", "언어"),
    "limits": ("Limits", "사용 한도"),
    "close": ("Close", "닫기"),
    "open": ("Open JSON…", "JSON 열기…"),
    "open_batch": ("Open multiple…", "여러 파일 열기…"),
    "sample": ("Load example", "예제 불러오기"),
    "preview": ("Update preview", "미리보기 갱신"),
    "export": ("Save file…", "파일 저장…"),
    "input": ("JSON input", "JSON 입력"),
    "table": ("Table preview", "표 미리보기"),
    "input_mode": ("Input format", "입력 형식"),
    "batch_picker": ("Workbook tab", "통합 문서 탭"),
    "auto": ("Detect input", "입력 자동 감지"),
    "output_format": ("Output format", "출력 형식"),
    "placeholder": ("Paste or type JSON or JSON Lines here.", "여기에 JSON 또는 JSONL을 붙여넣거나 입력하세요."),
    "empty_preview": ("Open a file or paste JSON to see its table here.", "파일을 열거나 JSON을 붙여넣으면 여기에 표가 표시됩니다."),
    "ready": ("Ready", "준비"),
    "edited": ("Input changed · Preview required to save", "입력 변경됨 · 저장하려면 미리보기 갱신 필요"),
    "pending_preview": ("Update the preview to see the table.", "미리보기를 갱신하면 표가 표시됩니다."),
    "error_status": ("Error · {message}", "오류 · {message}"),
    "summary": ("Rows {rows:,} · Columns {columns:,}", "{rows:,}행 · {columns:,}열"),
    "batch_summary": ("{count} tabs · {name} · Rows {rows:,} · Columns {columns:,}", "{count}개 탭 · {name} · {rows:,}행 · {columns:,}열"),
    "partial": (" · Preview shows {shown_rows:,} rows × {shown_columns:,} columns", " · 미리보기 {shown_rows:,}행 × {shown_columns:,}열 표시"),
    "saved": ("Saved {name}", "{name} 저장 완료"),
    "saved_batch": ("Saved {name} with {count} tabs", "{name} 저장 완료 · {count}개 탭"),
    "json_filter": (
        "JSON files (*.json *.jsonl *.ndjson *.jsonlines *.ldjson);;All files (*)",
        "JSON 파일 (*.json *.jsonl *.ndjson *.jsonlines *.ldjson);;모든 파일 (*)",
    ),
    "export_filter": ("{kind} file (*{extension})", "{kind} 파일 (*{extension})"),
    "invalid": ("Invalid JSON at line {line}, column {column}. Check the syntax and retry.", "JSON 형식 오류 ({line}줄, {column}번째 문자). 문법을 확인하고 다시 시도하세요."),
    "duplicate": ("An object has a duplicate key. Give its fields unique names.", "객체에 중복 키가 있습니다. 각 필드에 서로 다른 이름을 사용하세요."),
    "duplicate_line": ("Duplicate object key on line {line}. Give its fields unique names.", "{line}줄에 중복 키가 있습니다. 각 필드에 서로 다른 이름을 사용하세요."),
    "constant": ("Use valid JSON numbers. NaN and Infinity are unsupported.", "올바른 JSON 숫자를 입력하세요. NaN과 Infinity는 지원되지 않습니다."),
    "constant_line": ("Invalid number on line {line}. Replace NaN or Infinity with a valid JSON value.", "{line}줄의 숫자가 잘못되었습니다. NaN 또는 Infinity를 올바른 JSON 값으로 바꾸세요."),
    "depth": ("The JSON is nested too deeply. Reduce nesting and retry.", "JSON 중첩이 너무 깊습니다. 중첩을 줄인 뒤 다시 시도하세요."),
    "encoding": ("Use UTF-8 JSON. Save the input as UTF-8 and retry.", "UTF-8 JSON을 사용하세요. 입력을 UTF-8로 저장한 뒤 다시 시도하세요."),
    "unicode": ("Invalid Unicode in Excel row {row}, column {column}. Correct the source text.", "엑셀 {row}행 {column}열에 잘못된 유니코드가 있습니다. 원본 텍스트를 수정하세요."),
    "cell_limit": ("Text exceeds Excel's limit in row {row}, column {column}. Shorten the source value.", "{row}행 {column}열의 텍스트가 엑셀 셀 한도를 초과합니다. 원본 값을 줄이세요."),
    "rows": ("Too many records for one Excel sheet. Split the JSON and retry.", "엑셀 시트 한 장에 담을 수 있는 행 수를 초과했습니다. JSON을 나누어 다시 시도하세요."),
    "columns": ("Too many fields for one Excel sheet. Reduce the number of fields.", "엑셀 시트 한 장에 담을 수 있는 열 수를 초과했습니다. 필드 수를 줄이세요."),
    "read": ("Could not open the file. Check that it is available and readable.", "파일을 열 수 없습니다. 파일 위치와 읽기 권한을 확인하세요."),
    "save": ("Could not save the workbook. Check the destination or close the file in Excel and retry.", "통합 문서를 저장할 수 없습니다. 저장 위치를 확인하거나 엑셀에서 파일을 닫고 다시 시도하세요."),
    "extension": ("Use a filename ending in {extension}.", "파일 이름의 끝에 {extension}을 사용하세요."),
    "ods_character": ("ODS cannot store a character in {sheet}, row {row}, column {column}. Edit it or choose Excel format.", "ODS가 {sheet} 탭 {row}행 {column}열의 문자를 저장할 수 없습니다. 해당 문자를 편집하거나 Excel 형식을 선택하세요."),
    "batch_failure": ("Could not load the selection. {name} · {reason}", "선택한 파일을 불러오지 못했습니다. {name} · {reason}"),
    "loading": ("Reading and preparing the file…", "파일을 읽고 표를 만드는 중…"),
    "preparing": ("Preparing the preview…", "미리보기를 만드는 중…"),
    "saving": ("Writing the output file…", "출력 파일을 저장하는 중…"),
    "stream_notice": ("This file is processed on disk. View the result in Table preview.", "이 파일은 디스크에서 처리합니다. 결과는 표 미리보기에서 확인하세요."),
    "invalid_record": ("Invalid JSON near record {record}. Check the source syntax.", "{record}번째 레코드 부근에 JSON 형식 오류가 있습니다. 원본 문법을 확인하세요."),
    "source_changed": ("The input changed while being read. Open it again.", "읽는 동안 입력 파일이 바뀌었습니다. 다시 열어 주세요."),
    "scratch": ("Could not write temporary data. Check available disk space.", "임시 데이터를 쓸 수 없습니다. 디스크 여유 공간을 확인하세요."),
    "workbook_size": ("The workbook exceeds the output package size limit. Export fewer records at a time.", "통합 문서가 출력 패키지 크기 한도를 초과했습니다. 레코드를 나눠 저장하세요."),
    "batch_empty": ("Select at least one input file.", "입력 파일을 하나 이상 선택하세요."),
    "batch_format": ("Choose a workbook format for multiple tabs.", "여러 탭을 저장할 수 있는 통합 문서 형식을 선택하세요."),
    "locked": (
        "Could not write {name}. Close it in Excel or another program, or check that the folder is writable, then retry.",
        "{name}에 쓸 수 없습니다. 엑셀 등 다른 프로그램에서 파일을 닫거나 폴더 쓰기 권한을 확인한 뒤 다시 시도하세요.",
    ),
    "cancel": ("Cancel", "취소"),
    "cancelling": ("Cancelling… Temporary files are being removed.", "취소하는 중… 임시 파일을 정리하고 있습니다."),
    "cancelled": ("Cancelled. No file was changed.", "취소했습니다. 변경된 파일은 없습니다."),
    "quit_title": ("Task in progress", "작업 진행 중"),
    "quit_prompt": (
        "A file is still being read or saved. Cancel it and quit? Files being saved are left unchanged.",
        "파일을 읽거나 저장하는 중입니다. 작업을 취소하고 종료할까요? 저장 중이던 파일은 바뀌지 않습니다.",
    ),
    "quit_confirm": ("Cancel and quit", "취소하고 종료"),
    "keep_working": ("Keep working", "계속 진행"),
}

LIMITS = (
    """INPUT AND RESOURCES
Tablyfern sets no fixed size cap on files, pasted text, individual records or the number of files in a batch. Large files use temporary disk space. A single record still needs enough memory to be parsed, and pasted text stays in the editor. Unknown file extensions use automatic JSON/JSON Lines detection.

WORKSHEET AND OUTPUT
Each input file maps to one worksheet with up to {rows:,} data rows plus the header and {columns:,} columns, including continuation columns.
Imported text beyond {text:,} UTF-16 units or {line_feeds} line feeds per cell continues in adjacent columns. Join the parts in order to recover the text.
Preview shows up to {preview_rows} rows and {preview_columns} columns. Free space on both the temporary and output drives and the workbook package size also affect export.

FORMAT NOTES
CSV and TSV save one table with the same row and column bounds. Excel may infer cell types when importing text. Use XLSX for long IDs and leading zeros.
Excel can show an ODS used range wider than its actual data. For XLS or XLSB, open an exported XLSX in Excel and choose Save As.""",
    """입력 및 필요한 공간
파일 크기, 붙여넣기 크기, 레코드 크기, 한 번에 여는 파일 수에 앱이 정한 고정 한도는 없습니다. 큰 파일은 디스크 임시 공간을 사용합니다. 레코드 하나를 읽을 메모리가 필요하며, 붙여넣은 텍스트는 편집기에 남습니다. 알 수 없는 확장자는 JSON과 JSONL을 자동 감지합니다.

시트 및 출력
파일 하나당 시트 한 장으로 저장합니다. 시트마다 머리글 외 데이터 {rows:,}행, 이어지는 열을 포함해 {columns:,}열까지 지원합니다.
가져온 텍스트가 셀당 UTF-16 {text:,}단위 또는 줄바꿈 {line_feeds}개를 넘으면 옆 열로 나눕니다. 열 순서대로 이어 붙이면 원래 텍스트를 복원할 수 있습니다.
미리보기는 최대 {preview_rows}행과 {preview_columns}열을 보여줍니다. 임시 폴더와 저장 위치의 여유 공간, 통합 문서 패키지 크기도 저장에 영향을 줍니다.

형식별 참고 사항
CSV와 TSV는 같은 행·열 한도 안에서 표 하나를 저장합니다. Excel이 가져올 때 셀 유형을 추정할 수 있습니다. 긴 ID와 앞자리 0은 XLSX에서 보존할 수 있습니다.
Excel이 ODS의 사용 범위를 실제 데이터보다 넓게 표시할 수 있습니다. XLS 또는 XLSB는 내보낸 XLSX를 Excel에서 열어 다른 이름으로 저장하세요.""",
)


class FileJob(QThread):
    def __init__(self, action, on_success, fallback: str, parent):
        super().__init__(parent)
        self.action = action
        self.on_success = on_success
        self.fallback = fallback
        self.result = None
        self.error = None
        self.cancel_event = threading.Event()

    def cancel(self) -> None:
        self.cancel_event.set()

    def run(self):
        try:
            with cancellation(self.cancel_event):
                self.result = self.action()
        except Exception as exc:
            self.error = exc


def tables_in(result) -> list[Table]:
    """Tables produced by a job result, so a discarded result can be released."""
    if isinstance(result, Table):
        return [result]
    if isinstance(result, tuple):
        return [item for item in result if isinstance(item, Table)]
    if isinstance(result, list):
        return [item.table for item in result if isinstance(item, BatchSheet)]
    return []


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("workspace")
        self.setWindowTitle("Tablyfern")
        body_font = self.font()
        body_font.setFamilies(["Segoe UI Variable Text", "Segoe UI", body_font.family()])
        body_font.setPointSizeF(max(10.5, body_font.pointSizeF()))
        self.setFont(body_font)
        self.setMinimumSize(640, 520)
        self.resize(1120, 720)
        self.table: Table | None = None
        self.batch: list[BatchSheet] = []
        self.file_table: Table | None = None
        self.job: FileJob | None = None
        self.quit_requested = False
        self.held_tables: dict[int, Table] = {}
        self.source_name = "tablyfern"
        self.status_key = "ready"
        self.status_values = {}
        self.status_error = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 20, 24, 20)
        outer.setSpacing(16)
        self.content = QWidget()
        layout = QVBoxLayout(self.content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(16)
        outer.addWidget(self.content, 1)
        heading = QHBoxLayout()
        heading.setSpacing(8)
        self.brand = QLabel()
        self.brand.setObjectName("brandMark")
        self.brand.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.brand.setFixedSize(32, 40)
        heading.addWidget(self.brand)
        branding = QVBoxLayout()
        branding.setSpacing(0)
        title = QLabel("Tablyfern")
        title.setObjectName("appTitle")
        self.emphasize(title, 18)
        branding.addWidget(title)
        self.tagline = QLabel()
        self.tagline.setProperty("muted", True)
        branding.addWidget(self.tagline)
        heading.addLayout(branding)
        heading.addStretch()
        self.limits_button = QPushButton()
        self.limits_button.setProperty("quiet", True)
        self.style_button(self.limits_button)
        self.limits_button.clicked.connect(self.show_limits)
        heading.addWidget(self.limits_button)
        self.language = QComboBox()
        self.style_combo(self.language)
        self.language.addItems(["English", "한국어"])
        heading.addWidget(self.language)
        layout.addLayout(heading)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.buttons = {}
        self.output_format = QComboBox()
        self.style_combo(self.output_format)
        self.set_output_formats(FORMATS)
        for key, action in (
            ("open", self.open_json),
            ("open_batch", self.open_batch),
            ("sample", self.load_example),
            ("preview", self.update_preview),
            ("export", self.export),
        ):
            button = QPushButton()
            self.style_button(button)
            button.clicked.connect(action)
            self.buttons[key] = button
        self.buttons["open"].setObjectName("openButton")
        self.buttons["export"].setObjectName("exportButton")
        self.buttons["sample"].setProperty("quiet", True)
        for key in ("open", "open_batch", "sample"):
            button = self.buttons[key]
            actions.addWidget(button)
        actions.addStretch()
        layout.addLayout(actions)

        self.split = QSplitter(Qt.Orientation.Horizontal)
        self.split.setHandleWidth(12)
        self.split.setChildrenCollapsible(False)
        self.editor = QPlainTextEdit()
        self.editor.setObjectName("sourceEditor")
        editor_font = QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont)
        editor_font.setFamilies(["Consolas", editor_font.family()])
        editor_font.setPointSizeF(body_font.pointSizeF())
        self.editor.setFont(editor_font)
        self.editor.setMinimumHeight(64)
        self.editor.setTabChangesFocus(True)
        self.editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.editor.textChanged.connect(self.invalidate_preview)
        self.input_label = QLabel()
        self.input_label.setBuddy(self.editor)
        self.input_mode = QComboBox()
        self.style_combo(self.input_mode)
        self.input_mode.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        for label, mode in (("", "auto"), ("JSON", "json"), ("JSONL", "jsonl")):
            self.input_mode.addItem(label, mode)
        self.input_mode.currentIndexChanged.connect(self.invalidate_preview)
        self.batch_picker = QComboBox()
        self.style_combo(self.batch_picker)
        self.batch_picker.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.batch_picker.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.batch_picker.setMinimumContentsLength(8)
        self.batch_picker.currentIndexChanged.connect(self.update_preview)
        self.split.addWidget(self.panel(self.input_label, self.editor, self.input_mode))

        self.grid = QTableWidget()
        self.grid.setObjectName("previewGrid")
        self.grid.setAlternatingRowColors(True)
        self.grid.setShowGrid(False)
        self.grid.setTabKeyNavigation(False)
        self.grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.grid.verticalHeader().setDefaultSectionSize(36)
        self.grid.horizontalHeader().setDefaultSectionSize(128)
        self.grid.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.table_label = QLabel()
        self.table_label.setBuddy(self.grid)

        self.empty_preview = QLabel()
        self.empty_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_preview.setWordWrap(True)
        self.empty_preview.setMargin(12)
        self.empty_preview.setProperty("muted", True)
        self.preview_stack = QStackedWidget()
        self.preview_stack.setObjectName("previewStack")
        self.preview_stack.addWidget(self.empty_preview)
        self.preview_stack.addWidget(self.grid)
        self.preview_stack.setMinimumHeight(86)
        self.batch_label = QLabel()
        self.batch_label.setBuddy(self.batch_picker)
        self.sheet_selector = QWidget()
        sheets = QHBoxLayout(self.sheet_selector)
        sheets.setContentsMargins(0, 0, 0, 0)
        sheets.addWidget(self.batch_label)
        sheets.addWidget(self.batch_picker)
        self.sheet_selector.hide()
        self.split.addWidget(self.panel(
            self.table_label, self.preview_stack, self.buttons["preview"], self.sheet_selector,
        ))
        self.split.setStretchFactor(0, 1)
        self.split.setStretchFactor(1, 1)
        self.split.setSizes([380, 620])
        layout.addWidget(self.split, 1)

        footer = QHBoxLayout()
        footer.setSpacing(12)
        self.status = QLabel()
        self.status.setObjectName("statusLine")
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setWordWrap(True)
        self.status.setMinimumWidth(160)
        self.status.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        footer.addWidget(self.status, 1)
        self.cancel_button = QPushButton()
        self.style_button(self.cancel_button)
        self.cancel_button.clicked.connect(self.cancel_job)
        self.cancel_button.hide()
        footer.addWidget(self.cancel_button)
        self.output_label = QLabel()
        self.output_label.setObjectName("outputLabel")
        self.output_label.setBuddy(self.output_format)
        footer.addWidget(self.output_label)
        footer.addWidget(self.output_format)
        footer.addWidget(self.buttons["export"])
        outer.addLayout(footer)
        self.busy = QProgressBar()
        self.busy.setRange(0, 0)
        self.busy.setTextVisible(False)
        self.busy.setFixedHeight(4)
        self.busy.hide()
        outer.addWidget(self.busy)

        self.shortcuts = {}
        for key, sequence in (
            ("open", QKeySequence.StandardKey.Open),
            ("open_batch", "Ctrl+Shift+O"),
            ("preview", "Ctrl+Return"),
            ("export", QKeySequence.StandardKey.Save),
        ):
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.activated.connect(self.buttons[key].click)
            self.shortcuts[key] = shortcut
        focus_order = (
            self.limits_button, self.language, self.buttons["open"], self.buttons["open_batch"],
            self.buttons["sample"], self.input_mode, self.editor, self.buttons["preview"],
            self.batch_picker, self.grid, self.output_format, self.buttons["export"],
        )
        for previous, following in zip(focus_order, focus_order[1:]):
            self.setTabOrder(previous, following)
        hints = QApplication.styleHints()
        hints.colorSchemeChanged.connect(self.apply_theme)
        hints.accessibility().contrastPreferenceChanged.connect(self.apply_theme)
        self.language.currentIndexChanged.connect(self.translate)
        self.apply_theme()
        self.translate()
        self.editor.setFocus()

    @staticmethod
    def emphasize(widget: QWidget, size: int | None = None) -> None:
        font = widget.font()
        font.setBold(True)
        if size is not None:
            font.setPointSize(size)
        widget.setFont(font)

    @staticmethod
    def style_button(button: QPushButton) -> None:
        button.setProperty("workspaceButton", True)
        button.setMinimumHeight(36)
        button.setIconSize(QSize(18, 18))

    @staticmethod
    def style_combo(combo: QComboBox) -> None:
        combo.setProperty("workspaceCombo", True)
        combo.setMinimumHeight(36)

    @classmethod
    def panel(cls, label: QLabel, content: QWidget, *controls: QWidget) -> QFrame:
        panel = QFrame()
        panel.setProperty("panel", True)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(8)
        heading = QHBoxLayout()
        label.setProperty("panelTitle", True)
        label.setMinimumHeight(36)
        cls.emphasize(label)
        heading.addWidget(label)
        heading.addStretch()
        for control in controls:
            heading.addWidget(control)
        layout.addLayout(heading)
        layout.addWidget(content, 1)
        return panel

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        compact = event.size().width() < 790
        self.tagline.setVisible(not compact)
        margin = 16 if compact else 24
        self.layout().setContentsMargins(margin, 20, margin, 20)
        self.layout().setSpacing(12 if compact else 16)
        self.content.layout().setSpacing(12 if compact else 16)
        orientation = Qt.Orientation.Vertical if compact else Qt.Orientation.Horizontal
        if self.split.orientation() != orientation:
            self.split.setOrientation(orientation)
            extent = event.size().height() if orientation == Qt.Orientation.Vertical else event.size().width()
            self.split.setSizes([extent * 38 // 100, extent * 62 // 100])

    @Slot()
    def apply_theme(self) -> None:
        light = use_light_style()
        self.setStyleSheet(LIGHT_STYLE.replace("${arrow}", (ICON_DIR / "chevron-down.svg").as_posix()) if light else "")
        self.brand.setPixmap(stock_icon("sheet", light=light).pixmap(26, 26))
        self.grid.setShowGrid(not light)
        for key, icon in (("open", "folder-open"), ("open_batch", "files"),
                          ("preview", "refresh-cw"), ("export", "save")):
            self.buttons[key].setIcon(stock_icon(icon, light=light, inverse=key == "export"))

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
        self.batch_label.setText(self.text("batch_picker"))
        self.output_label.setText(self.text("output_format"))
        self.output_format.setAccessibleName(self.text("output_format"))
        self.editor.setAccessibleName(self.text("input"))
        self.editor.setPlaceholderText(self.text("placeholder"))
        self.table_label.setText(self.text("table"))
        self.grid.setAccessibleName(self.text("table"))
        for key, button in self.buttons.items():
            button.setText(self.text(key))
            if key in self.shortcuts:
                sequence = self.shortcuts[key].key().toString(QKeySequence.SequenceFormat.NativeText)
                button.setToolTip(f"{button.text()} ({sequence})")
        self.cancel_button.setText(self.text("cancel"))
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
            rows=MAX_ROWS - 1, columns=MAX_COLUMNS, text=MAX_TEXT, line_feeds=MAX_LINE_FEEDS,
            preview_rows=PREVIEW_ROWS, preview_columns=PREVIEW_COLUMNS,
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
        if error:
            message = self.text("error_status", message=message)
        if self.status.property("error") != error:
            self.status.setProperty("error", error)
            self.status.style().unpolish(self.status)
            self.status.style().polish(self.status)
            font = self.status.font()
            font.setBold(error)
            self.status.setFont(font)
        self.status.setText(message)
        self.sync_controls()

    def sync_controls(self) -> None:
        idle = self.job is None
        has_table = self.table is not None
        preview_needed = not has_table and not self.editor.isReadOnly() and not self.editor.document().isEmpty()
        self.buttons["preview"].setVisible(preview_needed)
        self.buttons["preview"].setEnabled(idle)
        self.buttons["export"].setEnabled(idle and has_table)
        for key in ("open", "open_batch", "sample"):
            self.buttons[key].setEnabled(idle)
        for control in (self.editor, self.input_mode, self.batch_picker, self.output_format, self.grid):
            control.setEnabled(idle)
        for key in self.shortcuts:
            self.shortcuts[key].setEnabled(self.buttons[key].isEnabled() and not self.buttons[key].isHidden())
        self.preview_stack.setCurrentWidget(self.grid if has_table else self.empty_preview)
        empty_key = self.status_key if not idle else "pending_preview" if preview_needed else "empty_preview"
        self.empty_preview.setText(self.text(empty_key))
        self.busy.setAccessibleName(self.status.text() if not idle else "")
        self.busy.setVisible(not idle)

    def invalidate_preview(self) -> None:
        if self.batch or self.file_table is not None:
            return
        if self.table is not None:
            self.table = None
            self.release_unused_tables()
            self.grid.clear()
            self.grid.setRowCount(0)
            self.grid.setColumnCount(0)
        key = "ready" if self.editor.document().isEmpty() else "edited"
        if self.status_key != key:
            self.show_status(key)

    def release_unused_tables(self) -> None:
        """Close the temporary files of tables the window no longer shows or exports."""
        live = [self.table, self.file_table, *(item.table for item in self.batch)]
        live = {id(table): table for table in live if table is not None}
        for key, table in self.held_tables.items():
            if key not in live:
                table.close()
        self.held_tables = live

    def release_all_tables(self) -> None:
        for table in self.held_tables.values():
            table.close()
        self.held_tables = {}

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
            self.batch_picker.addItem(item.name)
        self.batch_picker.blockSignals(False)
        self.sheet_selector.setVisible(bool(batch))
        self.input_mode.setVisible(not batch)
        self.editor.setReadOnly(bool(batch))
        self.set_output_formats(WORKBOOK_FORMATS if batch else FORMATS)
        self.release_unused_tables()

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
        focus_preview = self.buttons["preview"].hasFocus()
        self.invalidate_preview()
        source = self.editor.toPlainText()
        mode = self.input_mode.currentData()
        if exceeds_editor_limit(source):
            self.start_job("preparing", lambda: parse_table(source, mode),
                           lambda table: self.show_table(table, focus_preview=focus_preview), "read")
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            table = parse_table(source, mode)
        except ConversionError as exc:
            self.show_status(exc.code, error=True, **exc.details)
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.show_table(table, focus_preview=focus_preview)

    def show_table(self, table: Table, *, focus_preview: bool = False) -> None:
        self.render_table(table)
        self.show_status("summary", rows=len(table.rows), columns=len(table.headers))
        if focus_preview:
            self.grid.setFocus()

    def render_table(self, table: Table) -> None:
        self.grid.clear()
        headers = table.headers[:PREVIEW_COLUMNS]
        rows = list(islice(table.rows, PREVIEW_ROWS))
        self.grid.setColumnCount(len(headers))
        self.grid.setHorizontalHeaderLabels(headers)
        metrics = self.grid.horizontalHeader().fontMetrics()
        for column, header in enumerate(headers):
            self.grid.setColumnWidth(column, min(240, max(104, metrics.horizontalAdvance(header) + 24)))
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
        self.release_unused_tables()

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
        self.content.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.cancel_button.show()
        job.finished.connect(self.finish_job)
        job.start()

    @Slot()
    def cancel_job(self) -> None:
        if self.job is None:
            return
        self.job.cancel()
        self.cancel_button.setEnabled(False)
        self.show_status("cancelling")

    @Slot()
    def finish_job(self) -> None:
        job = self.job
        self.job = None
        self.cancel_button.hide()
        self.content.setEnabled(True)
        job.deleteLater()
        if self.quit_requested:
            for table in tables_in(job.result):
                table.close()
            QTimer.singleShot(0, self.close)
            return
        # A job that completed before it reached a checkpoint has finished its
        # work (a save has already replaced the file), so it reports success.
        if isinstance(job.error, Cancelled):
            self.show_status("cancelled")
        elif isinstance(job.error, BatchInputError):
            self.show_status("batch_failure", error=True, name=job.error.path.name, issue=job.error.issue)
        elif isinstance(job.error, ConversionError):
            self.show_status(job.error.code, error=True, **job.error.details)
        elif job.error is not None:
            self.show_status(job.fallback, error=True)
        else:
            job.on_success(job.result)

    def confirm_quit(self) -> bool:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle(self.text("quit_title"))
        box.setText(self.text("quit_prompt"))
        quit_button = box.addButton(self.text("quit_confirm"), QMessageBox.ButtonRole.DestructiveRole)
        keep = box.addButton(self.text("keep_working"), QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(keep)
        box.setEscapeButton(keep)
        box.exec()
        return box.clickedButton() is quit_button

    def closeEvent(self, event) -> None:
        if self.job is not None:
            # The worker thread must finish before the window and its QThread go
            # away. It stops at its next checkpoint, removes its temporary files
            # and finish_job() closes the window again.
            event.ignore()
            if self.quit_requested:
                self.show_status("cancelling")
                return
            if not self.confirm_quit():
                return
            if self.job is None:
                # The job finished while the question was open.
                QTimer.singleShot(0, self.close)
                return
            self.quit_requested = True
            self.cancel_job()
            return
        self.release_all_tables()
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
    if window.buttons["open"].icon().isNull() or window.brand.pixmap().isNull():
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
