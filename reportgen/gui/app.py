"""보고서 자동 생성기 - tkinter 마법사 UI.

5단계 흐름
----------
1단계  원본 엑셀 선택 및 읽기
2단계  템플릿 선택 (등록된 템플릿 / 신규 업로드)
3단계  매핑 확인 및 수정
4단계  일단위 -> 월단위 집계 설정과 미리보기
5단계  생성

각 단계는 :class:`ReportApp` 의 ``_build_stepN`` 메서드가 만든다. 업무 로직은
전부 ``reportgen`` 하위 모듈에 있고 여기서는 화면과 상태만 다룬다.
"""

from __future__ import annotations

import datetime as _dt
import os
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk
from typing import Any, Optional

from ..aggregator import (
    METHOD_LABELS,
    METHODS,
    AggregationSpec,
    MonthlyResult,
    aggregate_monthly,
    detect_date_column,
    suggest_methods,
)
from ..data_reader import ReadOptions, Table, list_sheets, list_table_blocks, read_table
from ..dateutils import month_label, parse_date
from ..errors import ReportGenError
from ..generator import GenerationRequest, Prepared, generate
from ..mapping import (
    BUILTIN_KEYS,
    Binding,
    MappingProfile,
    TemplateSlot,
    auto_match,
    load_mapping,
    mapping_path_for,
    save_mapping,
)
from ..templating import TemplateRegistry, open_template
from ..templating.excel import CELL_KEY_RE, ExcelTemplate
from .widgets import InlineCombo, InlineEntry, TableView, show_error

__all__ = ["ReportApp", "run"]

SOURCE_LABELS = {
    "column": "엑셀 컬럼",
    "literal": "고정값",
    "builtin": "자동 값",
    "blank": "비움",
}
_LABEL_TO_SOURCE = {v: k for k, v in SOURCE_LABELS.items()}
_METHOD_TO_LABEL = dict(METHOD_LABELS)
_LABEL_TO_METHOD = {v: k for k, v in METHOD_LABELS.items()}


class ReportApp(ttk.Frame):
    def __init__(self, master: tk.Tk, base_dir: str) -> None:
        super().__init__(master, padding=8)
        self.master_window = master
        self.base_dir = os.path.abspath(base_dir)
        self.template_dir = os.path.join(self.base_dir, "templates")
        self.mapping_dir = os.path.join(self.base_dir, "mappings")
        self.output_dir = os.path.join(self.base_dir, "output")
        for directory in (self.template_dir, self.mapping_dir, self.output_dir):
            os.makedirs(directory, exist_ok=True)
        self.registry = TemplateRegistry(self.template_dir)

        # ---- 상태 ----------------------------------------------------- #
        self.table: Optional[Table] = None
        self.handler = None
        self.slots: list[TemplateSlot] = []
        self.bindings: dict[str, Binding] = {}
        self.monthly: Optional[MonthlyResult] = None
        self.overrides: dict[str, dict[str, Any]] = {}
        self.template_path: str = ""

        self.pack(fill="both", expand=True)
        self._build()

    # ------------------------------------------------------------------ #
    # 화면 구성
    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.step1 = ttk.Frame(self.notebook, padding=10)
        self.step2 = ttk.Frame(self.notebook, padding=10)
        self.step3 = ttk.Frame(self.notebook, padding=10)
        self.step4 = ttk.Frame(self.notebook, padding=10)
        self.step5 = ttk.Frame(self.notebook, padding=10)
        for frame, title in (
            (self.step1, "1. 원본 엑셀"),
            (self.step2, "2. 템플릿"),
            (self.step3, "3. 매핑"),
            (self.step4, "4. 집계·미리보기"),
            (self.step5, "5. 생성"),
        ):
            self.notebook.add(frame, text=title)

        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._build_step4()
        self._build_step5()

        self.status = tk.StringVar(value="원본 엑셀 파일을 선택하는 것으로 시작합니다.")
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(6, 0))
        ttk.Separator(bar, orient="horizontal").pack(fill="x", pady=(0, 4))
        ttk.Label(bar, textvariable=self.status, anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Label(bar, text="완전 로컬 실행 · 외부 전송 없음", foreground="#666").pack(side="right")

        self._set_step_enabled(1)

    def _set_step_enabled(self, up_to: int) -> None:
        for index in range(5):
            state = "normal" if index < up_to else "disabled"
            self.notebook.tab(index, state=state)

    def _say(self, message: str) -> None:
        self.status.set(message)
        self.update_idletasks()

    # ------------------------------------------------------------------ #
    # 1단계 - 원본 엑셀
    # ------------------------------------------------------------------ #
    def _build_step1(self) -> None:
        frame = self.step1
        top = ttk.LabelFrame(frame, text="원본 엑셀 파일", padding=8)
        top.pack(fill="x")

        self.src_path = tk.StringVar()
        row = ttk.Frame(top)
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.src_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="찾아보기…", command=self._pick_source).pack(side="left", padx=(6, 0))

        options = ttk.LabelFrame(frame, text="읽기 옵션", padding=8)
        options.pack(fill="x", pady=(8, 0))

        line1 = ttk.Frame(options)
        line1.pack(fill="x", pady=2)
        ttk.Label(line1, text="시트", width=10).pack(side="left")
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(line1, textvariable=self.sheet_var, state="readonly", width=28)
        self.sheet_combo.pack(side="left")

        self.auto_detect = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            line1,
            text="표 위치·헤더 자동 감지",
            variable=self.auto_detect,
            command=self._toggle_auto,
        ).pack(side="left", padx=(16, 0))

        ttk.Button(line1, text="표 후보 찾기…", command=self._find_table_blocks).pack(
            side="left", padx=(16, 0)
        )

        line2 = ttk.Frame(options)
        line2.pack(fill="x", pady=2)
        ttk.Label(line2, text="셀 범위", width=10).pack(side="left")
        self.range_var = tk.StringVar()
        self.range_entry = ttk.Entry(line2, textvariable=self.range_var, width=16)
        self.range_entry.pack(side="left")
        ttk.Label(line2, text="예: B3:G40 (비우면 자동)", foreground="#666").pack(side="left", padx=(6, 0))

        ttk.Label(line2, text="헤더 행 수", width=10).pack(side="left", padx=(20, 0))
        self.header_rows = tk.IntVar(value=1)
        ttk.Spinbox(line2, from_=1, to=5, width=4, textvariable=self.header_rows).pack(side="left")

        self.transpose = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            line2, text="행/열 바꿔 읽기 (날짜가 가로로 늘어선 표)", variable=self.transpose
        ).pack(side="left", padx=(20, 0))

        line3 = ttk.Frame(options)
        line3.pack(fill="x", pady=2)
        self.allow_uncalculated = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            line3,
            text="계산되지 않은 수식 셀은 빈 값으로 넘어가기 (평소엔 끄세요)",
            variable=self.allow_uncalculated,
        ).pack(side="left")
        ttk.Label(
            line3,
            text="※ 원본을 엑셀에서 다시 열어 저장할 수 없을 때만 켜세요. 해당 값은 집계에서 빠집니다.",
            foreground="#a60",
        ).pack(side="left", padx=(6, 0))

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="데이터 읽기", command=self._load_source).pack(side="left")
        self.src_info = tk.StringVar(value="")
        ttk.Label(buttons, textvariable=self.src_info, foreground="#0a6").pack(side="left", padx=(10, 0))
        ttk.Button(buttons, text="다음 단계 ▶", command=lambda: self._goto(1)).pack(side="right")

        preview = ttk.LabelFrame(frame, text="미리보기 (최대 200행)", padding=6)
        preview.pack(fill="both", expand=True, pady=(8, 0))
        self.src_preview = TableView(preview, height=14)
        self.src_preview.pack(fill="both", expand=True)

        self._toggle_auto()

    def _toggle_auto(self) -> None:
        state = "disabled" if self.auto_detect.get() else "normal"
        self.range_entry.configure(state=state)

    def _find_table_blocks(self) -> None:
        """한 시트 안에 표가 여러 개 섞여 있을 때, 후보를 찾아 고르게 한다.

        실무 엑셀은 시트 하나에 표 하나만 깔끔히 있는 경우가 오히려 드물다.
        '자동 감지'가 여러 표를 한 덩어리로 잘못 잡는 문제를, 후보 목록을
        보여주고 사람이 고르는 방식으로 피해간다.
        """
        path = self.src_path.get()
        if not path:
            messagebox.showinfo("안내", "먼저 원본 엑셀 파일을 선택해 주세요.")
            return
        try:
            blocks = list_table_blocks(path, self.sheet_var.get() or None)
        except ReportGenError as exc:
            show_error("표 후보 찾기 실패", exc)
            return

        if not blocks:
            messagebox.showinfo("표 후보 찾기", "이 시트에서 표처럼 보이는 덩어리를 찾지 못했습니다.")
            return

        dialog = tk.Toplevel(self)
        dialog.title("표 후보 — 원하는 표를 고르세요")
        dialog.geometry("640x320")
        dialog.transient(self.master_window)

        ttk.Label(
            dialog,
            text=f"이 시트에서 표처럼 보이는 덩어리 {len(blocks)}개를 찾았습니다. "
            "하나를 골라 [이 표 사용] 을 누르면 셀 범위가 채워집니다.",
            wraplength=600,
            justify="left",
        ).pack(fill="x", padx=10, pady=(10, 6))

        listbox = tk.Listbox(dialog, height=10, font=("", 10))
        for block in blocks:
            listbox.insert("end", block.label())
        listbox.pack(fill="both", expand=True, padx=10)
        if blocks:
            listbox.selection_set(0)

        def use_selected() -> None:
            selection = listbox.curselection()
            if not selection:
                return
            block = blocks[selection[0]]
            self.auto_detect.set(False)
            self._toggle_auto()
            self.range_var.set(block.range)
            dialog.destroy()
            self._say(f"표 범위를 '{block.range}' 로 지정했습니다. [데이터 읽기]를 눌러 확인하세요.")

        buttons = ttk.Frame(dialog)
        buttons.pack(fill="x", padx=10, pady=10)
        ttk.Button(buttons, text="이 표 사용", command=use_selected).pack(side="right")
        ttk.Button(buttons, text="닫기", command=dialog.destroy).pack(side="right", padx=(0, 6))
        listbox.bind("<Double-1>", lambda _e: use_selected())

    def _pick_source(self) -> None:
        path = filedialog.askopenfilename(
            title="원본 엑셀 파일 선택",
            filetypes=[("엑셀 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        self.src_path.set(path)
        try:
            sheets = list_sheets(path)
        except ReportGenError as exc:
            show_error("파일 열기 실패", exc)
            return
        self.sheet_combo["values"] = sheets
        if sheets:
            self.sheet_var.set(sheets[0])
        self._say(f"시트 {len(sheets)}개를 찾았습니다. 시트를 고르고 [데이터 읽기]를 누르세요.")

    def _read_options(self) -> ReadOptions:
        return ReadOptions(
            sheet_name=self.sheet_var.get() or None,
            cell_range=None if self.auto_detect.get() else (self.range_var.get().strip() or None),
            header_rows=max(1, int(self.header_rows.get() or 1)),
            auto_detect=self.auto_detect.get(),
            transpose=self.transpose.get(),
            allow_uncalculated_formulas=self.allow_uncalculated.get(),
        )

    def _load_source(self) -> None:
        try:
            table = read_table(self.src_path.get(), self._read_options())
        except ReportGenError as exc:
            show_error("데이터 읽기 실패", exc)
            return
        except Exception as exc:  # noqa: BLE001
            show_error("데이터 읽기 실패", exc)
            return

        self.table = table
        self.monthly = None
        self.overrides = {}
        self.src_preview.load_matrix(table.preview(200), max_rows=200)
        self.src_info.set(f"컬럼 {len(table.columns)}개 · 데이터 {table.n_rows}행")
        if table.warnings:
            messagebox.showwarning("주의", "\n\n".join(table.warnings))
            self._say("원본을 읽었습니다 (일부 값을 빈 값으로 건너뜀). 2단계로 진행하세요.")
        else:
            self._say("원본을 읽었습니다. 2단계에서 템플릿을 고르세요.")
        self._set_step_enabled(2)
        self._refresh_aggregation_inputs()

    # ------------------------------------------------------------------ #
    # 2단계 - 템플릿
    # ------------------------------------------------------------------ #
    def _build_step2(self) -> None:
        frame = self.step2
        self.template_mode = tk.StringVar(value="registered")

        registered = ttk.LabelFrame(frame, text="등록된 템플릿 사용", padding=8)
        registered.pack(fill="x")
        row = ttk.Frame(registered)
        row.pack(fill="x")
        ttk.Radiobutton(row, text="사용", variable=self.template_mode, value="registered").pack(side="left")
        self.registered_var = tk.StringVar()
        self.registered_combo = ttk.Combobox(
            row, textvariable=self.registered_var, state="readonly", width=48
        )
        self.registered_combo.pack(side="left", padx=(8, 0))
        ttk.Button(row, text="목록 새로고침", command=self._refresh_registry).pack(side="left", padx=(6, 0))
        ttk.Button(row, text="폴더 열기", command=lambda: _open_folder(self.template_dir)).pack(
            side="left", padx=(4, 0)
        )

        upload = ttk.LabelFrame(frame, text="신규 템플릿 업로드", padding=8)
        upload.pack(fill="x", pady=(8, 0))
        row2 = ttk.Frame(upload)
        row2.pack(fill="x")
        ttk.Radiobutton(row2, text="사용", variable=self.template_mode, value="upload").pack(side="left")
        self.upload_path = tk.StringVar()
        ttk.Entry(row2, textvariable=self.upload_path).pack(side="left", fill="x", expand=True, padx=(8, 0))
        ttk.Button(row2, text="찾아보기…", command=self._pick_template).pack(side="left", padx=(6, 0))
        self.register_after = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            upload, text="templates/ 폴더에 복사해서 다음에도 목록에 보이게 하기", variable=self.register_after
        ).pack(anchor="w", pady=(4, 0))

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="템플릿 분석", command=self._analyze_template).pack(side="left")
        self.tpl_info = tk.StringVar(value="")
        ttk.Label(buttons, textvariable=self.tpl_info, foreground="#0a6").pack(side="left", padx=(10, 0))
        ttk.Button(buttons, text="다음 단계 ▶", command=lambda: self._goto(2)).pack(side="right")

        found = ttk.LabelFrame(frame, text="템플릿에서 찾은 항목", padding=6)
        found.pack(fill="both", expand=True, pady=(8, 0))
        self.slot_view = TableView(found, height=12)
        self.slot_view.pack(fill="both", expand=True)

        self._refresh_registry()

    def _refresh_registry(self) -> None:
        labels = self.registry.labels()
        self.registered_combo["values"] = labels
        if labels and not self.registered_var.get():
            self.registered_var.set(labels[0])

    def _pick_template(self) -> None:
        path = filedialog.askopenfilename(
            title="템플릿 파일 선택",
            filetypes=[("워드/엑셀 템플릿", "*.docx *.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if path:
            self.upload_path.set(path)
            self.template_mode.set("upload")

    def _resolve_template_path(self) -> str:
        if self.template_mode.get() == "registered":
            item = self.registry.find_by_label(self.registered_var.get())
            if item is None:
                raise ReportGenError(
                    "등록된 템플릿을 고르지 않았습니다.",
                    f"templates 폴더({self.template_dir})에 .docx / .xlsx 파일을 넣고 "
                    "[목록 새로고침]을 눌러 주세요.",
                )
            return item.path
        path = self.upload_path.get().strip()
        if not path:
            raise ReportGenError("템플릿 파일을 선택해 주세요.")
        if self.register_after.get():
            return self.registry.register(path).path
        return path

    def _analyze_template(self) -> None:
        try:
            path = self._resolve_template_path()
            handler = open_template(path)
            slots = handler.scan()
        except ReportGenError as exc:
            show_error("템플릿 분석 실패", exc)
            return
        except Exception as exc:  # noqa: BLE001
            show_error("템플릿 분석 실패", exc)
            return

        self.template_path = path
        self.handler = handler
        self.slots = slots
        self._refresh_registry()

        matrix: list[list[Any]] = [["항목(태그)", "종류", "위치", "횟수", "주변 내용"]]
        for slot in slots:
            kind = {"tag": "태그", "cell": "셀", "table": "표 삽입 위치"}.get(slot.kind, slot.kind)
            matrix.append([slot.key, kind, slot.where, slot.occurrences, slot.sample])
        self.slot_view.load_matrix(matrix)
        self.tpl_info.set(f"{handler.describe()} · 항목 {len(slots)}개")

        # 저장된 매핑이 있으면 불러오고, 없으면 이름으로 자동 매칭
        columns = self.table.columns if self.table else []
        profile = None
        try:
            profile = load_mapping(mapping_path_for(path, self.mapping_dir))
        except ReportGenError as exc:
            messagebox.showwarning("매핑 불러오기", str(exc))

        if profile and profile.bindings:
            self.bindings = dict(profile.bindings)
            for slot in slots:
                self.bindings.setdefault(slot.key, Binding())
            self._apply_profile_settings(profile)
            self._say("저장된 매핑을 불러왔습니다. 3단계에서 확인하세요.")
        else:
            self.bindings = auto_match(slots, columns)
            self._say("이름이 비슷한 항목을 자동으로 연결했습니다. 3단계에서 확인하세요.")

        self._refresh_mapping_view()
        self._set_step_enabled(5 if self.table else 3)
        self.notebook.select(self.step3)

    def _apply_profile_settings(self, profile: MappingProfile) -> None:
        agg = profile.aggregation or {}
        if not agg:
            return
        self.use_agg.set(bool(agg.get("enabled")))
        if agg.get("date_column"):
            self.date_col_var.set(agg["date_column"])
        self.default_method.set(_METHOD_TO_LABEL.get(agg.get("default_method", "sum"), "합산"))
        self.exclude_weekends.set(bool(agg.get("exclude_weekends")))
        self.multi_mode.set(agg.get("multi_month_mode", "separate"))
        if agg.get("base_year"):
            self.base_year.set(str(agg["base_year"]))
        if agg.get("base_month"):
            self.base_month.set(str(agg["base_month"]))
        self._saved_methods = dict(agg.get("methods") or {})
        self._refresh_method_view()

    # ------------------------------------------------------------------ #
    # 3단계 - 매핑
    # ------------------------------------------------------------------ #
    def _build_step3(self) -> None:
        frame = self.step3

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="이름으로 자동 매칭", command=self._auto_match).pack(side="left")
        ttk.Button(buttons, text="매핑 저장", command=self._save_mapping).pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text="매핑 불러오기…", command=self._load_mapping_dialog).pack(side="left", padx=(6, 0))
        self.add_cell_button = ttk.Button(
            buttons, text="엑셀 셀 좌표 추가…", command=self._add_cell_slot
        )
        self.add_cell_button.pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text="다음 단계 ▶", command=lambda: self._goto(3)).pack(side="right")

        table_frame = ttk.LabelFrame(frame, text="항목별 연결", padding=6)
        table_frame.pack(fill="both", expand=True, pady=(8, 0))
        self.map_view = TableView(table_frame, height=12)
        self.map_view.pack(fill="both", expand=True)
        self.map_view.tree.bind("<<TreeviewSelect>>", self._on_map_select)

        editor = ttk.LabelFrame(frame, text="선택한 항목 편집", padding=8)
        editor.pack(fill="x", pady=(8, 0))

        line = ttk.Frame(editor)
        line.pack(fill="x")
        ttk.Label(line, text="항목", width=8).pack(side="left")
        self.edit_key = tk.StringVar()
        ttk.Label(line, textvariable=self.edit_key, font=("", 10, "bold")).pack(side="left")

        line2 = ttk.Frame(editor)
        line2.pack(fill="x", pady=(6, 0))
        ttk.Label(line2, text="연결 방식", width=8).pack(side="left")
        self.edit_source = tk.StringVar(value="엑셀 컬럼")
        for label in SOURCE_LABELS.values():
            ttk.Radiobutton(
                line2, text=label, value=label, variable=self.edit_source, command=self._on_source_change
            ).pack(side="left", padx=(0, 10))

        line3 = ttk.Frame(editor)
        line3.pack(fill="x", pady=(6, 0))
        ttk.Label(line3, text="값", width=8).pack(side="left")
        self.edit_column = ttk.Combobox(line3, state="readonly", width=26)
        self.edit_column.pack(side="left")
        self.edit_literal = ttk.Entry(line3, width=28)
        self.edit_literal.pack(side="left", padx=(6, 0))
        self.edit_builtin = ttk.Combobox(line3, state="readonly", width=16, values=list(BUILTIN_KEYS))
        self.edit_builtin.pack(side="left", padx=(6, 0))

        line4 = ttk.Frame(editor)
        line4.pack(fill="x", pady=(6, 0))
        ttk.Label(line4, text="행 지정", width=8).pack(side="left")
        self.edit_row = ttk.Combobox(line4, state="readonly", width=14)
        self.edit_row.pack(side="left")
        ttk.Label(line4, text="숫자 형식").pack(side="left", padx=(14, 4))
        self.edit_format = ttk.Combobox(
            line4, width=10, values=["", ",.0f", ",.1f", ",.2f", ".1f", ".2f"]
        )
        self.edit_format.pack(side="left")
        ttk.Label(line4, text="단위").pack(side="left", padx=(14, 4))
        self.edit_suffix = ttk.Entry(line4, width=10)
        self.edit_suffix.pack(side="left")
        ttk.Button(line4, text="적용", command=self._apply_binding).pack(side="left", padx=(16, 0))

    def _refresh_mapping_view(self) -> None:
        columns = self.table.columns if self.table else []
        self.edit_column["values"] = columns
        rows = ["집계값 / 첫 행"]
        if self.table:
            rows += [f"{i + 1}행" for i in range(min(self.table.n_rows, 200))]
        self.edit_row["values"] = rows

        keys = [slot.key for slot in self.slots]
        for key in self.bindings:
            if key not in keys:
                keys.append(key)

        where = {slot.key: slot.where for slot in self.slots}
        matrix: list[list[Any]] = [["항목(태그)", "위치", "연결", "미리보기"]]
        for key in keys:
            binding = self.bindings.get(key, Binding())
            matrix.append([key, where.get(key, "(직접 추가)"), binding.describe(), self._peek(binding)])
        self.map_view.load_matrix(matrix)

        is_excel = isinstance(self.handler, ExcelTemplate)
        self.add_cell_button.configure(state="normal" if is_excel else "disabled")

    def _peek(self, binding: Binding) -> str:
        """매핑 화면에 보여줄 '지금 값이라면 이렇게 들어갑니다' 문자열."""
        if binding.source == "literal":
            return binding.literal
        if binding.source == "builtin":
            return f"<{binding.builtin}>"
        if binding.source == "column" and self.table:
            if self.monthly and binding.row is None and binding.column in self.monthly.columns:
                period = self.monthly.periods[0]
                return str(self.monthly.get(period, binding.column))
            if binding.column in self.table.columns and self.table.rows:
                index = binding.row or 0
                if 0 <= index < self.table.n_rows:
                    return str(self.table.cell(index, binding.column))
        return ""

    def _current_key(self) -> Optional[str]:
        values = self.map_view.selected_values()
        return values[0] if values else None

    def _on_map_select(self, _event: tk.Event) -> None:
        key = self._current_key()
        if key is None:
            return
        binding = self.bindings.get(key, Binding())
        self.edit_key.set(key)
        self.edit_source.set(SOURCE_LABELS.get(binding.source, "비움"))
        self.edit_column.set(binding.column)
        self.edit_literal.delete(0, "end")
        self.edit_literal.insert(0, binding.literal)
        self.edit_builtin.set(binding.builtin)
        self.edit_row.set("집계값 / 첫 행" if binding.row is None else f"{binding.row + 1}행")
        self.edit_format.set(binding.number_format)
        self.edit_suffix.delete(0, "end")
        self.edit_suffix.insert(0, binding.suffix)
        self._on_source_change()

    def _on_source_change(self) -> None:
        source = _LABEL_TO_SOURCE.get(self.edit_source.get(), "blank")
        self.edit_column.configure(state="readonly" if source == "column" else "disabled")
        self.edit_literal.configure(state="normal" if source == "literal" else "disabled")
        self.edit_builtin.configure(state="readonly" if source == "builtin" else "disabled")
        self.edit_row.configure(state="readonly" if source == "column" else "disabled")

    def _apply_binding(self) -> None:
        key = self.edit_key.get()
        if not key:
            messagebox.showinfo("안내", "먼저 위 표에서 항목을 하나 고르세요.")
            return
        source = _LABEL_TO_SOURCE.get(self.edit_source.get(), "blank")
        row_text = self.edit_row.get()
        row = None
        if source == "column" and row_text.endswith("행"):
            try:
                row = int(row_text[:-1]) - 1
            except ValueError:
                row = None
        self.bindings[key] = Binding(
            source=source,
            column=self.edit_column.get() if source == "column" else "",
            row=row,
            literal=self.edit_literal.get() if source == "literal" else "",
            builtin=self.edit_builtin.get() if source == "builtin" else "",
            number_format=self.edit_format.get().strip(),
            suffix=self.edit_suffix.get(),
        )
        self._refresh_mapping_view()
        self._say(f"'{key}' 연결을 수정했습니다.")

    def _auto_match(self) -> None:
        if not self.table:
            messagebox.showinfo("안내", "1단계에서 원본 엑셀을 먼저 읽어 주세요.")
            return
        self.bindings = auto_match(self.slots, self.table.columns)
        self._refresh_mapping_view()
        self._say("이름이 비슷한 항목을 자동으로 연결했습니다.")

    def _add_cell_slot(self) -> None:
        from tkinter import simpledialog

        if not isinstance(self.handler, ExcelTemplate):
            return
        sheets = self.handler.sheet_names()
        answer = simpledialog.askstring(
            "엑셀 셀 좌표 추가",
            "값을 넣을 위치를 [시트명]![셀] 형식으로 입력하세요.\n"
            f"이 템플릿의 시트: {', '.join(sheets)}\n\n예) {sheets[0]}!B3",
            parent=self,
        )
        if not answer:
            return
        answer = answer.strip()
        if not CELL_KEY_RE.match(answer):
            messagebox.showerror(
                "형식 오류",
                f"'{answer}' 는 올바른 형식이 아닙니다.\n'{sheets[0]}!B3' 처럼 입력해 주세요.",
            )
            return
        self.slots.append(TemplateSlot(key=answer, kind="cell", where=answer, sample="직접 추가"))
        self.bindings.setdefault(answer, Binding())
        self._refresh_mapping_view()

    def _save_mapping(self) -> None:
        if not self.template_path:
            messagebox.showinfo("안내", "2단계에서 템플릿을 먼저 분석해 주세요.")
            return
        profile = MappingProfile(
            template_name=os.path.basename(self.template_path),
            template_type=getattr(self.handler, "kind", ""),
            bindings=self.bindings,
            aggregation=self._aggregation_snapshot(),
            read_options=self._read_options().__dict__,
        )
        path = save_mapping(profile, mapping_path_for(self.template_path, self.mapping_dir))
        self._say(f"매핑을 저장했습니다: {path}")
        messagebox.showinfo("저장 완료", f"매핑을 저장했습니다.\n\n{path}")

    def _load_mapping_dialog(self) -> None:
        path = filedialog.askopenfilename(
            title="매핑 파일 선택",
            initialdir=self.mapping_dir,
            filetypes=[("매핑 파일", "*.json"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        try:
            profile = load_mapping(path)
        except ReportGenError as exc:
            show_error("매핑 불러오기 실패", exc)
            return
        if profile is None:
            return
        self.bindings = dict(profile.bindings)
        for slot in self.slots:
            self.bindings.setdefault(slot.key, Binding())
        self._apply_profile_settings(profile)
        self._refresh_mapping_view()
        self._say(f"매핑을 불러왔습니다: {os.path.basename(path)}")

    # ------------------------------------------------------------------ #
    # 4단계 - 집계
    # ------------------------------------------------------------------ #
    def _build_step4(self) -> None:
        frame = self.step4
        self._saved_methods: dict[str, str] = {}

        head = ttk.Frame(frame)
        head.pack(fill="x")
        self.use_agg = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            head,
            text="일단위 데이터를 월단위로 집계해서 사용",
            variable=self.use_agg,
            command=self._toggle_agg,
        ).pack(side="left")
        ttk.Button(head, text="다음 단계 ▶", command=lambda: self._goto(4)).pack(side="right")

        self.agg_body = ttk.Frame(frame)
        self.agg_body.pack(fill="both", expand=True, pady=(8, 0))

        settings = ttk.LabelFrame(self.agg_body, text="집계 설정", padding=8)
        settings.pack(fill="x")

        line1 = ttk.Frame(settings)
        line1.pack(fill="x", pady=2)
        ttk.Label(line1, text="날짜 컬럼", width=12).pack(side="left")
        self.date_col_var = tk.StringVar()
        self.date_col_combo = ttk.Combobox(line1, textvariable=self.date_col_var, state="readonly", width=26)
        self.date_col_combo.pack(side="left")
        ttk.Button(line1, text="자동 감지", command=self._detect_date_column).pack(side="left", padx=(6, 0))

        ttk.Label(line1, text="기준 연/월", width=10).pack(side="left", padx=(20, 0))
        self.base_year = tk.StringVar()
        self.base_month = tk.StringVar()
        ttk.Entry(line1, textvariable=self.base_year, width=6).pack(side="left")
        ttk.Label(line1, text="년").pack(side="left")
        ttk.Entry(line1, textvariable=self.base_month, width=4).pack(side="left")
        ttk.Label(line1, text="월  (‘5일’처럼 일자만 있을 때)", foreground="#666").pack(side="left")

        line2 = ttk.Frame(settings)
        line2.pack(fill="x", pady=2)
        self.exclude_weekends = tk.BooleanVar(value=False)
        ttk.Checkbutton(line2, text="주말(토·일) 제외", variable=self.exclude_weekends).pack(side="left")
        ttk.Label(line2, text="공휴일 제외").pack(side="left", padx=(20, 4))
        self.holidays = tk.StringVar()
        ttk.Entry(line2, textvariable=self.holidays, width=34).pack(side="left")
        ttk.Label(line2, text="예: 2026-01-01, 2026-03-01", foreground="#666").pack(side="left", padx=(6, 0))

        line3 = ttk.Frame(settings)
        line3.pack(fill="x", pady=2)
        ttk.Label(line3, text="기본 집계", width=12).pack(side="left")
        self.default_method = tk.StringVar(value="합산")
        ttk.Combobox(
            line3, textvariable=self.default_method, state="readonly",
            width=12, values=list(METHOD_LABELS.values()),
        ).pack(side="left")

        ttk.Label(line3, text="여러 달이 섞여 있으면", width=20).pack(side="left", padx=(20, 0))
        self.multi_mode = tk.StringVar(value="separate")
        ttk.Radiobutton(line3, text="월별 보고서 각각", variable=self.multi_mode, value="separate").pack(side="left")
        ttk.Radiobutton(
            line3, text="한 보고서에 월별로 나열", variable=self.multi_mode, value="wide"
        ).pack(side="left", padx=(8, 0))

        middle = ttk.Frame(self.agg_body)
        middle.pack(fill="both", expand=True, pady=(8, 0))

        methods = ttk.LabelFrame(middle, text="항목별 집계 방식 (방식 칸을 더블클릭)", padding=6)
        methods.pack(side="left", fill="both", expand=True)
        self.method_view = TableView(methods, height=10)
        self.method_view.pack(fill="both", expand=True)
        self.method_editor = InlineCombo(
            self.method_view.tree, "#2", list(METHOD_LABELS.values()), self._on_method_changed
        )

        months = ttk.LabelFrame(middle, text="대상 월 (선택 안 하면 전체)", padding=6)
        months.pack(side="left", fill="y", padx=(8, 0))
        self.month_list = tk.Listbox(months, selectmode="multiple", height=10, exportselection=False, width=14)
        self.month_list.pack(fill="both", expand=True)

        actions = ttk.Frame(self.agg_body)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="집계 미리보기 계산", command=self._compute_preview).pack(side="left")
        self.agg_info = tk.StringVar(value="")
        ttk.Label(actions, textvariable=self.agg_info, foreground="#0a6").pack(side="left", padx=(10, 0))

        preview = ttk.LabelFrame(
            self.agg_body, text="집계 결과 (값을 더블클릭하면 직접 고칠 수 있습니다)", padding=6
        )
        preview.pack(fill="both", expand=True, pady=(8, 0))
        self.agg_view = TableView(preview, height=8)
        self.agg_view.pack(fill="both", expand=True)
        self.agg_editor = InlineEntry(self.agg_view.tree, self._on_override)

        self._toggle_agg()

    def _toggle_agg(self) -> None:
        state = "normal" if self.use_agg.get() else "disabled"
        _set_state_recursive(self.agg_body, state)
        if not self.use_agg.get():
            self.monthly = None
            self.overrides = {}

    def _refresh_aggregation_inputs(self) -> None:
        if not self.table:
            return
        self.date_col_combo["values"] = self.table.columns
        self._detect_date_column(silent=True)
        self._refresh_method_view()
        self._refresh_month_list()

    def _detect_date_column(self, silent: bool = False) -> None:
        if not self.table:
            return
        found = detect_date_column(self.table, self._spec_for_detection())
        if found:
            self.date_col_var.set(found)
            if not silent:
                self._say(f"날짜 컬럼으로 '{found}' 를 골랐습니다.")
        elif not silent:
            messagebox.showinfo(
                "자동 감지",
                "날짜로 읽히는 컬럼을 찾지 못했습니다.\n"
                "직접 고르거나, '5일'처럼 일자만 있는 표라면 기준 연/월을 입력해 주세요.",
            )
        self._refresh_method_view()
        self._refresh_month_list()

    def _spec_for_detection(self) -> AggregationSpec:
        return AggregationSpec(
            base_year=_int_or_none(self.base_year.get()),
            base_month=_int_or_none(self.base_month.get()),
        )

    def _refresh_method_view(self) -> None:
        if not self.table:
            return
        date_column = self.date_col_var.get()
        suggested = suggest_methods(self.table, date_column)
        suggested.update(
            {k: v for k, v in getattr(self, "_saved_methods", {}).items() if k in suggested}
        )
        self._methods = suggested
        matrix: list[list[Any]] = [["항목", "집계 방식"]]
        for column, method in suggested.items():
            matrix.append([column, _METHOD_TO_LABEL.get(method, method)])
        self.method_view.load_matrix(matrix)

    def _on_method_changed(self, item: str, label: str) -> None:
        column = self.method_view.tree.set(item, "#1")
        self._methods[column] = _LABEL_TO_METHOD.get(label, "sum")

    def _refresh_month_list(self) -> None:
        self.month_list.delete(0, "end")
        if not self.table or not self.date_col_var.get():
            return
        base_year = _int_or_none(self.base_year.get())
        base_month = _int_or_none(self.base_month.get())
        keys: list[str] = []
        try:
            values = self.table.column_values(self.date_col_var.get())
        except KeyError:
            return
        for value in values:
            day = parse_date(value, base_year, base_month)
            if day is not None:
                key = f"{day.year:04d}-{day.month:02d}"
                if key not in keys:
                    keys.append(key)
        for key in sorted(keys):
            self.month_list.insert("end", key)

    def _selected_months(self) -> list[str]:
        return [self.month_list.get(i) for i in self.month_list.curselection()]

    def _aggregation_spec(self) -> AggregationSpec:
        return AggregationSpec(
            date_column=self.date_col_var.get() or None,
            methods=dict(getattr(self, "_methods", {})),
            default_method=_LABEL_TO_METHOD.get(self.default_method.get(), "sum"),
            exclude_weekends=self.exclude_weekends.get(),
            exclude_dates=_parse_holidays(self.holidays.get()),
            only_months=self._selected_months(),
            base_year=_int_or_none(self.base_year.get()),
            base_month=_int_or_none(self.base_month.get()),
            multi_month_mode=self.multi_mode.get(),
        )

    def _aggregation_snapshot(self) -> dict[str, Any]:
        spec = self._aggregation_spec()
        return {
            "enabled": self.use_agg.get(),
            "date_column": spec.date_column,
            "methods": spec.methods,
            "default_method": spec.default_method,
            "exclude_weekends": spec.exclude_weekends,
            "base_year": spec.base_year,
            "base_month": spec.base_month,
            "multi_month_mode": spec.multi_month_mode,
        }

    def _compute_preview(self) -> None:
        if not self.table:
            messagebox.showinfo("안내", "1단계에서 원본 엑셀을 먼저 읽어 주세요.")
            return
        try:
            monthly = aggregate_monthly(self.table, self._aggregation_spec())
        except ReportGenError as exc:
            show_error("집계 실패", exc)
            return
        for period, columns in self.overrides.items():
            for column, value in columns.items():
                if period in monthly.values and column in monthly.values[period]:
                    monthly.set(period, column, value)
        self.monthly = monthly
        self.agg_view.load_matrix(monthly.as_matrix())
        note = f"{len(monthly.periods)}개월 집계"
        if monthly.skipped_rows:
            note += f" · 날짜 인식 실패 {monthly.skipped_rows}행 제외"
        self.agg_info.set(note)
        self._refresh_mapping_view()
        self._say("집계했습니다. 값이 이상하면 표에서 직접 고칠 수 있습니다.")

    def _on_override(self, item: str, column: str, value: str) -> None:
        if self.monthly is None:
            return
        try:
            index = int(column.replace("#", "")) - 1
        except ValueError:
            return
        header = list(self.agg_view.tree["columns"])
        if index < 2 or index >= len(header):  # 0=연-월, 1=일수 는 편집 대상 아님
            return
        column_name = header[index]
        row_index = self.agg_view.tree.index(item)
        if row_index >= len(self.monthly.periods):
            return
        period = self.monthly.periods[row_index]
        parsed: Any = value
        try:
            parsed = float(value.replace(",", ""))
            if parsed == int(parsed):
                parsed = int(parsed)
        except (ValueError, AttributeError):
            pass
        self.monthly.set(period, column_name, parsed)
        self.overrides.setdefault(period, {})[column_name] = parsed
        self._refresh_mapping_view()
        self._say(f"{period} · {column_name} 값을 {parsed} 로 고쳤습니다.")

    # ------------------------------------------------------------------ #
    # 5단계 - 생성
    # ------------------------------------------------------------------ #
    def _build_step5(self) -> None:
        frame = self.step5

        out = ttk.LabelFrame(frame, text="저장 위치", padding=8)
        out.pack(fill="x")
        row = ttk.Frame(out)
        row.pack(fill="x")
        self.out_dir = tk.StringVar(value=self.output_dir)
        ttk.Entry(row, textvariable=self.out_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="폴더 선택…", command=self._pick_output).pack(side="left", padx=(6, 0))
        ttk.Button(row, text="폴더 열기", command=lambda: _open_folder(self.out_dir.get())).pack(
            side="left", padx=(4, 0)
        )
        ttk.Label(
            out,
            text="파일 이름은 [템플릿명]_[생성일시] 형식으로 자동 지정됩니다.",
            foreground="#666",
        ).pack(anchor="w", pady=(4, 0))

        options = ttk.Frame(frame)
        options.pack(fill="x", pady=(8, 0))
        self.include_table = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options,
            text="{{#표}} / {% for r in rows %} 자리에 표 데이터도 채우기",
            variable=self.include_table,
        ).pack(side="left")

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(10, 0))
        self.generate_button = ttk.Button(actions, text="보고서 생성", command=self._generate)
        self.generate_button.pack(side="left")
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=180)
        self.progress.pack(side="left", padx=(10, 0))

        result = ttk.LabelFrame(frame, text="결과", padding=6)
        result.pack(fill="both", expand=True, pady=(8, 0))
        self.result_text = tk.Text(result, height=12, wrap="word")
        scroll = ttk.Scrollbar(result, orient="vertical", command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scroll.set, state="disabled")
        self.result_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def _pick_output(self) -> None:
        path = filedialog.askdirectory(title="저장 폴더 선택", initialdir=self.out_dir.get())
        if path:
            self.out_dir.set(path)

    def _build_request(self) -> GenerationRequest:
        return GenerationRequest(
            source_path=self.src_path.get(),
            read_options=self._read_options(),
            template_path=self.template_path,
            bindings=self.bindings,
            use_aggregation=self.use_agg.get(),
            aggregation=self._aggregation_spec(),
            output_dir=self.out_dir.get() or self.output_dir,
            multi_month_mode=self.multi_mode.get(),
            include_table=self.include_table.get(),
            overrides=self.overrides,
        )

    def _generate(self) -> None:
        if not self.table or not self.handler:
            messagebox.showinfo("안내", "1~2단계를 먼저 마쳐 주세요.")
            return

        request = self._build_request()
        prepared = Prepared(
            table=self.table,
            monthly=self.monthly if self.use_agg.get() else None,
            handler=self.handler,
            slots=self.slots,
            warnings=[],
        )
        if self.use_agg.get() and self.monthly is None:
            try:
                prepared.monthly = aggregate_monthly(self.table, request.aggregation)
            except ReportGenError as exc:
                show_error("집계 실패", exc)
                return

        self.generate_button.configure(state="disabled")
        self.progress.start(12)
        self._say("보고서를 만드는 중입니다…")

        holder: dict[str, Any] = {}

        def work() -> None:
            try:
                holder["result"] = generate(request, prepared)
            except BaseException as exc:  # noqa: BLE001 - 스레드 밖으로 넘긴다
                holder["error"] = exc
                holder["trace"] = traceback.format_exc()

        thread = threading.Thread(target=work, daemon=True)
        thread.start()
        self._poll_generation(thread, holder)

    def _poll_generation(self, thread: threading.Thread, holder: dict[str, Any]) -> None:
        if thread.is_alive():
            self.after(120, lambda: self._poll_generation(thread, holder))
            return

        self.progress.stop()
        self.generate_button.configure(state="normal")

        if "error" in holder:
            self._write_result(holder.get("trace", ""))
            show_error("보고서 생성 실패", holder["error"])
            self._say("생성에 실패했습니다.")
            return

        result = holder["result"]
        self._write_result(result.summary())
        self._say(f"완료: 파일 {len(result.files)}개")
        messagebox.showinfo("생성 완료", result.summary())

    def _write_result(self, text: str) -> None:
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.result_text.insert("1.0", f"[{stamp}]\n{text}\n")
        self.result_text.configure(state="disabled")

    # ------------------------------------------------------------------ #
    def _goto(self, current: int) -> None:
        frames = [self.step1, self.step2, self.step3, self.step4, self.step5]
        if current == 1 and self.table is None:
            messagebox.showinfo("안내", "먼저 [데이터 읽기]를 눌러 원본을 읽어 주세요.")
            return
        if current == 2 and self.handler is None:
            messagebox.showinfo("안내", "먼저 [템플릿 분석]을 눌러 주세요.")
            return
        self._set_step_enabled(5)
        self.notebook.select(frames[current])


# --------------------------------------------------------------------------- #
# 보조
# --------------------------------------------------------------------------- #
def _set_state_recursive(widget: tk.Misc, state: str) -> None:
    """하위 위젯을 한꺼번에 켜고 끈다.

    콤보박스는 그냥 'normal' 로 되돌리면 직접 타이핑이 가능해져 버리므로,
    목록에서만 고르도록 'readonly' 로 되돌린다.
    """
    for child in widget.winfo_children():
        target = state
        if state == "normal" and isinstance(child, ttk.Combobox):
            target = "readonly"
        try:
            child.configure(state=target)  # type: ignore[call-arg]
        except (tk.TclError, TypeError):
            pass
        _set_state_recursive(child, state)


def _int_or_none(text: str) -> Optional[int]:
    try:
        return int(str(text).strip())
    except (ValueError, AttributeError):
        return None


def _parse_holidays(text: str) -> list[_dt.date]:
    days: list[_dt.date] = []
    for chunk in str(text or "").replace(";", ",").split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        day = parse_date(chunk)
        if day is not None:
            days.append(day)
    return days


def _open_folder(path: str) -> None:
    """탐색기/파인더로 폴더를 연다 (네트워크 접근 없음)."""
    path = os.path.abspath(path or ".")
    if not os.path.isdir(path):
        messagebox.showinfo("안내", f"폴더가 없습니다:\n{path}")
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:  # noqa: BLE001
        messagebox.showinfo("안내", f"폴더를 열지 못했습니다.\n{path}\n({exc})")


def run(base_dir: Optional[str] = None) -> None:
    """GUI 실행 진입점."""
    base = base_dir or os.getcwd()
    root = tk.Tk()
    root.title("사내 보고서 자동 생성기")
    root.geometry("1080x760")
    root.minsize(900, 640)

    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    elif "clam" in style.theme_names():
        style.theme_use("clam")

    ReportApp(root, base)
    root.mainloop()
