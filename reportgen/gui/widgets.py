"""GUI 공통 위젯 조각."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Any, Callable, Iterable, Optional, Sequence

__all__ = [
    "TableView",
    "LabeledEntry",
    "show_error",
    "InlineCombo",
    "InlineEntry",
    "fit_columns",
]


class TableView(ttk.Frame):
    """스크롤바가 달린 ttk.Treeview 래퍼 (표 미리보기용)."""

    def __init__(self, master: tk.Misc, columns: Sequence[str] = (), height: int = 10) -> None:
        super().__init__(master)
        self.tree = ttk.Treeview(self, columns=list(columns), show="headings", height=height)
        y_scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

    def set_columns(self, columns: Sequence[str], widths: Optional[Sequence[int]] = None) -> None:
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(columns)
        for index, name in enumerate(columns):
            self.tree.heading(name, text=name)
            width = widths[index] if widths and index < len(widths) else 120
            self.tree.column(name, width=width, anchor="w", stretch=True)

    def set_rows(self, rows: Iterable[Sequence[Any]]) -> None:
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert("", "end", values=["" if v is None else v for v in row])

    def load_matrix(self, matrix: Sequence[Sequence[Any]], max_rows: int = 300) -> None:
        """첫 행을 헤더로 보고 표 전체를 갈아 끼운다."""
        if not matrix:
            self.set_columns(["(비어 있음)"])
            self.set_rows([])
            return
        header = [str(v) if v is not None else "" for v in matrix[0]]
        self.set_columns(header, fit_columns(matrix))
        self.set_rows(matrix[1 : max_rows + 1])

    def selected_index(self) -> Optional[int]:
        selection = self.tree.selection()
        if not selection:
            return None
        return self.tree.index(selection[0])

    def selected_values(self) -> Optional[list[Any]]:
        selection = self.tree.selection()
        if not selection:
            return None
        return list(self.tree.item(selection[0], "values"))


def fit_columns(matrix: Sequence[Sequence[Any]], minimum: int = 70, maximum: int = 260) -> list[int]:
    """내용 길이를 보고 적당한 컬럼 너비를 정한다."""
    if not matrix:
        return []
    width_count = max(len(row) for row in matrix)
    widths: list[int] = []
    for index in range(width_count):
        longest = 0
        for row in matrix[:60]:
            value = row[index] if index < len(row) else ""
            longest = max(longest, len(str(value if value is not None else "")))
        widths.append(max(minimum, min(maximum, 12 + longest * 9)))
    return widths


class LabeledEntry(ttk.Frame):
    """라벨 + 입력칸 한 줄."""

    def __init__(
        self,
        master: tk.Misc,
        label: str,
        width: int = 30,
        label_width: int = 12,
        variable: Optional[tk.Variable] = None,
    ) -> None:
        super().__init__(master)
        ttk.Label(self, text=label, width=label_width, anchor="w").pack(side="left")
        self.var = variable or tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self.var, width=width)
        self.entry.pack(side="left", fill="x", expand=True)

    def get(self) -> str:
        return self.var.get()

    def set(self, value: str) -> None:
        self.var.set(value)


class InlineCombo:
    """Treeview 셀 위에 콤보박스를 띄워 값을 고르게 하는 편집기."""

    def __init__(
        self,
        tree: ttk.Treeview,
        column: str,
        values: Sequence[str],
        on_commit: Callable[[str, str], None],
    ) -> None:
        self.tree = tree
        self.column = column
        self.values = list(values)
        self.on_commit = on_commit
        self.combo: Optional[ttk.Combobox] = None
        self.item: str = ""
        tree.bind("<Double-1>", self._open, add="+")
        tree.bind("<Return>", self._open, add="+")

    def _open(self, event: tk.Event) -> None:
        self._close()
        item = self.tree.focus()
        if not item:
            return
        if getattr(event, "x", None) is not None and event.type == tk.EventType.ButtonPress:
            if self.tree.identify_column(event.x) != self.column:
                return
        box = self.tree.bbox(item, self.column)
        if not box:
            return
        x, y, width, height = box
        self.item = item
        combo = ttk.Combobox(self.tree, values=self.values, state="readonly")
        current = self.tree.set(item, self.column)
        if current in self.values:
            combo.set(current)
        combo.place(x=x, y=y, width=width, height=height)
        combo.focus_set()
        combo.bind("<<ComboboxSelected>>", self._commit)
        combo.bind("<Escape>", lambda _e: self._close())
        combo.bind("<FocusOut>", lambda _e: self._close())
        self.combo = combo

    def _commit(self, _event: tk.Event) -> None:
        if self.combo is None:
            return
        value = self.combo.get()
        item = self.item
        self._close()
        self.tree.set(item, self.column, value)
        self.on_commit(item, value)

    def _close(self) -> None:
        if self.combo is not None:
            self.combo.destroy()
            self.combo = None


class InlineEntry:
    """Treeview 셀을 그 자리에서 직접 고쳐 쓰게 하는 편집기 (집계값 수동 보정)."""

    def __init__(
        self,
        tree: ttk.Treeview,
        on_commit: Callable[[str, str, str], None],
        editable_columns: Optional[Sequence[str]] = None,
    ) -> None:
        self.tree = tree
        self.on_commit = on_commit
        self.editable_columns = list(editable_columns) if editable_columns else None
        self.entry: Optional[ttk.Entry] = None
        self.item = ""
        self.column = ""
        tree.bind("<Double-1>", self._open, add="+")

    def _open(self, event: tk.Event) -> None:
        self._close()
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item or not column:
            return
        if self.editable_columns is not None and column not in self.editable_columns:
            return
        box = self.tree.bbox(item, column)
        if not box:
            return
        x, y, width, height = box
        self.item, self.column = item, column
        entry = ttk.Entry(self.tree)
        entry.insert(0, self.tree.set(item, column))
        entry.select_range(0, "end")
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus_set()
        entry.bind("<Return>", self._commit)
        entry.bind("<Escape>", lambda _e: self._close())
        entry.bind("<FocusOut>", self._commit)
        self.entry = entry

    def _commit(self, _event: tk.Event) -> None:
        if self.entry is None:
            return
        value = self.entry.get()
        item, column = self.item, self.column
        self._close()
        self.tree.set(item, column, value)
        self.on_commit(item, column, value)

    def _close(self) -> None:
        if self.entry is not None:
            self.entry.destroy()
            self.entry = None


def show_error(title: str, error: BaseException) -> None:
    """예외를 사용자용 메시지 박스로 띄운다."""
    from tkinter import messagebox

    from ..errors import ReportGenError

    if isinstance(error, ReportGenError):
        messagebox.showerror(title, str(error))
    else:
        messagebox.showerror(
            title,
            f"예상치 못한 오류가 발생했습니다.\n\n{type(error).__name__}: {error}",
        )
