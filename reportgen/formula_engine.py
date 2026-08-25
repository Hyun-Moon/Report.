"""엑셀 수식 계산 - ``formulas`` 패키지로 워크북 전체를 계산한다.

'계산된 값 캐시가 없는' 엑셀 파일(프로그램이 만들었거나, 재계산 없이
저장된 경우)에서, 다른 시트를 참조하는 수식이나 VLOOKUP 같은 복잡한
함수를 포함해 **엑셀 수식 대부분을 실제로 계산**해서 값을 채운다.

``formulas`` 는 셀 하나씩이 아니라 워크북 전체를 계산 그래프로 만들어
한 번에 계산한다. 그래서 시트 간 참조도 자연스럽게 풀리고, VLOOKUP·
INDEX·MATCH 같은 함수도 대부분 지원한다. 대신 계산 자체가 가볍지 않으므로
(파일 크기에 따라 1초 안팎), 캐시 없는 셀이 하나라도 있을 때만 호출하고
한 번 계산한 결과는 그대로 재사용한다.

그래도 정말로 못 푸는 수식(참조가 아예 깨졌거나, 순환 참조 등)은
``#REF!``, ``#NAME?`` 같은 오류값으로 돌아온다. 이런 값은 진짜 값으로
쓰지 않고 '계산 실패'로 취급해서, 상위 코드가 기존 안전장치
(:class:`~reportgen.errors.FormulaCacheError` 또는 '빈 값으로 넘어가기'
옵션)로 처리하게 넘긴다.
"""

from __future__ import annotations

import os

# formulas 패키지가 내부적으로 tqdm 진행률 표시줄을 콘솔에 찍는데, GUI
# 프로그램에서는 지저분하기만 하므로 끈다. import 전에 설정해야 먹는다.
os.environ.setdefault("TQDM_DISABLE", "1")

import re  # noqa: E402
from typing import Any, Optional  # noqa: E402

from openpyxl.cell.cell import is_date_format  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.utils.datetime import from_excel  # noqa: E402

__all__ = ["WorkbookFormulaEngine"]

#: 엑셀이 계산 실패 시 셀에 표시하는 오류 값 패턴. 진짜 값으로 취급하면 안 된다.
#: (일반 텍스트가 우연히 '#'로 시작하는 경우와 헷갈리지 않도록 정확한 오류
#: 표기만 매칭한다 - 예: '#1위' 같은 라벨은 오류가 아니다.)
_ERROR_RE = re.compile(r"^#(REF|NAME|DIV/0|VALUE|N/A|NULL|NUM|SPILL|CALC)[!?]?$")


class WorkbookFormulaEngine:
    """워크북 전체의 수식을 ``formulas`` 로 한 번에 계산해서 캐시한다.

    ``get(sheet_name, row, col, number_format)`` 로 셀 값을 조회한다.
    계산은 첫 조회 때 딱 한 번만 하고(무겁다), 이후 조회는 메모리에
    캐시된 결과에서 바로 꺼낸다.
    """

    def __init__(self, path: str) -> None:
        self.path = path
        self._basename = os.path.basename(path)
        self._solution: Optional[dict] = None
        self._load_error: Optional[Exception] = None
        #: '이 셀이 진짜 수식인지' 가볍게 확인하는 용도 (계산은 안 하는
        #: data_only=False 워크북 - formulas 의 무거운 전체 계산과는 별개).
        self._formula_workbook = None
        #: 값을 못 구한 (sheet, row, col) 목록. 상위 코드가 이걸로
        #: FormulaCacheError / 경고 메시지를 만든다.
        self.unresolved: set[tuple[str, int, int]] = set()

    def get(self, sheet_name: str, row: int, col: int, number_format: str = "") -> Any:
        if not self._is_formula_cell(sheet_name, row, col):
            return None  # 진짜 빈 셀 - 오류도 경고도 낼 일이 아니다

        self._ensure_solved()
        if self._solution is None:
            self.unresolved.add((sheet_name, row, col))
            return None

        coordinate = f"{get_column_letter(col)}{row}"
        # formulas 패키지는 내부적으로 시트 이름을 대문자로 바꿔 키를 만든다
        # (한글처럼 대소문자가 없는 문자는 .upper() 를 해도 그대로다).
        key = f"'[{self._basename}]{sheet_name.upper()}'!{coordinate}"
        entry = self._solution.get(key)
        if entry is None:
            self.unresolved.add((sheet_name, row, col))
            return None

        value = _unwrap(getattr(entry, "value", entry))
        if value is _UNRESOLVED:
            self.unresolved.add((sheet_name, row, col))
            return None

        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if is_date_format(number_format):
                try:
                    return from_excel(value)
                except (ValueError, OverflowError):
                    pass  # 날짜로 못 바꾸면 그냥 숫자로 둔다
        return value

    def close(self) -> None:
        if self._formula_workbook is not None:
            self._formula_workbook.close()
            self._formula_workbook = None

    # ------------------------------------------------------------------ #
    def _is_formula_cell(self, sheet_name: str, row: int, col: int) -> bool:
        if self._formula_workbook is None:
            from openpyxl import load_workbook

            self._formula_workbook = load_workbook(self.path, data_only=False, read_only=False)
        sheet = self._formula_workbook[sheet_name]
        return sheet.cell(row=row, column=col).data_type == "f"

    def _ensure_solved(self) -> None:
        if self._solution is not None or self._load_error is not None:
            return
        try:
            import formulas  # 무거운 의존성(scipy 포함)이라 실제로 쓸 때만 불러온다

            model = formulas.ExcelModel().loads(self.path).finish()
            self._solution = model.calculate()
        except Exception as exc:  # noqa: BLE001 - 계산 자체가 실패하면 전부 포기하고 안전망에 맡긴다
            self._load_error = exc
            self._solution = None


def _unwrap(raw: Any) -> Any:
    """``formulas`` 가 돌려주는 numpy 배열/스칼라를 파이썬 기본형으로 바꾼다."""
    try:
        import numpy as np
    except ImportError:  # pragma: no cover - formulas 의 필수 의존성이라 보통 없을 일 없음
        np = None  # type: ignore[assignment]

    if np is not None and isinstance(raw, np.ndarray):
        if raw.size == 0:
            return _UNRESOLVED
        raw = raw.reshape(-1)[0]
    if hasattr(raw, "item"):
        try:
            raw = raw.item()
        except (ValueError, TypeError):
            pass
    if isinstance(raw, str):
        text = raw.strip().upper()
        if _ERROR_RE.match(text):
            return _UNRESOLVED
    return raw


_UNRESOLVED = object()
