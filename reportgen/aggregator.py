"""일단위 데이터를 월단위로 집계한다.

핵심 개념
---------
``AggregationSpec``  어느 컬럼이 날짜이고, 항목별로 어떤 방식(합산/평균/최대/최소)
                     으로 묶을지에 대한 설정.
``MonthlyResult``    ``[연-월][컬럼] -> 값`` 형태의 집계 결과. 미리보기 화면과
                     보고서 생성이 모두 이 결과만 바라본다.

빠진 날짜(예: 28일치만 있음)는 그냥 없는 대로 계산한다. 평균은 '실제로 값이
있는 날' 기준이며, 며칠치가 들어갔는지는 ``day_counts`` 로 확인할 수 있다.
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import Any, Iterable, Optional

from .data_reader import Table
from .dateutils import month_key, month_label, parse_date
from .errors import AggregationError

__all__ = [
    "METHODS",
    "METHOD_LABELS",
    "AggregationSpec",
    "MonthlyResult",
    "detect_date_column",
    "aggregate_monthly",
    "suggest_methods",
    "to_number",
]

METHODS = ("sum", "mean", "max", "min", "count", "first", "last", "text_join")

METHOD_LABELS = {
    "sum": "합산",
    "mean": "평균",
    "max": "최댓값",
    "min": "최솟값",
    "count": "건수",
    "first": "첫 값",
    "last": "마지막 값",
    "text_join": "문자열 이어붙이기",
}

# 컬럼 이름으로 집계 방식을 추천할 때 쓰는 힌트
_HINTS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("mean", ("온도", "습도", "평균", "율", "률", "효율", "단가", "지수", "농도", "역률")),
    ("max", ("최대", "피크", "최고", "peak", "max")),
    ("min", ("최소", "최저", "min")),
    ("sum", ("사용량", "합계", "금액", "요금", "수량", "발생", "생산", "판매", "건수", "실적")),
)

_RE_NUMBER = re.compile(r"^-?[\d,]*\.?\d+$")


# --------------------------------------------------------------------------- #
# 설정 / 결과
# --------------------------------------------------------------------------- #
@dataclass
class AggregationSpec:
    date_column: Optional[str] = None
    methods: dict[str, str] = field(default_factory=dict)
    default_method: str = "sum"
    exclude_weekends: bool = False
    exclude_dates: list[_dt.date] = field(default_factory=list)
    #: 'YYYY-MM' 목록. 비어 있으면 전체 월을 대상으로 한다.
    only_months: list[str] = field(default_factory=list)
    #: 'MM/DD' 나 '5일' 처럼 연/월 정보가 없는 표기를 보완한다.
    base_year: Optional[int] = None
    base_month: Optional[int] = None
    #: 여러 달이 섞여 있을 때: 'separate' 는 월별 보고서, 'wide' 는 한 보고서에 월별 컬럼
    multi_month_mode: str = "separate"

    def method_for(self, column: str) -> str:
        return self.methods.get(column, self.default_method)


@dataclass
class MonthlyResult:
    """월별 집계 결과."""

    periods: list[str]  # 'YYYY-MM' 오름차순
    columns: list[str]  # 날짜 컬럼을 제외한 값 컬럼
    values: dict[str, dict[str, Any]] = field(default_factory=dict)
    day_counts: dict[str, int] = field(default_factory=dict)
    used_days: dict[str, list[_dt.date]] = field(default_factory=dict)
    skipped_rows: int = 0
    date_column: str = ""

    def get(self, period: str, column: str) -> Any:
        return self.values.get(period, {}).get(column)

    def set(self, period: str, column: str, value: Any) -> None:
        """미리보기 화면에서 사용자가 손으로 고친 값을 반영한다."""
        self.values.setdefault(period, {})[column] = value

    def as_matrix(self) -> list[list[Any]]:
        """미리보기 표: 첫 행이 헤더."""
        head = ["연-월", "일수"] + list(self.columns)
        rows = [head]
        for period in self.periods:
            row: list[Any] = [month_label(period), self.day_counts.get(period, 0)]
            row.extend(self.get(period, col) for col in self.columns)
            rows.append(row)
        return rows

    def context_for(self, period: str) -> dict[str, Any]:
        """한 달치를 템플릿 컨텍스트(태그 -> 값)로 펼친다."""
        data = dict(self.values.get(period, {}))
        data["연-월"] = period
        data["대상월"] = month_label(period)
        data["집계일수"] = self.day_counts.get(period, 0)
        return data

    def wide_context(self) -> dict[str, Any]:
        """여러 달을 한 보고서에 넣을 때: ``'사용량 (2026-01)'`` 형태로 펼친다."""
        data: dict[str, Any] = {}
        for period in self.periods:
            for col in self.columns:
                data[f"{col} ({period})"] = self.get(period, col)
        data["대상월"] = " · ".join(month_label(p) for p in self.periods)
        data["연-월"] = ", ".join(self.periods)
        data["집계일수"] = sum(self.day_counts.values())
        return data


# --------------------------------------------------------------------------- #
# 날짜 컬럼 탐지
# --------------------------------------------------------------------------- #
_DATE_NAME_HINTS = ("날짜", "일자", "일시", "년월", "연월", "date", "day", "월일", "일")


def _name_hints_date(name: str) -> bool:
    lowered = str(name).lower()
    return any(hint.lower() in lowered for hint in _DATE_NAME_HINTS)


def detect_date_column(table: Table, spec: Optional[AggregationSpec] = None) -> Optional[str]:
    """날짜로 해석되는 값이 가장 많은 컬럼을 고른다.

    주의할 점: ``12500`` 같은 평범한 숫자도 엑셀 일련번호로는 해석이 된다.
    그래서 **맨숫자 컬럼은 이름에 날짜 힌트가 있을 때만** 날짜 후보로 본다.
    그러지 않으면 '사용량' 컬럼이 날짜로 잡히는 사고가 난다.
    """
    spec = spec or AggregationSpec()
    best_name: Optional[str] = None
    best_score = 0.0

    for index, name in enumerate(table.columns):
        values = [row[index] if index < len(row) else None for row in table.rows]
        non_empty = [v for v in values if v is not None and str(v).strip() != ""]
        if not non_empty:
            continue

        hinted = _name_hints_date(name)
        evidence = 0
        for value in non_empty:
            if isinstance(value, _dt.date):  # datetime 도 date 의 하위형
                evidence += 1
                continue
            if isinstance(value, str):
                # 순수 숫자 문자열은 아래 '맨숫자' 규칙을 따른다
                if not value.strip().replace(".", "", 1).lstrip("-").isdigit():
                    if parse_date(value, spec.base_year, spec.base_month) is not None:
                        evidence += 1
                    continue
            # 여기부터는 맨숫자(또는 숫자만 든 문자열)
            if not hinted:
                continue
            if parse_date(value, spec.base_year, spec.base_month) is not None:
                evidence += 1

        ratio = evidence / len(non_empty)
        score = ratio + (0.15 if hinted else 0.0)
        if ratio >= 0.6 and score > best_score:
            best_score = score
            best_name = name

    return best_name


def suggest_methods(table: Table, date_column: Optional[str]) -> dict[str, str]:
    """컬럼 이름을 보고 집계 방식을 추천한다."""
    out: dict[str, str] = {}
    for name in table.columns:
        if name == date_column:
            continue
        lowered = name.lower()
        chosen = None
        for method, keywords in _HINTS:
            if any(k.lower() in lowered for k in keywords):
                chosen = method
                break
        if chosen is None:
            chosen = "sum" if _looks_numeric(table, name) else "first"
        out[name] = chosen
    return out


def _looks_numeric(table: Table, column: str) -> bool:
    values = table.column_values(column)
    non_empty = [v for v in values if v is not None and str(v).strip() != ""]
    if not non_empty:
        return False
    numeric = sum(1 for v in non_empty if to_number(v) is not None)
    return numeric / len(non_empty) >= 0.7


# --------------------------------------------------------------------------- #
# 집계
# --------------------------------------------------------------------------- #
def aggregate_monthly(table: Table, spec: AggregationSpec) -> MonthlyResult:
    """일단위 ``table`` 을 월단위로 묶는다."""
    date_column = spec.date_column or detect_date_column(table, spec)
    if not date_column:
        raise AggregationError(
            "날짜 컬럼을 찾지 못했습니다.",
            "집계 화면에서 날짜가 들어 있는 컬럼을 직접 지정해 주세요. "
            "'1일, 2일...' 처럼 일자만 있는 표는 기준 연도/월도 함께 입력해야 합니다.",
        )
    if date_column not in table.columns:
        raise AggregationError(
            f"'{date_column}' 컬럼이 원본 데이터에 없습니다.",
            "원본 시트를 바꿨다면 집계 설정을 다시 확인해 주세요.",
        )

    value_columns = [c for c in table.columns if c != date_column]
    excluded = set(spec.exclude_dates)

    buckets: dict[str, dict[str, list[Any]]] = {}
    days: dict[str, set[_dt.date]] = {}
    skipped = 0
    date_index = table.index_of(date_column)

    for row in table.rows:
        raw = row[date_index] if date_index < len(row) else None
        day = parse_date(raw, spec.base_year, spec.base_month)
        if day is None:
            skipped += 1
            continue
        if spec.exclude_weekends and day.weekday() >= 5:
            continue
        if day in excluded:
            continue
        key = month_key(day)
        if spec.only_months and key not in spec.only_months:
            continue

        bucket = buckets.setdefault(key, {c: [] for c in value_columns})
        days.setdefault(key, set()).add(day)
        for col in value_columns:
            idx = table.index_of(col)
            bucket[col].append(row[idx] if idx < len(row) else None)

    if not buckets:
        raise AggregationError(
            "집계할 수 있는 날짜 데이터가 없습니다.",
            f"'{date_column}' 컬럼의 값이 날짜로 인식되지 않았거나, "
            "제외 조건(주말/공휴일/대상월)이 모든 행을 걸러냈습니다.",
        )

    periods = sorted(buckets)
    values: dict[str, dict[str, Any]] = {}
    for period in periods:
        values[period] = {
            col: _apply(spec.method_for(col), buckets[period][col])
            for col in value_columns
        }

    return MonthlyResult(
        periods=periods,
        columns=value_columns,
        values=values,
        day_counts={p: len(days.get(p, ())) for p in periods},
        used_days={p: sorted(days.get(p, ())) for p in periods},
        skipped_rows=skipped,
        date_column=date_column,
    )


def _apply(method: str, raw_values: Iterable[Any]) -> Any:
    values = list(raw_values)
    if method == "count":
        return sum(1 for v in values if v is not None and str(v).strip() != "")
    if method == "text_join":
        texts = [str(v).strip() for v in values if v is not None and str(v).strip() != ""]
        return ", ".join(dict.fromkeys(texts))
    if method in ("first", "last"):
        filled = [v for v in values if v is not None and str(v).strip() != ""]
        if not filled:
            return None
        return filled[0] if method == "first" else filled[-1]

    numbers = [n for n in (to_number(v) for v in values) if n is not None]
    if not numbers:
        return None
    if method == "sum":
        return _tidy(sum(numbers))
    if method == "mean":
        return _tidy(sum(numbers) / len(numbers))
    if method == "max":
        return _tidy(max(numbers))
    if method == "min":
        return _tidy(min(numbers))
    raise AggregationError(
        f"알 수 없는 집계 방식입니다: {method}",
        "합산/평균/최댓값/최솟값/건수 중에서 골라 주세요.",
    )


def to_number(value: Any) -> Optional[float]:
    """엑셀 셀 값을 숫자로 바꾼다. 숫자가 아니면 ``None``."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return None if number != number else number  # NaN 제외
    if isinstance(value, _dt.date):
        return None
    text = str(value).strip()
    if not text:
        return None
    # '1,234', '1 234', '123 kWh', '45%' 같은 표기를 흡수한다
    text = text.replace(",", "").replace(" ", "")
    percent = text.endswith("%")
    if percent:
        text = text[:-1]
    text = re.sub(r"[^\d.\-+eE]", "", text)
    if not text or not _RE_NUMBER.match(text.replace("+", "")):
        try:
            number = float(text)
        except ValueError:
            return None
    else:
        try:
            number = float(text)
        except ValueError:
            return None
    if percent:
        number /= 100.0
    return number


def _tidy(number: float) -> float | int:
    """부동소수점 찌꺼기(0.30000000000000004)를 정리한다."""
    rounded = round(number, 6)
    if abs(rounded - round(rounded)) < 1e-9:
        return int(round(rounded))
    return rounded
