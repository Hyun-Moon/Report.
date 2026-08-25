"""여러 형식으로 들어오는 날짜 값을 ``datetime.date`` 로 정규화한다.

실제 사내 엑셀에서 만나는 형태를 최대한 흡수하는 것이 목표다::

    datetime/date 객체        2026-01-05
    엑셀 일련번호             45657        (1899-12-30 기준)
    'YYYY-MM-DD'              2026-01-05
    'YYYY.MM.DD' / 'YYYY/MM/DD'
    'YYYYMMDD'                20260105
    'MM/DD' / 'M-D'           1/5         (기준 연도 필요)
    'YYYY년 M월 D일' / 'M월 D일'
    '5일' / 5                 (기준 연-월 필요)
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import Optional

__all__ = [
    "EXCEL_EPOCH",
    "parse_date",
    "excel_serial_to_date",
    "month_key",
    "month_label",
    "is_weekend",
]

# 엑셀은 1900년을 윤년으로 잘못 취급하므로 1899-12-30 을 0번으로 두면 1900-03-01
# 이후 날짜가 정확히 맞는다. (그 이전 날짜는 엑셀 자체가 틀리므로 재현하지 않는다.)
EXCEL_EPOCH = _dt.date(1899, 12, 30)

# 엑셀 일련번호로 인정할 범위: 1900-01-01 ~ 9999-12-31
_SERIAL_MIN = 1
_SERIAL_MAX = 2958465

_RE_YMD = re.compile(r"^\s*(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})\s*일?\s*$")
_RE_YM = re.compile(r"^\s*(\d{4})\s*[-./년]\s*(\d{1,2})\s*월?\s*$")
_RE_COMPACT = re.compile(r"^\s*(\d{4})(\d{2})(\d{2})\s*$")
_RE_MD = re.compile(r"^\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})\s*일?\s*$")
_RE_D = re.compile(r"^\s*(\d{1,2})\s*일\s*$")
# 'YYYY-MM-DD HH:MM(:SS)' 같이 시간이 붙은 경우 앞부분만 떼어낸다.
_RE_DATETIME_PREFIX = re.compile(r"^\s*(\d{4}[-./]\d{1,2}[-./]\d{1,2})[ T]\d")
# '2026-07-01(수)' 처럼 끝에 요일/비고가 괄호로 덧붙은 경우 그 부분을 뗀다.
_RE_TRAILING_PAREN = re.compile(r"\s*[(（][^()（）]{1,6}[)）]\s*$")


def excel_serial_to_date(serial: float) -> _dt.date:
    """엑셀 일련번호를 날짜로 변환한다."""
    return EXCEL_EPOCH + _dt.timedelta(days=int(serial))


def parse_date(
    value: object,
    base_year: Optional[int] = None,
    base_month: Optional[int] = None,
) -> Optional[_dt.date]:
    """``value`` 를 날짜로 해석한다. 실패하면 ``None``.

    ``base_year`` / ``base_month`` 는 'MM/DD' 나 '5일' 처럼 정보가 모자란
    표기를 보완하는 데 쓰인다.
    """
    if value is None:
        return None

    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value

    if isinstance(value, bool):  # bool 은 int 의 하위형이라 먼저 걸러낸다.
        return None

    if isinstance(value, (int, float)):
        number = float(value)
        if number != number:  # NaN
            return None
        # 1~31 처럼 작은 수는 일련번호가 아니라 '일(day)' 로 본다.
        if 1 <= number <= 31 and base_year and base_month:
            return _safe_date(base_year, base_month, int(number))
        if _SERIAL_MIN <= number <= _SERIAL_MAX:
            return excel_serial_to_date(number)
        return None

    if not isinstance(value, str):
        return None

    text = value.strip()
    if not text:
        return None

    prefix = _RE_DATETIME_PREFIX.match(text)
    if prefix:
        text = prefix.group(1)
    else:
        text = _RE_TRAILING_PAREN.sub("", text)

    m = _RE_YMD.match(text)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    m = _RE_COMPACT.match(text)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    m = _RE_YM.match(text)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), 1)

    m = _RE_D.match(text)
    if m and base_year and base_month:
        return _safe_date(base_year, base_month, int(m.group(1)))

    m = _RE_MD.match(text)
    if m:
        year = base_year or _dt.date.today().year
        return _safe_date(year, int(m.group(1)), int(m.group(2)))

    # 숫자만 들어 있는 문자열 (일련번호 또는 일)
    if text.replace(".", "", 1).isdigit():
        return parse_date(float(text), base_year, base_month)

    return None


def _safe_date(year: int, month: int, day: int) -> Optional[_dt.date]:
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None


def month_key(day: _dt.date) -> str:
    """``date`` -> ``'2026-01'``."""
    return f"{day.year:04d}-{day.month:02d}"


def month_label(key: str) -> str:
    """``'2026-01'`` -> ``'2026년 1월'``."""
    try:
        year, month = key.split("-")
        return f"{int(year)}년 {int(month)}월"
    except (ValueError, AttributeError):
        return key


def is_weekend(day: _dt.date) -> bool:
    return day.weekday() >= 5
