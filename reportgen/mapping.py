"""'엑셀 컬럼 <-> 템플릿 태그' 매핑 모델과 JSON 저장/로드.

매핑 파일은 ``mappings/<템플릿파일명>.json`` 에 저장되며, 같은 템플릿을 다시
쓸 때 자동으로 불러온다.
"""

from __future__ import annotations

import datetime as _dt
import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Optional

from .aggregator import MonthlyResult
from .data_reader import Table
from .dateutils import month_label
from .errors import MappingError

__all__ = [
    "Binding",
    "MappingProfile",
    "TemplateSlot",
    "BUILTIN_KEYS",
    "auto_match",
    "normalize_text",
    "mapping_path_for",
    "load_mapping",
    "save_mapping",
    "resolve_context",
]

BUILTIN_KEYS = {
    "오늘": "오늘 날짜 (2026-08-25)",
    "현재일시": "생성 일시 (2026-08-25 14:03)",
    "생성일시": "생성 일시 (2026-08-25 14:03)",
    "원본파일": "원본 엑셀 파일명",
    "템플릿명": "템플릿 파일명",
    "대상월": "집계 대상 연-월 (2026년 1월)",
    "연-월": "집계 대상 키 (2026-01)",
    "집계일수": "집계에 들어간 날짜 수",
    "작성자": "윈도우 로그인 계정명",
}


@dataclass
class Binding:
    """템플릿 슬롯 하나에 어떤 값을 넣을지에 대한 규칙."""

    source: str = "blank"  # 'column' | 'literal' | 'builtin' | 'blank'
    column: str = ""
    #: 집계를 쓰지 않을 때 원본 표의 몇 번째 행을 쓸지 (0-based). None 이면 첫 행.
    row: Optional[int] = None
    literal: str = ""
    builtin: str = ""
    #: 특정 달로 못박을 때 쓰는 'YYYY-MM'. 비우면 생성 중인 달을 따라간다.
    period: str = ""
    #: 파이썬 format 스펙. 예: ',.0f' -> 1,234 / '.1f' -> 12.3
    number_format: str = ""
    #: 값 뒤에 붙일 단위 문자열. 예: ' kWh'
    suffix: str = ""

    def to_json(self) -> dict[str, Any]:
        out = {"source": self.source}
        for key in ("column", "literal", "builtin", "period", "number_format", "suffix"):
            value = getattr(self, key)
            if value:
                out[key] = value
        if self.row is not None:
            out["row"] = self.row
        return out

    @classmethod
    def from_json(cls, data: dict[str, Any]) -> "Binding":
        allowed = set(cls.__annotations__)
        return cls(**{k: v for k, v in (data or {}).items() if k in allowed})

    def describe(self) -> str:
        if self.source == "column":
            base = f"컬럼: {self.column}"
            if self.period:
                base += f" [{self.period}]"
            if self.row is not None:
                base += f" ({self.row + 1}행)"
            return base
        if self.source == "literal":
            return f"고정값: {self.literal}"
        if self.source == "builtin":
            return f"자동: {self.builtin}"
        return "(비움)"


@dataclass
class TemplateSlot:
    """템플릿에서 찾아낸 '값이 들어갈 자리'."""

    key: str
    kind: str = "tag"  # 'tag' | 'cell'
    where: str = ""
    sample: str = ""
    occurrences: int = 1
    #: 자동 매칭에 쓸 텍스트가 key 와 다를 때만 채운다. ``{{태그}}`` 는 태그
    #: 이름이 곧 key 라서 비워 두지만, 태그가 없는 엑셀 서식에서 라벨 옆 빈
    #: 칸을 추론한 경우는 key 가 셀 좌표(``'Sheet1'!B3``)이므로 매칭에는
    #: 이 필드(라벨 텍스트)를 대신 쓴다.
    match_text: str = ""

    def to_json(self) -> dict[str, Any]:
        return dict(self.__dict__)


@dataclass
class MappingProfile:
    template_name: str = ""
    template_type: str = ""  # 'word' | 'excel'
    bindings: dict[str, Binding] = field(default_factory=dict)
    #: 집계 설정 스냅샷 (다음에 열 때 그대로 복원)
    aggregation: dict[str, Any] = field(default_factory=dict)
    #: 원본 읽기 설정 스냅샷
    read_options: dict[str, Any] = field(default_factory=dict)
    saved_at: str = ""
    note: str = ""

    def to_json(self) -> dict[str, Any]:
        return {
            "version": 1,
            "template_name": self.template_name,
            "template_type": self.template_type,
            "bindings": {k: v.to_json() for k, v in self.bindings.items()},
            "aggregation": self.aggregation,
            "read_options": self.read_options,
            "saved_at": self.saved_at or _dt.datetime.now().isoformat(timespec="seconds"),
            "note": self.note,
        }

    @classmethod
    def from_json(cls, data: dict[str, Any]) -> "MappingProfile":
        return cls(
            template_name=data.get("template_name", ""),
            template_type=data.get("template_type", ""),
            bindings={
                key: Binding.from_json(value)
                for key, value in (data.get("bindings") or {}).items()
            },
            aggregation=data.get("aggregation") or {},
            read_options=data.get("read_options") or {},
            saved_at=data.get("saved_at", ""),
            note=data.get("note", ""),
        )

    def filled_keys(self) -> list[str]:
        return [k for k, b in self.bindings.items() if b.source != "blank"]


# --------------------------------------------------------------------------- #
# 자동 매칭
# --------------------------------------------------------------------------- #
_NORMALIZE = re.compile(r"[\s_\-()\[\]{}·.]+")
#: ``사용량 (2026-01)`` 처럼 태그 이름 끝에 달이 붙어 있는 형태
_PERIOD_SUFFIX = re.compile(r"^(?P<base>.*?)\s*[(\[]\s*(?P<period>\d{4}-\d{2})\s*[)\]]\s*$")


def _norm(text: str) -> str:
    return _NORMALIZE.sub("", str(text)).lower()


def normalize_text(text: str) -> str:
    """``_norm`` 의 공개 버전. 다른 모듈이 '완전 일치인지'를 직접 비교할 때 쓴다."""
    return _norm(text)


def auto_match(slots: list[TemplateSlot], columns: list[str]) -> dict[str, Binding]:
    """태그 이름과 컬럼 이름이 비슷하면 자동으로 연결한다."""
    by_norm = {_norm(c): c for c in columns}
    result: dict[str, Binding] = {}

    for slot in slots:
        key = slot.key
        #: 실제로 이름을 비교할 텍스트. 보통은 key 와 같지만(=태그 이름),
        #: 라벨 추론으로 만들어진 슬롯은 key 가 셀 좌표라서 따로 둔 라벨
        #: 텍스트(match_text)를 대신 비교해야 한다.
        match_key = slot.match_text or key

        # 1) 내장 키 먼저
        if match_key in BUILTIN_KEYS:
            result[key] = Binding(source="builtin", builtin=match_key)
            continue

        # 2) '사용량 (2026-01)' 처럼 달이 못박힌 태그는 그 달의 집계값으로 연결
        suffix = _PERIOD_SUFFIX.match(match_key)
        if suffix:
            base = by_norm.get(_norm(suffix.group("base")))
            if base:
                result[key] = Binding(
                    source="column", column=base, period=suffix.group("period")
                )
                continue

        normalized = _norm(match_key)
        # 2) 완전 일치
        if normalized in by_norm:
            result[key] = Binding(source="column", column=by_norm[normalized])
            continue
        # 3) 부분 일치 (가장 긴 것 우선)
        candidates = [
            original
            for norm_col, original in by_norm.items()
            if norm_col and (norm_col in normalized or normalized in norm_col)
        ]
        if candidates:
            candidates.sort(key=len, reverse=True)
            result[key] = Binding(source="column", column=candidates[0])
            continue
        result[key] = Binding(source="blank")

    return result


# --------------------------------------------------------------------------- #
# 저장 / 로드
# --------------------------------------------------------------------------- #
def mapping_path_for(template_path: str, mapping_dir: str) -> str:
    stem = os.path.splitext(os.path.basename(template_path))[0]
    safe = re.sub(r'[\\/:*?"<>|]+', "_", stem)
    return os.path.join(mapping_dir, f"{safe}.json")


def save_mapping(profile: MappingProfile, path: str) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    profile.saved_at = _dt.datetime.now().isoformat(timespec="seconds")
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(profile.to_json(), handle, ensure_ascii=False, indent=2)
    return os.path.abspath(path)


def load_mapping(path: str) -> Optional[MappingProfile]:
    if not path or not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as handle:
            return MappingProfile.from_json(json.load(handle))
    except (json.JSONDecodeError, OSError, TypeError) as exc:
        raise MappingError(
            f"매핑 파일을 읽지 못했습니다: {os.path.basename(path)}",
            f"파일이 손상되었을 수 있습니다. 삭제 후 다시 매핑해 주세요. ({exc})",
        ) from exc


# --------------------------------------------------------------------------- #
# 값 채우기
# --------------------------------------------------------------------------- #
def resolve_context(
    slots: list[TemplateSlot],
    bindings: dict[str, Binding],
    table: Table,
    monthly: Optional[MonthlyResult] = None,
    period: Optional[str] = None,
    meta: Optional[dict[str, Any]] = None,
    strict: bool = False,
    pool: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    """슬롯별로 실제 값을 계산해 ``{키: 값}`` 컨텍스트를 만든다.

    ``pool`` 은 '이름이 그대로 맞으면 알아서 채워 주는' 값 모음이다. 예를 들어
    월별 컬럼 나열 모드에서 ``{{사용량 (2026-01)}}`` 같은 태그는 따로 매핑하지
    않아도 집계 결과에서 바로 값을 찾아 넣는다.
    """
    meta = dict(meta or {})
    pool = pool or {}
    context: dict[str, Any] = {}
    missing: list[str] = []

    for slot in slots:
        binding = bindings.get(slot.key, Binding())
        try:
            value = _resolve_one(binding, table, monthly, period, meta)
        except KeyError:
            raise MappingError(
                f"'{slot.key}' 에 연결된 컬럼 '{binding.column}' 이(가) 원본 데이터에 없습니다.",
                "원본 엑셀을 바꾸었다면 3단계 매핑 화면에서 다시 지정해 주세요.",
            ) from None

        if binding.source == "blank":
            if slot.key in pool:
                value = pool[slot.key]
            elif slot.key in BUILTIN_KEYS:
                value = _builtin_value(slot.key, table, monthly, period, meta)
            else:
                missing.append(slot.key)
                value = ""
        context[slot.key] = _format_value(value, binding)

    if strict and missing:
        raise MappingError(
            "아직 값이 연결되지 않은 항목이 있습니다: " + ", ".join(missing[:10]),
            "3단계 매핑 화면에서 각 항목에 컬럼이나 고정값을 지정해 주세요.",
        )
    return context


def _resolve_one(
    binding: Binding,
    table: Table,
    monthly: Optional[MonthlyResult],
    period: Optional[str],
    meta: dict[str, Any],
) -> Any:
    if binding.source == "literal":
        return binding.literal
    if binding.source == "builtin":
        return _builtin_value(binding.builtin, table, monthly, period, meta)
    if binding.source != "column" or not binding.column:
        return None

    # 집계 결과가 있고 해당 컬럼이 집계 대상이면, 원본 행이 아니라 집계값을 쓴다.
    if monthly is not None and binding.row is None and binding.column in monthly.columns:
        target = binding.period or period
        if target:
            return monthly.get(target, binding.column)
        if len(monthly.periods) == 1:
            return monthly.get(monthly.periods[0], binding.column)
        # 여러 달을 한 보고서에 담는 경우: 어느 달인지 알 수 없으므로 전부 나열한다.
        # (원본 첫 행 값을 슬쩍 넣으면 '한 달치 값'처럼 보여 오해를 부른다.)
        return " / ".join(
            "" if monthly.get(p, binding.column) is None else str(monthly.get(p, binding.column))
            for p in monthly.periods
        )

    if binding.column not in table.columns:
        raise KeyError(binding.column)
    row_index = binding.row if binding.row is not None else 0
    if not table.rows:
        return None
    if row_index < 0 or row_index >= table.n_rows:
        return None
    return table.cell(row_index, binding.column)


def _builtin_value(
    name: str,
    table: Table,
    monthly: Optional[MonthlyResult],
    period: Optional[str],
    meta: dict[str, Any],
) -> Any:
    now = _dt.datetime.now()
    if name == "오늘":
        return now.strftime("%Y-%m-%d")
    if name in ("현재일시", "생성일시"):
        return now.strftime("%Y-%m-%d %H:%M")
    if name == "원본파일":
        return os.path.basename(table.source_path or meta.get("source_file", ""))
    if name == "템플릿명":
        return meta.get("template_name", "")
    if name == "작성자":
        return meta.get("author") or os.environ.get("USERNAME") or os.environ.get("USER") or ""
    if name == "대상월":
        if period:
            return month_label(period)
        if monthly and monthly.periods:
            return " · ".join(month_label(p) for p in monthly.periods)
        return month_label(now.strftime("%Y-%m"))  # 집계를 안 쓸 때는 이번 달
    if name == "연-월":
        if period:
            return period
        if monthly and monthly.periods:
            return ", ".join(monthly.periods)
        return now.strftime("%Y-%m")
    if name == "집계일수":
        if monthly is None:
            return table.n_rows
        if period:
            return monthly.day_counts.get(period, 0)
        return sum(monthly.day_counts.values())
    return meta.get(name, "")


def _format_value(value: Any, binding: Binding) -> Any:
    if value is None:
        return ""
    if binding.number_format:
        try:
            value = format(float(value), binding.number_format)
        except (TypeError, ValueError):
            pass
    if binding.suffix:
        value = f"{value}{binding.suffix}"
    return value
