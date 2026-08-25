"""전체 흐름 오케스트레이션: 읽기 -> 집계 -> 매핑 -> 렌더링.

GUI 는 :class:`GenerationRequest` 하나를 채워서 :func:`generate` 를 부르기만
하면 된다. 콘솔에서 배치로 돌릴 때도 같은 함수를 쓴다.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
from dataclasses import dataclass, field
from typing import Any, Optional

from .aggregator import AggregationSpec, MonthlyResult, aggregate_monthly
from .data_reader import ReadOptions, Table, read_table
from .dateutils import month_label
from .errors import ReportGenError
from .mapping import Binding, TemplateSlot, resolve_context
from .templating import TemplateHandler, open_template

__all__ = [
    "GenerationRequest",
    "GenerationResult",
    "generate",
    "prepare",
    "build_output_name",
]

_INVALID_NAME = re.compile(r'[\\/:*?"<>|]+')


@dataclass
class GenerationRequest:
    source_path: str = ""
    read_options: ReadOptions = field(default_factory=ReadOptions)
    template_path: str = ""
    bindings: dict[str, Binding] = field(default_factory=dict)
    use_aggregation: bool = False
    aggregation: AggregationSpec = field(default_factory=AggregationSpec)
    output_dir: str = "output"
    #: 'separate' = 월별로 보고서 각각, 'wide' = 한 보고서에 월별 컬럼 나열
    multi_month_mode: str = "separate"
    #: ``{{#표}}`` / ``{% for r in rows %}`` 에 넣을 표 데이터를 만들지 여부
    include_table: bool = True
    #: 표에 넣을 컬럼 (비우면 전부)
    table_columns: list[str] = field(default_factory=list)
    #: 미리보기 화면에서 사람이 손으로 고친 값 {'2026-01': {'사용량': 123}}
    overrides: dict[str, dict[str, Any]] = field(default_factory=dict)
    author: str = ""


@dataclass
class GenerationResult:
    files: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    monthly: Optional[MonthlyResult] = None
    table: Optional[Table] = None

    def summary(self) -> str:
        if not self.files:
            return "생성된 파일이 없습니다."
        lines = [f"보고서 {len(self.files)}개를 만들었습니다.", ""]
        lines.extend(f"  · {path}" for path in self.files)
        if self.warnings:
            lines.append("")
            lines.append("[참고]")
            lines.extend(f"  - {w}" for w in self.warnings)
        return "\n".join(lines)


@dataclass
class Prepared:
    """생성 직전 상태. 미리보기 화면이 이 값을 그대로 보여준다."""

    table: Table
    monthly: Optional[MonthlyResult]
    handler: TemplateHandler
    slots: list[TemplateSlot]
    warnings: list[str] = field(default_factory=list)


def prepare(request: GenerationRequest) -> Prepared:
    """읽기 + 집계 + 템플릿 스캔까지만 수행한다 (파일은 만들지 않는다)."""
    warnings: list[str] = []
    table = read_table(request.source_path, request.read_options)
    warnings.extend(table.warnings)
    if not table.rows:
        warnings.append("원본 표에서 데이터 행을 찾지 못했습니다. 셀 범위를 확인해 주세요.")

    monthly: Optional[MonthlyResult] = None
    if request.use_aggregation:
        monthly = aggregate_monthly(table, request.aggregation)
        _apply_overrides(monthly, request.overrides)
        if monthly.skipped_rows:
            warnings.append(
                f"날짜로 읽지 못한 행 {monthly.skipped_rows}개는 집계에서 제외했습니다."
            )
        for period in monthly.periods:
            count = monthly.day_counts.get(period, 0)
            expected = _days_in_month(period)
            if count < expected:
                warnings.append(
                    f"{month_label(period)}: {expected}일 중 {count}일치만 집계되었습니다."
                )

    handler = open_template(request.template_path)
    slots = handler.scan()
    if not slots:
        warnings.append(
            f"템플릿 '{handler.name}' 에서 {{{{태그}}}} 를 찾지 못했습니다. "
            "매핑 화면에서 셀 좌표(예: Sheet1!B3)를 직접 추가할 수 있습니다."
            if handler.kind == "excel"
            else f"템플릿 '{handler.name}' 에서 {{{{태그}}}} 를 찾지 못했습니다."
        )

    return Prepared(table=table, monthly=monthly, handler=handler, slots=slots, warnings=warnings)


def generate(request: GenerationRequest, prepared: Optional[Prepared] = None) -> GenerationResult:
    """보고서 파일을 실제로 만든다."""
    prepared = prepared or prepare(request)
    table, monthly, handler, slots = (
        prepared.table,
        prepared.monthly,
        prepared.handler,
        prepared.slots,
    )
    result = GenerationResult(warnings=list(prepared.warnings), monthly=monthly, table=table)

    meta = {
        "template_name": handler.name,
        "source_file": os.path.basename(table.source_path),
        "author": request.author,
    }

    # 매핑 키 중에 템플릿 스캔에 안 잡힌 것(엑셀 셀 좌표 직접 지정)도 슬롯으로 추가
    effective_slots = list(slots)
    known = {s.key for s in slots}
    for key, binding in request.bindings.items():
        if key not in known and binding.source != "blank":
            effective_slots.append(TemplateSlot(key=key, kind="cell", where=key))

    timestamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    periods: list[Optional[str]]
    if monthly is None:
        periods = [None]
    elif request.multi_month_mode == "wide" or len(monthly.periods) <= 1:
        periods = [monthly.periods[0]] if len(monthly.periods) == 1 else [None]
    else:
        periods = list(monthly.periods)

    wide = request.multi_month_mode == "wide" and monthly is not None and len(monthly.periods) > 1

    for period in periods:
        pool: dict[str, Any] = {}
        if monthly is not None:
            pool.update(monthly.wide_context() if wide else {})
            if period:
                pool.update(
                    {f"{c} ({period})": monthly.get(period, c) for c in monthly.columns}
                )

        context = resolve_context(
            effective_slots,
            request.bindings,
            table,
            monthly=monthly,
            period=period,
            meta=meta,
            pool=pool,
        )
        table_data = _build_table_data(request, table, monthly, period, wide)
        output_path = os.path.join(
            request.output_dir,
            build_output_name(handler, timestamp, period if len(periods) > 1 else None),
        )
        result.files.append(handler.render(context, output_path, table_data))

    return result


def build_output_name(
    handler: TemplateHandler, timestamp: str, period: Optional[str] = None
) -> str:
    """``[템플릿명]_[생성일시].docx`` (월별 생성 시 연-월을 덧붙인다)."""
    stem = _INVALID_NAME.sub("_", handler.stem)
    suffix = f"_{period}" if period else ""
    return f"{stem}{suffix}_{timestamp}{handler.extension}"


# --------------------------------------------------------------------------- #
# 보조
# --------------------------------------------------------------------------- #
def _build_table_data(
    request: GenerationRequest,
    table: Table,
    monthly: Optional[MonthlyResult],
    period: Optional[str],
    wide: bool,
) -> dict[str, Any]:
    if not request.include_table:
        return {"columns": [], "rows": [], "include_header": True}

    if monthly is not None:
        matrix = monthly.as_matrix()
        columns = matrix[0]
        rows = matrix[1:]
        if period and not wide:
            rows = [
                row
                for row, key in zip(rows, monthly.periods)
                if key == period
            ]
        return {"columns": columns, "rows": rows, "include_header": True}

    wanted = request.table_columns or table.columns
    missing = [c for c in wanted if c not in table.columns]
    if missing:
        wanted = [c for c in wanted if c in table.columns]
    indexes = [table.index_of(c) for c in wanted]
    rows = [[row[i] if i < len(row) else None for i in indexes] for row in table.rows]
    return {"columns": wanted, "rows": rows, "include_header": True}


def _apply_overrides(monthly: MonthlyResult, overrides: dict[str, dict[str, Any]]) -> None:
    for period, columns in (overrides or {}).items():
        for column, value in (columns or {}).items():
            monthly.set(period, column, value)


def _days_in_month(period: str) -> int:
    import calendar

    try:
        year, month = (int(x) for x in period.split("-"))
        return calendar.monthrange(year, month)[1]
    except (ValueError, TypeError):
        return 0


def run_from_profile(
    source_path: str,
    template_path: str,
    profile,
    output_dir: str = "output",
) -> GenerationResult:
    """저장된 매핑 프로필로 바로 생성한다 (배치/명령줄용)."""
    from .mapping import MappingProfile

    if not isinstance(profile, MappingProfile):
        raise ReportGenError("매핑 프로필이 올바르지 않습니다.")

    read_options = ReadOptions(**(profile.read_options or {}))
    request = GenerationRequest(
        source_path=source_path,
        read_options=read_options,
        template_path=template_path,
        bindings=profile.bindings,
        output_dir=output_dir,
    )
    agg = profile.aggregation or {}
    if agg.get("enabled"):
        request.use_aggregation = True
        request.aggregation = AggregationSpec(
            date_column=agg.get("date_column"),
            methods=agg.get("methods") or {},
            default_method=agg.get("default_method", "sum"),
            exclude_weekends=bool(agg.get("exclude_weekends")),
            base_year=agg.get("base_year"),
            base_month=agg.get("base_month"),
        )
        request.multi_month_mode = agg.get("multi_month_mode", "separate")
    return generate(request)
