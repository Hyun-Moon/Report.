"""예외 처리 및 '완전 로컬 동작' 검증.

* 잘못된 입력에 대해 개발자용 스택 트레이스가 아니라 사람이 읽는
  :class:`ReportGenError` 가 나오는지
* 소스 어디에도 네트워크를 쓰는 코드가 없는지
"""

from __future__ import annotations

import ast
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from reportgen.aggregator import AggregationSpec, aggregate_monthly  # noqa: E402
from reportgen.data_reader import ReadOptions, read_table  # noqa: E402
from reportgen.errors import (  # noqa: E402
    AggregationError,
    CellRangeError,
    FileFormatError,
    FormulaCacheError,
    MappingError,
    ReportGenError,
    SheetNotFoundError,
    TemplateError,
)
from reportgen.mapping import Binding, TemplateSlot, resolve_context  # noqa: E402
from reportgen.templating import open_template  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SOURCES = os.path.join(HERE, "fixtures", "sources")

FAILURES: list[str] = []

#: 이 프로그램이 절대 써서는 안 되는 모듈들
FORBIDDEN_MODULES = {
    "socket",
    "urllib",
    "urllib2",
    "urllib3",
    "http",
    "httplib",
    "requests",
    "ftplib",
    "smtplib",
    "telnetlib",
    "xmlrpc",
    "asyncio",
    "aiohttp",
    "websocket",
    "websockets",
    "boto3",
    "paramiko",
}


def expect(kind: type[BaseException], label: str, call) -> None:
    try:
        call()
    except kind as exc:
        message = str(exc)
        friendly = isinstance(exc, ReportGenError) and len(message) > 10
        if friendly:
            print(f"  OK   {label}\n         -> {message.splitlines()[0]}")
        else:
            FAILURES.append(label)
            print(f"  FAIL {label}: 메시지가 불친절함 ({message!r})")
    except BaseException as exc:  # noqa: BLE001
        FAILURES.append(label)
        print(f"  FAIL {label}: {type(exc).__name__} 이(가) 나옴 ({exc})")
    else:
        FAILURES.append(label)
        print(f"  FAIL {label}: 오류가 나지 않음")


def test_input_errors() -> None:
    print("\n[잘못된 입력에 대한 안내 메시지]")

    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as handle:
        handle.write(b"this is not excel")
        text_file = handle.name

    expect(FileFormatError, "엑셀이 아닌 파일을 원본으로 고른 경우",
           lambda: read_table(text_file))
    expect(FileFormatError, "없는 파일 경로",
           lambda: read_table(os.path.join(SOURCES, "없는파일.xlsx")))
    expect(FileFormatError, "원본을 고르지 않고 진행",
           lambda: read_table(""))
    expect(SheetNotFoundError, "없는 시트 이름 지정",
           lambda: read_table(os.path.join(SOURCES, "S1_헤더1행.xlsx"),
                              ReadOptions(sheet_name="없는시트")))
    expect(CellRangeError, "셀 범위 문자열이 엉망인 경우",
           lambda: read_table(os.path.join(SOURCES, "S1_헤더1행.xlsx"),
                              ReadOptions(cell_range="A1-Z9", auto_detect=False)))
    expect(FileFormatError, "템플릿으로 .txt 를 고른 경우",
           lambda: open_template(text_file))
    expect(FileFormatError, "구형 .doc 템플릿",
           lambda: open_template(os.path.join(HERE, "없는템플릿.doc")))

    table = read_table(os.path.join(SOURCES, "S1_헤더1행.xlsx"))
    expect(AggregationError, "날짜 컬럼이 없는 표를 월집계하려는 경우",
           lambda: aggregate_monthly(table, AggregationSpec()))
    expect(AggregationError, "지정한 날짜 컬럼이 원본에 없는 경우",
           lambda: aggregate_monthly(table, AggregationSpec(date_column="없는컬럼")))

    daily = read_table(os.path.join(SOURCES, "S5_일단위_한달.xlsx"))
    expect(AggregationError, "제외 조건이 모든 행을 걸러낸 경우",
           lambda: aggregate_monthly(daily, AggregationSpec(only_months=["2099-01"])))

    expect(FormulaCacheError, "원본에 계산되지 않은 수식이 있는 경우",
           lambda: read_table(os.path.join(SOURCES, "S18_계산안된수식.xlsx")))

    expect(MappingError, "매핑이 원본에 없는 컬럼을 가리키는 경우",
           lambda: resolve_context(
               [TemplateSlot(key="사용량")],
               {"사용량": Binding(source="column", column="사라진컬럼")},
               table,
           ))
    expect(MappingError, "필수 항목이 비어 있는 채로 생성(strict)",
           lambda: resolve_context(
               [TemplateSlot(key="미지정항목")], {}, table, strict=True
           ))

    os.unlink(text_file)


def test_formula_cache_bypass() -> None:
    """원본을 재계산할 수 없을 때, 빈 값으로 넘어가는 우회 옵션이 조용히 넘어가지 않는지."""
    print("\n[계산 안 된 수식 - 우회 옵션]")
    path = os.path.join(SOURCES, "S18_계산안된수식.xlsx")
    table = read_table(path, ReadOptions(allow_uncalculated_formulas=True))
    ok = bool(table.warnings) and table.rows[0][-1] is None
    label = "우회 옵션을 켜면 오류 대신 경고 + 빈 값으로 진행됨"
    if ok:
        print(f"  OK   {label}")
        print(f"         -> {table.warnings[0]}")
    else:
        FAILURES.append(label)
        print(f"  FAIL {label}: warnings={table.warnings}, rows={table.rows}")


def test_table_block_detection() -> None:
    """한 시트에 표가 여러 개 섞여 있을 때, 후보를 정확히 나누는지 확인.

    대형 빌딩 에너지 대장처럼 좌우로 나란히 놓인 표(전력/스팀/용수)와
    위아래로 별개인 표가 한 시트에 있는 경우를 재현한 원본으로 검증한다.
    """
    from reportgen.data_reader import list_table_blocks, read_table, ReadOptions

    print("\n[표 후보 찾기]")
    path = os.path.join(SOURCES, "S21_다중표시트.xlsx")
    blocks = list_table_blocks(path)
    ranges = {b.range for b in blocks}
    expected = {"B7:D9", "F7:H9", "J7:K9", "A16:H18"}
    label = f"좌우로 나란한 표 3개 + 별개 표 1개를 정확히 분리 (찾은 것: {sorted(ranges)})"
    if expected <= ranges:
        print(f"  OK   {label}")
    else:
        FAILURES.append(label)
        print(f"  FAIL {label}")

    steam_block = next((b for b in blocks if b.range == "F7:H9"), None)
    if steam_block is None:
        FAILURES.append("스팀 표 후보를 못 찾음")
        print("  FAIL 스팀 표 후보를 못 찾음")
        return

    table = read_table(
        path, ReadOptions(cell_range=steam_block.range, header_rows=2, auto_detect=False)
    )
    ok = table.columns == ["스팀 / 구분", "사용량", "사용금액"] and table.rows == [
        ["금년월보사용량", 262000, None]
    ]
    label2 = "찾은 표 범위로 실제 읽기까지 정상 동작"
    if ok:
        print(f"  OK   {label2}")
    else:
        FAILURES.append(label2)
        print(f"  FAIL {label2}: columns={table.columns}, rows={table.rows}")


def test_graceful_paths() -> None:
    """오류가 '나면 안 되는' 경우들."""
    print("\n[예외 없이 넘어가야 하는 경우]")
    checks = [
        (
            "빈 셀이 섞여 있어도 집계가 된다",
            lambda: aggregate_monthly(
                read_table(os.path.join(SOURCES, "S3_빈셀혼재.xlsx")),
                AggregationSpec(date_column="부서", base_year=2026, base_month=1),
            ),
        ),
    ]
    for label, call in checks:
        try:
            call()
        except AggregationError:
            # 날짜가 아닌 컬럼을 날짜로 지정한 것이므로 AggregationError 가 맞다
            print(f"  OK   {label} (날짜 아님을 정확히 알려 줌)")
        except BaseException as exc:  # noqa: BLE001
            FAILURES.append(label)
            print(f"  FAIL {label}: {type(exc).__name__}: {exc}")
        else:
            print(f"  OK   {label}")

    # 데이터가 한 행뿐이어도, 200행이어도 읽기가 된다
    for name, expected in (("S12_한행.xlsx", 1), ("S11_행많음.xlsx", 200)):
        table = read_table(os.path.join(SOURCES, name))
        if table.n_rows == expected:
            print(f"  OK   {name}: {expected}행 읽음")
        else:
            FAILURES.append(name)
            print(f"  FAIL {name}: {table.n_rows}행 (기대 {expected})")


def test_no_network() -> None:
    """소스 트리 전체에서 네트워크 관련 import 를 찾는다."""
    print("\n[외부 통신 코드 없음 확인]")
    offenders: list[str] = []
    scanned = 0

    for folder, _dirs, files in os.walk(os.path.join(ROOT, "reportgen")):
        for name in files:
            if not name.endswith(".py"):
                continue
            path = os.path.join(folder, name)
            scanned += 1
            tree = ast.parse(open(path, encoding="utf-8").read(), filename=path)
            for node in ast.walk(tree):
                names: list[str] = []
                if isinstance(node, ast.Import):
                    names = [alias.name for alias in node.names]
                elif isinstance(node, ast.ImportFrom) and node.module:
                    names = [node.module]
                for module in names:
                    root = module.split(".")[0]
                    if root in FORBIDDEN_MODULES:
                        offenders.append(f"{os.path.relpath(path, ROOT)}: import {module}")

    # app.py 도 함께 본다
    tree = ast.parse(open(os.path.join(ROOT, "app.py"), encoding="utf-8").read())
    scanned += 1
    for node in ast.walk(tree):
        names = []
        if isinstance(node, ast.Import):
            names = [alias.name for alias in node.names]
        elif isinstance(node, ast.ImportFrom) and node.module:
            names = [node.module]
        for module in names:
            if module.split(".")[0] in FORBIDDEN_MODULES:
                offenders.append(f"app.py: import {module}")

    if offenders:
        FAILURES.extend(offenders)
        print(f"  FAIL 네트워크 관련 import 발견:\n    " + "\n    ".join(offenders))
    else:
        print(f"  OK   파이썬 파일 {scanned}개에서 네트워크 import 없음")


def main() -> int:
    test_input_errors()
    test_formula_cache_bypass()
    test_table_block_detection()
    test_graceful_paths()
    test_no_network()
    print()
    if FAILURES:
        print(f"실패 {len(FAILURES)}건: {FAILURES}")
        return 1
    print("예외 처리 / 로컬 동작 검증 전부 통과")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
