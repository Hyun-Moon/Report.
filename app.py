#!/usr/bin/env python3
"""사내 보고서 자동 생성기 진입점.

GUI 실행::

    python app.py

명령줄에서 저장된 매핑으로 바로 생성 (배치용)::

    python app.py --cli --source 원본.xlsx --template templates/양식.docx
"""

from __future__ import annotations

import argparse
import os
import sys


def base_directory() -> str:
    """실행 파일(exe) 로 묶였을 때도 templates/ 폴더를 제대로 찾도록 한다."""
    if getattr(sys, "frozen", False):  # PyInstaller
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="사내 보고서 자동 생성기")
    parser.add_argument("--cli", action="store_true", help="GUI 없이 명령줄로 생성")
    parser.add_argument("--source", help="원본 엑셀 파일")
    parser.add_argument("--template", help="템플릿 파일 (.docx / .xlsx)")
    parser.add_argument("--mapping", help="매핑 JSON 경로 (생략하면 mappings/ 에서 자동 탐색)")
    parser.add_argument("--out", default=None, help="저장 폴더")
    args = parser.parse_args(argv)

    base = base_directory()

    if not args.cli:
        from reportgen.gui import run

        run(base)
        return 0

    if not args.source or not args.template:
        parser.error("--cli 모드에서는 --source 와 --template 이 필요합니다.")

    from reportgen.errors import ReportGenError
    from reportgen.generator import run_from_profile
    from reportgen.mapping import load_mapping, mapping_path_for

    mapping_dir = os.path.join(base, "mappings")
    mapping_path = args.mapping or mapping_path_for(args.template, mapping_dir)

    try:
        profile = load_mapping(mapping_path)
        if profile is None:
            print(f"매핑 파일이 없습니다: {mapping_path}", file=sys.stderr)
            print("GUI 에서 한 번 매핑을 저장한 뒤 다시 실행해 주세요.", file=sys.stderr)
            return 2
        result = run_from_profile(
            args.source, args.template, profile, args.out or os.path.join(base, "output")
        )
    except ReportGenError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(result.summary())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
