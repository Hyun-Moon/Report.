#!/usr/bin/env python3
"""전체 검증 실행기.

    python run_tests.py            # 샘플 생성 + 양식 검증 + 예외 검증
    python run_tests.py --gui      # GUI 스모크 테스트까지 (화면이 필요)

GUI 테스트는 tkinter 와 디스플레이가 있어야 한다. 리눅스 서버라면
``xvfb-run -a python run_tests.py --gui`` 로 돌리면 된다.
"""

from __future__ import annotations

import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))


def step(title: str, argv: list[str]) -> int:
    print("\n" + "=" * 70)
    print(f"  {title}")
    print("=" * 70)
    return subprocess.call([sys.executable, *argv], cwd=ROOT)


def main() -> int:
    want_gui = "--gui" in sys.argv
    failures: list[str] = []

    if step("샘플 원본·템플릿 생성", ["tests/make_fixtures.py"]) != 0:
        failures.append("샘플 생성")
    if step("에너지 도메인 샘플 생성", ["tests/make_energy_fixtures.py"]) != 0:
        failures.append("에너지 샘플 생성")
    if step("양식 커버리지 검증", ["tests/test_matrix.py"]) != 0:
        failures.append("양식 커버리지")
    if step("에너지 도메인 문서 검증", ["tests/test_energy_matrix.py"]) != 0:
        failures.append("에너지 도메인 문서")
    if step("예외 처리 / 로컬 동작 검증", ["tests/test_errors.py"]) != 0:
        failures.append("예외 처리")
    if want_gui and step("GUI 스모크 테스트", ["tests/test_gui_smoke.py"]) != 0:
        failures.append("GUI")

    print("\n" + "=" * 70)
    if failures:
        print(f"실패한 단계: {', '.join(failures)}")
        return 1
    print("전체 검증 통과")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
