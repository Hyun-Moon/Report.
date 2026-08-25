"""GUI 스모크 테스트.

화면을 실제로 띄우고(가상 디스플레이 사용) 1~5단계를 버튼 클릭과 같은 순서로
호출해서, 위젯 구성과 단계 간 상태 전달이 깨지지 않는지 확인한다.

    xvfb-run -a python tests/test_gui_smoke.py
"""

from __future__ import annotations

import os
import sys
import tempfile
import tkinter as tk
from tkinter import messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from reportgen.gui.app import ReportApp  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")
FAILURES: list[str] = []
DIALOGS: list[tuple[str, str]] = []


def _silence_dialogs() -> None:
    """모달 대화상자는 사람이 눌러 줘야 닫히므로 자동 테스트에서는 막아 둔다."""

    def stub(kind: str):
        def _call(title: str = "", message: str = "", *_args, **_kwargs):
            DIALOGS.append((kind, f"{title}: {str(message).splitlines()[0][:70]}"))
            return True

        return _call

    messagebox.showinfo = stub("info")
    messagebox.showerror = stub("error")
    messagebox.showwarning = stub("warning")


def check(condition: bool, label: str) -> None:
    if condition:
        print(f"  OK   {label}")
    else:
        FAILURES.append(label)
        print(f"  FAIL {label}")


def run_case(app: ReportApp, source: str, template: str, aggregate: bool, label: str) -> None:
    print(f"\n[{label}]")
    app.src_path.set(os.path.join(FIXTURES, "sources", source))
    app.sheet_var.set("")
    app.auto_detect.set(True)
    app._load_source()
    check(app.table is not None and app.table.n_rows > 0, "1단계: 원본 읽기")

    app.template_mode.set("upload")
    app.upload_path.set(os.path.join(FIXTURES, "templates", template))
    app.register_after.set(False)
    app._analyze_template()
    check(app.handler is not None, "2단계: 템플릿 분석")
    check(bool(app.bindings), "3단계: 자동 매칭으로 매핑 생성")

    app.use_agg.set(aggregate)
    app._toggle_agg()
    if aggregate:
        app._refresh_aggregation_inputs()
        app._compute_preview()
        check(app.monthly is not None and bool(app.monthly.periods), "4단계: 집계 미리보기")

    app._generate()
    # 생성은 스레드에서 돌기 때문에 완료될 때까지 이벤트 루프를 돌려 준다
    for _ in range(400):
        app.update()
        if str(app.generate_button["state"]) == "normal" and app.result_text.get("1.0", "end").strip():
            break
        app.after(20)
        app.update_idletasks()
    text = app.result_text.get("1.0", "end")
    check("보고서" in text and "만들었습니다" in text, f"5단계: 생성 ({text.strip().splitlines()[-1][:60]})")


def main() -> int:
    if not os.path.isdir(FIXTURES):
        print("먼저 python tests/make_fixtures.py 를 실행해 주세요.", file=sys.stderr)
        return 2

    _silence_dialogs()
    workdir = tempfile.mkdtemp(prefix="reportgen_gui_")
    root = tk.Tk()
    root.geometry("1100x800")
    app = ReportApp(root, workdir)
    app.out_dir.set(os.path.join(workdir, "output"))

    try:
        run_case(app, "S1_헤더1행.xlsx", "W3_병합표.docx", False, "워드 병합표 · 집계 없음")
        run_case(app, "S1_헤더1행.xlsx", "X4_서식유지.xlsx", False, "엑셀 서식유지 · 집계 없음")
        run_case(app, "S5_일단위_한달.xlsx", "W2_단순표.docx", True, "워드 단순표 · 월집계")
        run_case(app, "S6_일단위_여러달.xlsx", "X1_단일시트.xlsx", True, "엑셀 · 여러 달 집계")

        # 매핑 저장/불러오기
        app._save_mapping()
        saved = os.path.join(workdir, "mappings")
        check(bool(os.listdir(saved)), "매핑 JSON 저장")
    finally:
        root.destroy()

    print("\n[뜬 대화상자]")
    for kind, text in DIALOGS:
        print(f"  {kind:7s} {text}")
    errors = [d for d in DIALOGS if d[0] == "error"]
    if errors:
        FAILURES.append(f"오류 대화상자 {len(errors)}건")

    print()
    if FAILURES:
        print(f"실패 {len(FAILURES)}건: {FAILURES}")
        return 1
    print("GUI 스모크 테스트 전부 통과")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
