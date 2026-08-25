#!/usr/bin/env python3
"""PyInstaller 로 단일 exe 를 만든다.

    pip install pyinstaller
    python build_exe.py

``dist/보고서생성기.exe`` 가 만들어진다. templates / mappings / output 폴더는
exe 옆에 두면 되고, 없으면 프로그램이 처음 실행될 때 알아서 만든다.

메모
----
* ``--onefile`` 은 실행할 때마다 임시 폴더에 압축을 푸는 방식이라 시작이 조금
  느리다. 배포가 편한 대신 첫 실행이 3~5초 걸릴 수 있다.
* 사내 백신이 서명 없는 exe 를 막는 경우가 있다. 그럴 때는 ``--onedir`` 로
  만들어 폴더째 배포하는 편이 통과가 잘 된다.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "보고서생성기"


def main() -> int:
    if shutil.which("pyinstaller") is None:
        print("pyinstaller 가 없습니다.  pip install pyinstaller  로 먼저 설치해 주세요.")
        return 1

    onefile = "--onedir" not in sys.argv
    command = [
        "pyinstaller",
        "--noconfirm",
        "--clean",
        "--onefile" if onefile else "--onedir",
        "--windowed",  # 콘솔 창 숨김
        "--name",
        APP_NAME,
        # 이 프로그램은 네트워크를 쓰지 않으므로 관련 모듈을 통째로 뺀다.
        # 용량이 줄고, '외부 통신 없음'이 빌드 산출물 수준에서도 보장된다.
        "--exclude-module", "urllib3",
        "--exclude-module", "requests",
        "--exclude-module", "http",
        "--exclude-module", "email",
        "--exclude-module", "pandas",
        "--exclude-module", "numpy",
        "--exclude-module", "matplotlib",
        "--exclude-module", "PIL",
        "--exclude-module", "pytest",
        os.path.join(ROOT, "app.py"),
    ]

    print(" ".join(command))
    code = subprocess.call(command, cwd=ROOT)
    if code != 0:
        return code

    target = os.path.join(ROOT, "dist")
    for folder in ("templates", "mappings", "output"):
        os.makedirs(os.path.join(target, folder), exist_ok=True)
    # 등록된 템플릿을 함께 배포한다
    source_templates = os.path.join(ROOT, "templates")
    if os.path.isdir(source_templates):
        for name in os.listdir(source_templates):
            src = os.path.join(source_templates, name)
            if os.path.isfile(src) and not name.startswith("~$"):
                shutil.copy2(src, os.path.join(target, "templates", name))

    print(f"\n완료: {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
