"""템플릿 공통 인터페이스."""

from __future__ import annotations

import os
import re
from abc import ABC, abstractmethod
from typing import Any, Optional

from ..errors import FileFormatError, TemplateError
from ..mapping import TemplateSlot

__all__ = [
    "TemplateHandler",
    "TAG_RE",
    "TABLE_ANCHOR_RE",
    "SIMPLE_TAG_RE",
    "template_kind",
    "open_template",
    "is_simple_tag",
]

#: 문서/셀 안에서 찾는 플레이스홀더. ``{{ 태그명 }}``
TAG_RE = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)

#: 표 반복 시작점. ``{{#표}}`` 또는 ``{{#table}}``
TABLE_ANCHOR_RE = re.compile(r"^\s*#\s*(?:표|table)\s*$", re.IGNORECASE)

#: '단순 태그'로 인정하는 이름. jinja 문법 문자가 없으면 단순 태그로 본다.
#: 괄호는 ``{{사용량 (2026-01)}}`` 같은 실제 태그 이름에 자주 쓰이므로 허용한다.
SIMPLE_TAG_RE = re.compile(r"^[^{}%|.\[\]!<>=*/,'\"]+$")

#: ``round(1)`` 처럼 '이름 바로 뒤에 여는 괄호' 면 함수 호출로 본다.
_CALL_RE = re.compile(r"[\w가-힣]\(")

_JINJA_WORDS = {"if", "for", "endif", "endfor", "else", "elif", "in", "not", "and", "or"}


def is_simple_tag(expression: str) -> bool:
    """``{{...}}`` 안의 내용이 '그냥 이름'인지 판단한다.

    ``사용량`` / ``연-월`` / ``사용량 (2026-01)`` -> True (매핑 항목으로 노출)
    ``r.사용량`` / ``값|round(1)`` / ``round(1)`` -> False (사용자가 쓴 jinja 식)
    """
    text = (expression or "").strip()
    if not text:
        return False
    if not SIMPLE_TAG_RE.match(text):
        return False
    if _CALL_RE.search(text):
        return False
    if text.split()[0].lower() in _JINJA_WORDS:
        return False
    return True


class TemplateHandler(ABC):
    """워드/엑셀 템플릿의 공통 동작."""

    #: 'word' 또는 'excel'
    kind: str = ""
    #: 결과 파일 확장자
    extension: str = ""

    def __init__(self, path: str) -> None:
        if not path or not os.path.isfile(path):
            raise FileFormatError(f"템플릿 파일을 찾을 수 없습니다: {path}")
        self.path = os.path.abspath(path)
        self.name = os.path.basename(path)
        self.stem = os.path.splitext(self.name)[0]

    @abstractmethod
    def scan(self) -> list[TemplateSlot]:
        """템플릿 안의 태그/셀 목록을 훑는다."""

    @abstractmethod
    def render(self, context: dict[str, Any], output_path: str, table_data: Optional[dict] = None) -> str:
        """``context`` 를 채워 ``output_path`` 에 저장하고 경로를 돌려준다."""

    def describe(self) -> str:
        return f"{self.name} ({'워드' if self.kind == 'word' else '엑셀'} 템플릿)"


def template_kind(path: str) -> str:
    ext = os.path.splitext(path or "")[1].lower()
    if ext == ".docx":
        return "word"
    if ext in (".xlsx", ".xlsm"):
        return "excel"
    if ext == ".doc":
        raise FileFormatError(
            "구형 .doc 템플릿은 지원하지 않습니다.",
            "워드에서 '다른 이름으로 저장' -> .docx 로 변환한 뒤 사용해 주세요.",
        )
    if ext == ".xls":
        raise FileFormatError(
            "구형 .xls 템플릿은 지원하지 않습니다.",
            "엑셀에서 '다른 이름으로 저장' -> .xlsx 로 변환한 뒤 사용해 주세요.",
        )
    raise FileFormatError(
        f"'{ext or '확장자 없음'}' 은(는) 템플릿으로 쓸 수 없습니다.",
        ".docx 또는 .xlsx 파일을 골라 주세요.",
    )


def open_template(path: str) -> TemplateHandler:
    """확장자를 보고 알맞은 핸들러를 만든다."""
    from .excel import ExcelTemplate
    from .word import WordTemplate

    kind = template_kind(path)
    handler: TemplateHandler = WordTemplate(path) if kind == "word" else ExcelTemplate(path)
    try:
        handler.scan()
    except TemplateError:
        raise
    except Exception as exc:  # noqa: BLE001 - 사용자에게 원인을 알려주기 위함
        raise TemplateError(
            f"템플릿을 여는 중 문제가 생겼습니다: {os.path.basename(path)}",
            f"파일이 손상되었거나 암호가 걸려 있을 수 있습니다. ({exc})",
        ) from exc
    return handler
