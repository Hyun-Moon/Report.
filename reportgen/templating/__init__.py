"""템플릿 스캔/렌더링 계층.

워드(docxtpl)와 엑셀(openpyxl)을 같은 인터페이스(:class:`TemplateHandler`)로
감싸서, 상위 계층(generator/GUI)이 두 형식을 구분하지 않고 다룰 수 있게 한다.
"""

from __future__ import annotations

from .base import TemplateHandler, open_template, template_kind
from .excel import ExcelTemplate, infer_label_slots
from .registry import TemplateRegistry
from .word import WordTemplate

__all__ = [
    "TemplateHandler",
    "WordTemplate",
    "ExcelTemplate",
    "TemplateRegistry",
    "open_template",
    "template_kind",
    "infer_label_slots",
]
