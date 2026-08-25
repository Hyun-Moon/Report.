"""``templates/`` 폴더에 등록된 템플릿 목록 관리."""

from __future__ import annotations

import os
import shutil
from dataclasses import dataclass
from typing import Optional

from ..errors import FileFormatError
from .base import template_kind

__all__ = ["RegisteredTemplate", "TemplateRegistry"]

_SUPPORTED = (".docx", ".xlsx", ".xlsm")


@dataclass
class RegisteredTemplate:
    name: str
    path: str
    kind: str

    def label(self) -> str:
        return f"[{'워드' if self.kind == 'word' else '엑셀'}] {self.name}"


class TemplateRegistry:
    """``templates/`` 폴더를 훑어 드롭다운에 쓸 목록을 만든다."""

    def __init__(self, directory: str) -> None:
        self.directory = os.path.abspath(directory)
        os.makedirs(self.directory, exist_ok=True)

    def list(self) -> list[RegisteredTemplate]:
        items: list[RegisteredTemplate] = []
        for entry in sorted(os.listdir(self.directory)):
            path = os.path.join(self.directory, entry)
            if not os.path.isfile(path):
                continue
            if entry.startswith("~$"):  # 오피스 임시 파일
                continue
            if os.path.splitext(entry)[1].lower() not in _SUPPORTED:
                continue
            try:
                kind = template_kind(path)
            except FileFormatError:
                continue
            items.append(RegisteredTemplate(entry, path, kind))
        return items

    def labels(self) -> list[str]:
        return [item.label() for item in self.list()]

    def find_by_label(self, label: str) -> Optional[RegisteredTemplate]:
        for item in self.list():
            if item.label() == label:
                return item
        return None

    def register(self, source_path: str, overwrite: bool = True) -> RegisteredTemplate:
        """외부 템플릿을 ``templates/`` 폴더로 복사해 등록한다."""
        kind = template_kind(source_path)
        name = os.path.basename(source_path)
        target = os.path.join(self.directory, name)
        if os.path.abspath(source_path) != os.path.abspath(target):
            if os.path.exists(target) and not overwrite:
                stem, ext = os.path.splitext(name)
                index = 2
                while os.path.exists(target):
                    name = f"{stem}({index}){ext}"
                    target = os.path.join(self.directory, name)
                    index += 1
            shutil.copy2(source_path, target)
        return RegisteredTemplate(name, target, kind)
