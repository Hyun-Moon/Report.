"""사용자 친화적인 예외 정의.

GUI 계층은 :class:`ReportGenError` 만 붙잡아서 그대로 메시지 박스에 띄우면 된다.
따라서 모든 메시지는 개발자가 아니라 '사무실에서 쓰는 사람'이 읽을 것을 전제로 쓴다.
"""

from __future__ import annotations


class ReportGenError(Exception):
    """프로그램이 사용자에게 그대로 보여줘도 되는 오류."""

    def __init__(self, message: str, hint: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.hint = hint

    def __str__(self) -> str:  # pragma: no cover - 단순 포맷
        if self.hint:
            return f"{self.message}\n\n[해결 방법] {self.hint}"
        return self.message


class FileFormatError(ReportGenError):
    """확장자가 다르거나 열 수 없는 파일."""


class SheetNotFoundError(ReportGenError):
    """지정한 시트가 없는 경우."""


class CellRangeError(ReportGenError):
    """셀 범위 문자열이 잘못된 경우."""


class HeaderError(ReportGenError):
    """헤더를 인식하지 못한 경우."""


class MappingError(ReportGenError):
    """매핑이 비어 있거나 존재하지 않는 컬럼을 가리키는 경우."""


class TemplateError(ReportGenError):
    """템플릿 파일 자체가 잘못된 경우."""


class AggregationError(ReportGenError):
    """집계 설정이 잘못되었거나 날짜를 인식하지 못한 경우."""
