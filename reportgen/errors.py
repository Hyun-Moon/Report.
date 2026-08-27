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


class FormulaCacheError(ReportGenError):
    """수식 셀에 '계산된 값' 캐시가 없어 읽을 수 없는 경우.

    엑셀 파일은 수식 자체와 별개로 마지막으로 계산된 결과값을 셀에 함께
    저장해 둔다. 이 프로그램은 그 결과값만 읽으므로, 프로그램이 만들었거나
    LibreOffice 등에서 재계산 없이 저장된 파일은 수식 칸이 비어 보인다.
    """


class MappingError(ReportGenError):
    """매핑이 비어 있거나 존재하지 않는 컬럼을 가리키는 경우."""


class TemplateError(ReportGenError):
    """템플릿 파일 자체가 잘못된 경우."""


class AggregationError(ReportGenError):
    """집계 설정이 잘못되었거나 날짜를 인식하지 못한 경우."""


class MultiFileError(ReportGenError):
    """'여러 파일 모아 월간표 만들기'(하루 1파일 취합) 관련 오류."""
