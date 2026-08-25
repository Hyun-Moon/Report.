"""사내 보고서 자동 생성기 (완전 로컬 동작).

모듈 구성
---------
``reportgen.data_reader``   원본 엑셀 읽기
``reportgen.dateutils``     여러 형식의 날짜 파싱
``reportgen.aggregator``    일단위 -> 월단위 집계
``reportgen.templating``    워드/엑셀 템플릿 스캔 및 렌더링
``reportgen.mapping``       "엑셀 컬럼 <-> 템플릿 태그" 매핑 저장/로드
``reportgen.generator``     전체 흐름 오케스트레이션
``reportgen.gui``           tkinter 마법사 UI

이 패키지의 어떤 모듈도 네트워크를 사용하지 않는다.
"""

from __future__ import annotations

__version__ = "1.0.0"
__all__ = ["__version__"]
