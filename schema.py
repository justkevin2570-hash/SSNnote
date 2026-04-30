"""
schema.py — 공문 구조 Pydantic 스키마

Structured Output / 검증에 사용. Field description에 세부 규칙을 포함하여
프롬프트 본문을 줄이면서도 AI가 정확한 형식을 따르도록 유도.
"""

from pydantic import BaseModel, Field


class OfficialDocument(BaseModel):
    body: str = Field(
        description=(
            "공문 본문 전체. 반드시 다음 형식을 정확히 따라야 함:\n"
            "1. 관련: [관련문서번호]\n"
            "2. [제목]을/를 다음과 같이 [보고/운영/발송/제출/공고/지출]하고자 합니다.\n"
            "  가. 일시: [YYYY. M. D.(요일) HH:MM~HH:MM]\n"
            "  나. 장소: [장소]\n"
            "  다. 대상/인원: [대상]\n"
            "  라. 내용: [내용]\n"
            "붙임  [파일명] 1부.  끝.\n\n"
            "규칙:\n"
            "- 날짜는 반드시 '2025. 3. 12.' 형식, 숫자 뒤에 마침표.\n"
            "- 가나다 항목: 앞에 공백 2칸.  1)2)3) 항목: 앞에 공백 4칸.\n"
            "- 문장 끝은 '~함', '~바람', '~고자 함' 공문체 사용.\n"
            "- 마크다운(**·##) 절대 사용 금지.\n"
            "- 표는 HTML <table border='1' style='border-collapse:collapse'> 태그.\n"
            "- 마지막은 반드시 '끝.'으로 끝나야 함."
        )
    )
