from datetime import date

import pytest

from doc_ai import (normalize_date, normalize_time, normalize_analysis, parse_json_object,
                    related_line, summary_body, context_block,
                    build_analyze_prompt, summary_text)

TODAY = date(2026, 11, 5)


@pytest.mark.parametrize('raw,expected', [
    ('2026-11-12', '2026-11-12'),
    ('2026. 11. 12.(목)', '2026-11-12'),
    ('2026년 11월 12일', '2026-11-12'),
    ('11. 12.(목)까지', '2026-11-12'),
    ('11월 12일', '2026-11-12'),
    ('1월 10일', '2027-01-10'),     # 연도 없는 지난 날짜 → 내년
    ('10월 20일', '2026-10-20'),    # 최근 지난 날짜는 올해 유지
    ('2026-02-30', ''),             # 없는 날짜
    ('', ''), ('기한 없음', ''), (None, ''),
])
def test_normalize_date(raw, expected):
    assert normalize_date(raw, TODAY) == expected


@pytest.mark.parametrize('raw,expected', [
    ('17:00', '17:00'), ('17시', '17:00'), ('오후 5시 30분', '17:30'), ('9:5', '09:05'),
    ('25:00', ''), ('', ''), ('정오', ''),
])
def test_normalize_time(raw, expected):
    assert normalize_time(raw) == expected


def test_parse_json_object_variants():
    assert parse_json_object('```json\n{"title": "a"}\n```') == {'title': 'a'}
    assert parse_json_object('결과입니다: {"title": "b"} 감사합니다') == {'title': 'b'}
    with pytest.raises(ValueError):
        parse_json_object('JSON 없음')


def test_normalize_analysis_fills_and_cleans():
    a = normalize_analysis({
        'title': ' 2026. 방과후학교 운영 실태 조사 ',
        'doc_number': '경기도교육청 교육과정과-12345',
        'doc_date': '2026. 11. 3.',
        'deadline': '11. 14.(금)', 'deadline_time': '17시',
        'summary': '- 운영 현황 조사\n- 강사 계약 현황 포함',
        'todos': ['1. 조사표 작성', '2) 공문으로 제출', '', 'x', 'y', 'z', 'w'],
        'needs_reply': '예',
    }, TODAY)
    assert a['title'] == '2026. 방과후학교 운영 실태 조사'
    assert a['doc_date'] == '2026-11-03' and a['deadline'] == '2026-11-14' and a['deadline_time'] == '17:00'
    assert a['doc_number'] == '교육과정과-12345(2026. 11. 3.)'          # 기관명 빼고 날짜까지
    assert a['summary'] == ['운영 현황 조사', '강사 계약 현황 포함']
    assert 'todos' not in a                                             # 할 일은 받지 않음
    assert 'attachments' not in a and 'sender' not in a and a['needs_reply'] is True


def test_normalize_analysis_garbage():
    a = normalize_analysis({'title': 3, 'summary': {'x': 1}}, TODAY)
    assert a['title'] == '' and a['summary'] == [] and a['needs_reply'] is False


A = {'title': '2026. 방과후학교 운영 실태 조사', 'doc_number': '교육과정과-12345', 'doc_date': '2026-11-03',
     'deadline': '2026-11-14', 'deadline_time': '17:00',
     'summary': ['운영 현황 조사'], 'needs_reply': True}


def test_summary_helpers():
    assert related_line(A) == '교육과정과-12345(2026. 11. 3.)'
    assert related_line({**A, 'doc_date': ''}) == '교육과정과-12345'
    body = summary_body(A)
    assert '· 운영 현황 조사' in body and '할 일' not in body
    assert '기한: 2026. 11. 14.' in body and '17:00' in body
    assert summary_body({'summary': []}) == ''
    s = summary_text('2026. 방과후학교 운영 실태 조사', '교육과정과-12345(2026. 11. 3.)', body)
    assert s.startswith('2026. 방과후학교 운영 실태 조사 (교육과정과-12345(2026. 11. 3.))\n· 운영 현황 조사')
    assert summary_text('', '', '') == ''
    assert summary_text('', '번호-1', '') == '번호-1'


def test_prompts_contain_context():
    c = context_block('제목', '교육과정과-12345(2026. 11. 3.)', '· 요약')
    assert c.startswith('[참고: 받은 공문]') and '교육과정과-12345' in c and '· 요약' in c
    q = build_analyze_prompt('본문', today=TODAY)
    assert '[공문 내용]\n본문' in q and '2026-11-05' in q
    assert '이미지' in build_analyze_prompt('', today=TODAY)


def test_format_doc_number_official_style():
    from doc_ai import format_doc_number as f
    assert f('경기도교육청 중등교육과-12345', '2026-11-03') == '중등교육과-12345(2026. 11. 3.)'
    assert f('중등교육과-12345(2026. 11. 03.)') == '중등교육과-12345(2026. 11. 3.)'      # 괄호 날짜 우선, 0 없음
    assert f('중등교육과 - 12 345 (2026.11.3)', '2025-01-01') == '중등교육과-12345(2026. 11. 3.)'
    assert f('교육과정과-12345（2026. 1. 5.）') == '교육과정과-12345(2026. 1. 5.)'
    assert f('중등교육과-12345') == '중등교육과-12345'                                   # 날짜 모르면 번호만
    assert f(f('중등교육과-12345', '2026-11-03'), '2026-12-01') == '중등교육과-12345(2026. 11. 3.)'
    assert f('') == '' and f('번호 없음') == '번호 없음'


def test_ocr_doc_number_has_no_zero_padding():
    from capture import _normalize_doc_number
    assert _normalize_doc_number('중 등 교 육 과-12345(2026. 11. 03.)') == '중등교육과-12345(2026. 11. 3.)'
