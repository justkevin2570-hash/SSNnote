from datetime import date

import pytest

from doc_ai import (normalize_date, normalize_time, normalize_analysis, parse_json_object,
                    related_line, task_deadline, task_notes, reply_title, build_reply_prompt,
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
    assert a['todos'][:2] == ['조사표 작성', '공문으로 제출'] and len(a['todos']) == 5
    assert 'attachments' not in a and 'sender' not in a and a['needs_reply'] is True


def test_normalize_analysis_garbage():
    a = normalize_analysis({'title': 3, 'summary': {'x': 1}}, TODAY)
    assert a['title'] == '' and a['summary'] == [] and a['needs_reply'] is False


A = {'title': '2026. 방과후학교 운영 실태 조사', 'doc_number': '교육과정과-12345', 'doc_date': '2026-11-03',
     'deadline': '2026-11-14', 'deadline_time': '17:00',
     'summary': ['운영 현황 조사'], 'todos': ['조사표 작성'], 'needs_reply': True}


def test_task_helpers():
    assert related_line(A) == '교육과정과-12345(2026. 11. 3.)'
    assert related_line({**A, 'doc_date': ''}) == '교육과정과-12345'
    assert task_deadline(A) == '2026-11-14 17:00'
    assert task_deadline({**A, 'deadline_time': ''}) == '2026-11-14'
    assert task_deadline({**A, 'deadline': ''}) == ''
    notes = task_notes(A)
    assert '☐ 조사표 작성' in notes and '- 운영 현황 조사' in notes and '발신' not in notes and '붙임' not in notes
    s = summary_text(A)
    assert s.startswith('2026. 방과후학교 운영 실태 조사 (교육과정과-12345(2026. 11. 3.))') and '기한: 2026. 11. 14. 17:00' in s


def test_reply_title():
    assert reply_title(A) == '2026. 방과후학교 운영 실태 제출'
    assert reply_title({'title': '연수 참가 신청 안내'}) == '연수 참가 신청 제출'
    assert reply_title({'title': '학교폭력 실태 결과 보고'}) == '학교폭력 실태 결과 보고'
    assert reply_title({'title': ''}) == ''


def test_prompts_contain_context():
    p = build_reply_prompt(A, profile_text='- 학교명: 서서니고', instruction='인원은 3명')
    assert '교육과정과-12345(2026. 11. 3.)' in p and '서서니고' in p and '인원은 3명' in p and '"body"' in p
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
