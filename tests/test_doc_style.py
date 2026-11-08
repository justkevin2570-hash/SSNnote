"""doc_style — 비슷한 저장 공문 고르기·작성 규칙 (UI 없음)."""
import doc_style as ds


def doc(title, content='1. 본문\n끝.', doc_type='general'):
    return {'title': title, 'content': content, 'doc_type': doc_type}


DOCS = [   # db.get_official_documents()처럼 최근 순
    doc('2026. 1학기 교무회의 운영 계획'),
    doc('2026학년도 교내 수학경시대회 결과 보고', doc_type='result_report'),
    doc('2025학년도 교내 수학경시대회 운영 계획'),
    doc('학업성적관리위원회 회의록 제출', doc_type='submit'),
    doc('현장체험학습 실시 계획'),
]


def titles(rs):
    return [r['title'] for r in rs]


def test_finds_similar_title_first():
    rs = ds.search_similar('2026학년도 교내 수학경시대회 운영 계획', DOCS, doc_type='general')
    assert titles(rs)[0] == '2025학년도 교내 수학경시대회 운영 계획'
    assert '2026학년도 교내 수학경시대회 결과 보고' in titles(rs)


def test_same_type_preferred():
    rs = ds.search_similar('2026학년도 교내 수학경시대회 결과 보고서', DOCS, doc_type='result_report', top_k=1)
    assert titles(rs) == ['2026학년도 교내 수학경시대회 결과 보고']


def test_common_words_alone_do_not_match_but_same_type_fills():
    # '운영 계획'만 겹침 → 비슷한 공문은 아니지만 같은 종류 최근 공문으로 채움 (score 0)
    rs = ds.search_similar('독서 동아리 운영 계획', DOCS, doc_type='general')
    assert len(rs) == 2 and all(r['doc_type'] == 'general' for r in rs)
    rs = ds.search_similar('아무 관련 없는 제목', DOCS, doc_type='purchase')
    assert rs == []


def test_dedupes_same_title_and_skips_empty():
    docs = [doc('교무회의 운영 계획', '새 본문'), doc('교무회의 운영 계획', '옛 본문'), doc('교무회의 자료', '')]
    rs = ds.search_similar('교무회의 운영 계획', docs, top_k=3)
    assert [r['content'] for r in rs] == ['새 본문']


def test_empty_inputs():
    assert ds.search_similar('교무회의', []) == []
    assert ds.search_similar('', DOCS) == []
    assert ds.examples_prompt([]) == ''


def test_examples_prompt_truncates_long_body():
    body = '가' * (ds.EXAMPLE_MAX_CHARS + 100)
    p = ds.examples_prompt([doc('긴 공문', body)])
    assert '[예시 1] 제목: 긴 공문' in p and '이하 생략' in p and '가' * (ds.EXAMPLE_MAX_CHARS + 1) not in p


def test_style_notes_roundtrip(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    assert ds.load_style_notes() == '' and ds.style_notes_prompt() == ''
    ds.save_style_notes('  대상은 1~3학년 전체 학생으로  ')
    assert ds.load_style_notes() == '대상은 1~3학년 전체 학생으로'
    assert '대상은 1~3학년 전체 학생으로' in ds.style_notes_prompt()
    ds.save_style_notes('가' * (ds.STYLE_NOTES_MAX + 50))
    assert len(ds.load_style_notes()) == ds.STYLE_NOTES_MAX
