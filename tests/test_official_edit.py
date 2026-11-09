"""저장한 공문 수정 (db.update_official_document)."""
import pytest


@pytest.fixture
def db(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    return db


def test_update_changes_only_given_fields(db):
    doc_id = db.save_official_document('교무회의 운영 계획', '교무-1', '옛 본문', 'general')
    db.update_official_document(doc_id, '교무회의 운영 계획(수정)', '교무-2', '새 본문')
    d = next(x for x in db.get_official_documents() if x['id'] == doc_id)
    assert d['title'] == '교무회의 운영 계획(수정)'
    assert d['doc_number'] == '교무-2'
    assert d['content'] == '새 본문'
    assert d['doc_type'] == 'general'          # 종류는 그대로
    assert len(db.get_official_documents()) == 1
