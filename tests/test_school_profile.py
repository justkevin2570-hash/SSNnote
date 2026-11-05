import importlib

import pytest


@pytest.fixture
def sp(tmp_path, monkeypatch):
    import db
    monkeypatch.setattr(db, 'DB_DIR', str(tmp_path))
    monkeypatch.setattr(db, 'DB_PATH', str(tmp_path / 'ssnnote.db'))
    db.init_db()
    import school_profile
    return importlib.reload(school_profile)


def test_profile_roundtrip_and_prompt(sp):
    assert sp.prompt_block() == ''
    sp.save_profile({'school': '서서니고', 'department': '교무기획부', 'name': '홍길동',
                     'phone': '031-123-4567', 'mask_names': '김철수, 이영희'})
    p = sp.load_profile()
    assert p['school'] == '서서니고' and p['position'] == ''
    block = sp.prompt_block()
    assert '학교명: 서서니고' in block and '연락처: 031-123-4567' in block and '직위' not in block


def test_privacy_default_on_and_toggle(sp):
    assert sp.privacy_enabled() is True
    sp.set_privacy_enabled(False)
    assert sp.privacy_enabled() is False


def test_mask_for_ai_uses_profile(sp):
    sp.save_profile({'name': '홍길동', 'phone': '031-123-4567', 'mask_names': '김철수, 홍길동'})
    masked, counts = sp.mask_for_ai('담당 홍길동(031-123-4567), 학생 김철수 보호자 010-1111-2222')
    assert '홍길동' in masked and '031-123-4567' in masked      # 담당자 본인·학교 연락처는 유지
    assert '김○○' in masked and '010-****-2222' in masked
    assert counts == {'phone': 1, 'name': 1}
