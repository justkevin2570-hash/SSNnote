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


def test_privacy_default_on_and_toggle(sp):
    assert sp.privacy_enabled() is True
    sp.set_privacy_enabled(False)
    assert sp.privacy_enabled() is False


def test_mask_for_ai(sp):
    masked, counts = sp.mask_for_ai('학생 김철수 보호자 010-1111-2222')
    assert '김○○' in masked and '010-****-2222' in masked
    assert counts.get('phone') == 1
    assert not hasattr(sp, 'load_profile') and not hasattr(sp, 'prompt_block')
