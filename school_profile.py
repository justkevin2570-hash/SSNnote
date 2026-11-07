"""school_profile.py — 개인정보 보호 모드 설정 (저장은 db.app_settings).

(예전의 학교 정보 설정은 없앴다. 이름은 호환을 위해 그대로 둔다.)
"""

from db import get_setting, set_setting
from privacy import mask_pii

PRIVACY_KEY = 'privacy_mode'


def privacy_enabled() -> bool:
    """기본값은 켬. 켜져 있으면 이미지 대신 PC에서 읽은 글자를, 개인정보를 가린 뒤 보낸다."""
    return get_setting(PRIVACY_KEY, '1') != '0'


def set_privacy_enabled(enabled: bool):
    set_setting(PRIVACY_KEY, '1' if enabled else '0')


def mask_for_ai(text: str):
    """AI로 보내기 전에 개인정보를 가린다. (masked, counts)"""
    return mask_pii(text)
