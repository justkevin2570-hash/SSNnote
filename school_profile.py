"""school_profile.py — 학교 프로필과 개인정보 보호 설정.

학교명·부서·담당자 등은 공문 초안 프롬프트에 자동으로 들어가고,
연락처는 개인정보 가리기에서 예외로 처리된다. 저장은 db.app_settings.
"""

import json

from db import get_setting, set_setting
from privacy import mask_pii, parse_name_list

PROFILE_KEY = 'school_profile'
PRIVACY_KEY = 'privacy_mode'

# (키, 화면 이름, 입력 예시)
PROFILE_FIELDS = [
    ('school',      '학교명',          '예: 서서니고등학교'),
    ('department',  '부서',            '예: 교무기획부'),
    ('name',        '담당자',          '예: 홍길동'),
    ('position',    '직위·직급',       '예: 교사'),
    ('phone',       '연락처',          '예: 031-123-4567'),
    ('approval',    '결재 라인',       '예: 교사 → 부장 → 교감 → 교장'),
    ('attachments', '자주 쓰는 붙임',  '예: 운영 계획서, 참가 신청서'),
]


def load_profile() -> dict:
    try:
        data = json.loads(get_setting(PROFILE_KEY, '{}') or '{}')
    except (ValueError, TypeError):
        data = {}
    profile = {k: str(data.get(k, '')).strip() for k, _, _ in PROFILE_FIELDS}
    profile['mask_names'] = str(data.get('mask_names', '')).strip()
    return profile


def save_profile(profile: dict):
    keys = [k for k, _, _ in PROFILE_FIELDS] + ['mask_names']
    set_setting(PROFILE_KEY, json.dumps({k: str(profile.get(k, '')).strip() for k in keys},
                                        ensure_ascii=False))


def prompt_block(profile: dict = None) -> str:
    """공문 초안 프롬프트에 붙일 학교 정보 문단. 비어 있으면 ''."""
    p = profile if profile is not None else load_profile()
    lines = [f'- {label}: {p[k]}' for k, label, _ in PROFILE_FIELDS if p.get(k)]
    if not lines:
        return ''
    return '[작성자 정보 — 필요한 곳(담당자, 문의처, 붙임 등)에 반영]\n' + '\n'.join(lines) + '\n'


# ── 개인정보 보호 모드 ─────────────────────────────────────────

def privacy_enabled() -> bool:
    """기본값은 켬. 켜져 있으면 이미지 대신 PC에서 읽은 글자를, 개인정보를 가린 뒤 보낸다."""
    return get_setting(PRIVACY_KEY, '1') != '0'


def set_privacy_enabled(enabled: bool):
    set_setting(PRIVACY_KEY, '1' if enabled else '0')


def mask_for_ai(text: str, profile: dict = None):
    """프로필 기준(가릴 이름 목록, 예외 연락처)으로 개인정보를 가린다. (masked, counts)"""
    p = profile if profile is not None else load_profile()
    exempt = [p.get('phone', '')]
    keep_names = {p.get('name', '')}  # 담당자 본인 이름은 공문에 들어가야 하므로 유지
    names = [n for n in parse_name_list(p.get('mask_names', '')) if n not in keep_names]
    return mask_pii(text, names=names, exempt=exempt)
