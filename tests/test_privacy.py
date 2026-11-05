from privacy import mask_pii, describe, parse_name_list


def test_masks_rrn_phone_email():
    text = '학생 주민번호 080512-3123456, 학부모 010-1234-5678, 메일 parent@naver.com'
    masked, counts = mask_pii(text)
    assert '3123456' not in masked and '080512-3******' in masked
    assert '010-****-5678' in masked
    assert '***@naver.com' in masked
    assert counts == {'rrn': 1, 'phone': 1, 'email': 1}


def test_phone_variants():
    masked, counts = mask_pii('01012345678 / 010 1234 5678 / 02-123-4567 / 031)123-4567')
    assert masked.count('****') == 4
    assert counts['phone'] == 4


def test_does_not_touch_doc_numbers_dates_amounts():
    text = '중등교육과-12345(2026. 11. 5.) 예산 1,200,000원 2026-11-12까지 제출'
    masked, counts = mask_pii(text)
    assert masked == text and counts == {}


def test_exempt_keeps_school_contact():
    masked, _ = mask_pii('문의: 031-123-4567, 학부모 010-9999-8888', exempt=['031-123-4567'])
    assert '031-123-4567' in masked and '010-****-8888' in masked


def test_names_longest_first():
    masked, counts = mask_pii('김민수, 김민수진 학생', names=['김민수', '김민수진'])
    assert masked == '김○○, 김○○○ 학생'
    assert counts == {'name': 2}


def test_single_char_names_ignored():
    masked, counts = mask_pii('김 선생님', names=['김'])
    assert masked == '김 선생님' and counts == {}


def test_describe_and_parse():
    assert describe({'phone': 2, 'name': 1}) == '전화번호 2건, 이름 1건'
    assert parse_name_list('홍길동, 김철수\n이영희  박') == ['홍길동', '김철수', '이영희']
