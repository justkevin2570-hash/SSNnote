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


def test_rrn_without_hyphen_and_date_check():
    masked, counts = mask_pii('주민번호 0805123123456 / 080512 3123456 / 9913011234567')
    assert masked.count('080512-3******') == 2
    assert '9913011234567' in masked            # 13월은 생년월일이 아님
    assert counts == {'rrn': 2}


def test_more_phone_formats():
    masked, counts = mask_pii('+82 10-1234-5678, 0311234567, 021234567, 07012345678')
    assert masked == '+82 10-****-5678, 031-****-4567, 02-****-4567, 070-****-5678'
    assert counts == {'phone': 4}


def test_account_apartment_birth():
    masked, counts = mask_pii('농협 352-1234-5678-13, 행복아파트 101동 1203호, 생년월일: 2008. 5. 12.')
    assert '***-****-****-**' in masked and '○○○동 ○○○호' in masked
    assert '2008' not in masked and '생년월일: ○○○○. ○○. ○○.' in masked
    assert counts == {'account': 1, 'address': 1, 'birth': 1}


def test_unregistered_names_in_clear_contexts():
    masked, counts = mask_pii('홍길동 학생, 2학년 3반 5번 김철수는, 5번 이영희, 보호자: 박영수, 예금주 최민호')
    assert masked == '홍○○ 학생, 2학년 3반 5번 김○○는, 5번 이○○, 보호자: 박○○, 예금주 최○○'
    assert counts == {'name': 5}


def test_name_heuristics_leave_common_words():
    text = '신청한 학생, 장애인 학생, 어려운 학생, 학부모 설명회, 보호자 동의서, 3번 문항을 확인'
    assert mask_pii(text) == (text, {})


def test_registered_name_needs_word_boundary():
    masked, counts = mask_pii('이수증 발급, 이수 학생은, 이수가 왔다, 이수에게', names=['이수'])
    assert masked == '이수증 발급, 이○ 학생은, 이○가 왔다, 이○에게'
    assert counts == {'name': 3}


def test_exempt_name_is_kept_by_heuristics():
    masked, _ = mask_pii('담당자 성명: 홍길동, 학생 성명: 김철수', exempt=['홍길동'])
    assert '홍길동' in masked and '김○○' in masked


def test_looks_masked():
    from privacy import looks_masked
    assert looks_masked('홍○○ 학생') and looks_masked('010-****-5678')
    assert not looks_masked('교무회의') and not looks_masked('')
