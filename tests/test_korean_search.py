from korean_search import chosung, score, best_score


def test_chosung():
    assert chosung('교') == 'ㄱ' and chosung('힣') == 'ㅎ' and chosung('A') == 'A'


def test_match_kinds_and_order():
    t = '교무회의 자료 제출'
    prefix, sub, cho, loose = score('교무', t), score('회의', t), score('ㄱㅁㅎㅇ', t), score('교자제', t)
    assert prefix > sub > cho > loose > 0
    assert score('ㄱㅁ회의', t) > 0          # 초성 + 글자 섞기
    assert score('자료제출', t) > 0           # 띄어쓰기 무시
    assert score('xyz', t) == 0 and score('ㅋㅋ', t) == 0
    assert score('', t) == 1


def test_case_insensitive_and_best():
    assert score('ai', '📸 AI 공문 캡처') > 0
    assert best_score('예산', '₩ 예산 관리', 'budget') > best_score('예산', '공문 작성')
