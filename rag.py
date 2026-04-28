"""
rag.py — 경량 RAG: 임베딩 기반 유사 공문 검색

- 모델: paraphrase-multilingual-MiniLM-L12-v2 (한국어 지원, ~470MB)
- 최초 1회 로드 후 싱글톤 캐싱
- 코사인 유사도로 top-k 검색
"""

import os
import struct
import numpy as np

_encoder = None
_MODEL_NAME = 'paraphrase-multilingual-MiniLM-L12-v2'


def _get_encoder():
    global _encoder
    if _encoder is None:
        from sentence_transformers import SentenceTransformer
        _encoder = SentenceTransformer(_MODEL_NAME)
    return _encoder


def encode(text: str) -> np.ndarray:
    """단일 텍스트 → 임베딩 벡터"""
    return _get_encoder().encode(text, normalize_embeddings=True)


def encode_batch(texts: list[str]) -> np.ndarray:
    """여러 텍스트 → 임베딩 행렬"""
    return _get_encoder().encode(texts, normalize_embeddings=True)


def embedding_to_bytes(vec: np.ndarray) -> bytes:
    """float32 벡터 → SQLite BLOB"""
    return struct.pack(f'{len(vec)}f', *vec)


def bytes_to_embedding(data: bytes) -> np.ndarray:
    """SQLite BLOB → float32 벡터"""
    return np.array(struct.unpack(f'{len(data)//4}f', data), dtype=np.float32)


def cosine_similarity(a: np.ndarray, b: np.ndarray) -> float:
    """코사인 유사도 (정규화된 벡터는 내적과 동일)"""
    return float(np.dot(a, b))


def search_similar(query: str, embeddings: list[dict], top_k: int = 3) -> list[dict]:
    """
    쿼리와 유사한 임베딩 docs 검색.
    embeddings: db.get_all_embeddings() 결과
    반환: [{'id': ..., 'title': ..., 'content': ..., 'doc_type': ..., 'score': ...}]
    """
    if not embeddings:
        return []

    query_vec = encode(query)
    results = []
    for emb in embeddings:
        blob = emb.get('embedding')
        if not blob:
            continue
        doc_vec = bytes_to_embedding(blob)
        score = cosine_similarity(query_vec, doc_vec)
        results.append({
            'id': emb['id'],
            'title': emb['title'],
            'content': emb['content'],
            'doc_type': emb['doc_type'],
            'score': score,
        })

    results.sort(key=lambda x: x['score'], reverse=True)
    return results[:top_k]


def make_search_prompt(docs: list[dict]) -> str:
    """검색된 문서들 → 프롬프트에 포함할 few-shot 문자열"""
    if not docs:
        return ''
    lines = ['\n참고할 기존 공문 예시:']
    for i, d in enumerate(docs, 1):
        lines.append(f'\n[예시 {i}]')
        if d.get('title'):
            lines.append(f'제목: {d["title"]}')
        if d.get('content'):
            lines.append(d['content'])
    return '\n'.join(lines)
