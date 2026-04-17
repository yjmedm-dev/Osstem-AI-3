"""
ChromaDB 기반 벡터 스토어.
한국어 지원 다국어 임베딩 모델(paraphrase-multilingual-MiniLM-L12-v2)을 사용한다.
"""
from __future__ import annotations

import chromadb
from chromadb.config import Settings
from sentence_transformers import SentenceTransformer
from pathlib import Path
from typing import Any

from analysis.chunk_builder import Chunk

# 다국어 소형 모델 — 한국어 포함, 속도 빠름
_EMBED_MODEL = "paraphrase-multilingual-MiniLM-L12-v2"
_COLLECTION  = "bspl_rag"


class VectorStore:
    def __init__(self, persist_dir: str | Path):
        self._dir = Path(persist_dir)
        self._dir.mkdir(parents=True, exist_ok=True)

        self._client = chromadb.PersistentClient(
            path=str(self._dir),
            settings=Settings(anonymized_telemetry=False),
        )
        self._model = SentenceTransformer(_EMBED_MODEL)
        self._col = self._client.get_or_create_collection(
            name=_COLLECTION,
            metadata={"hnsw:space": "cosine"},
        )

    # ── 인덱싱 ───────────────────────────────────────────────────────

    def index(self, chunks: list[Chunk], reset: bool = False) -> None:
        if reset:
            self._client.delete_collection(_COLLECTION)
            self._col = self._client.get_or_create_collection(
                name=_COLLECTION,
                metadata={"hnsw:space": "cosine"},
            )

        texts     = [c.text for c in chunks]
        ids       = [c.chunk_id for c in chunks]
        metadatas = [c.metadata for c in chunks]

        print(f"임베딩 생성 중 ({len(chunks)}개 청크)...")
        embeddings = self._model.encode(texts, show_progress_bar=True).tolist()

        self._col.upsert(
            ids=ids,
            embeddings=embeddings,
            documents=texts,
            metadatas=metadatas,
        )
        print(f"인덱싱 완료: {self._col.count()}개 문서 저장됨")

    # ── 검색 ─────────────────────────────────────────────────────────

    def search(
        self,
        query: str,
        n_results: int = 5,
        where: dict[str, Any] | None = None,
    ) -> list[dict]:
        query_emb = self._model.encode([query]).tolist()
        kwargs: dict[str, Any] = {
            "query_embeddings": query_emb,
            "n_results": min(n_results, self._col.count()),
            "include": ["documents", "metadatas", "distances"],
        }
        if where:
            kwargs["where"] = where

        result = self._col.query(**kwargs)
        hits = []
        for doc, meta, dist in zip(
            result["documents"][0],
            result["metadatas"][0],
            result["distances"][0],
        ):
            hits.append({"text": doc, "metadata": meta, "score": 1 - dist})
        return hits

    def count(self) -> int:
        return self._col.count()
