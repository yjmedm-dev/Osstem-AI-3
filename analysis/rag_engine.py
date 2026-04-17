"""
RAG 엔진 — 벡터 검색 결과를 컨텍스트로 붙여 Claude API에 질의한다.
"""
from __future__ import annotations

import anthropic
from analysis.vector_store import VectorStore

_SYSTEM_PROMPT = """당신은 해외법인 재무 분석 전문가입니다.
Osstem Implant 해외관리2팀이 매월 수행하는 법인 월마감 결산 데이터를 기반으로 답변합니다.

답변 규칙:
1. 제공된 [컨텍스트]에 있는 수치만 사용하세요. 수치를 지어내지 마세요.
2. 금액은 현지 통화 단위를 명시하세요 (RUB, UZS, UAH, KZT).
3. 전년 대비 증감률과 마진 지표를 적극 활용해 분석하세요.
4. 컨텍스트에 없는 정보는 "해당 정보는 데이터에 없습니다"라고 답하세요.
5. 답변은 핵심을 먼저 말하고, 근거 수치를 뒤에 제시하세요.
"""


class RAGEngine:
    def __init__(self, vector_store: VectorStore, api_key: str, model: str = "claude-sonnet-4-6"):
        self._vs    = vector_store
        self._client = anthropic.Anthropic(api_key=api_key)
        self._model  = model
        self._history: list[dict] = []  # 대화 히스토리

    def ask(self, question: str, n_context: int = 6, corp_filter: str | None = None) -> str:
        # 1) 관련 청크 검색
        where = {"corp": corp_filter} if corp_filter else None
        hits  = self._vs.search(question, n_results=n_context, where=where)

        if not hits:
            return "인덱싱된 데이터가 없습니다. 먼저 `index` 명령으로 데이터를 인덱싱하세요."

        # 2) 컨텍스트 구성
        context_parts = []
        for i, h in enumerate(hits, 1):
            meta = h["metadata"]
            label = f"[출처 {i}: {meta.get('corp','?')} / {meta.get('type','?')}]"
            context_parts.append(f"{label}\n{h['text']}")
        context = "\n\n".join(context_parts)

        # 3) 메시지 구성 (멀티턴)
        user_msg = f"[컨텍스트]\n{context}\n\n[질문]\n{question}"
        self._history.append({"role": "user", "content": user_msg})

        # 4) Claude API 호출
        response = self._client.messages.create(
            model=self._model,
            max_tokens=2048,
            system=_SYSTEM_PROMPT,
            messages=self._history,
        )
        answer = response.content[0].text

        # 5) 히스토리에 답변 추가 (멀티턴 유지)
        self._history.append({"role": "assistant", "content": answer})

        # 히스토리가 너무 길면 앞부분 정리 (최근 10턴 유지)
        if len(self._history) > 20:
            self._history = self._history[-20:]

        return answer

    def reset_history(self) -> None:
        self._history = []
        print("대화 히스토리를 초기화했습니다.")
