"""
결산 데이터 RAG 챗봇 CLI.

사용법:
  python -m analysis.chatbot index          # 데이터 인덱싱 (최초 1회)
  python -m analysis.chatbot chat           # 대화 시작
  python -m analysis.chatbot index --reset  # 인덱스 재구축
"""
from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from config.settings import ANTHROPIC_API_KEY
from analysis.chunk_builder import build_all_chunks
from analysis.vector_store import VectorStore
from analysis.rag_engine import RAGEngine

_VECTOR_DIR = BASE_DIR / "data" / "vector_db"

HELP_TEXT = """
명령어:
  /reset    — 대화 히스토리 초기화
  /quit     — 종료
  /help     — 도움말
  /법인 이름 — 해당 법인 데이터만 검색 (예: /러시아)
"""

BANNER = """
╔══════════════════════════════════════════════╗
║   Osstem 해외법인 결산 AI 챗봇 (2603)         ║
║   커버 법인: 러시아 / 우즈베키스탄 / 우크라이나 / 카자흐스탄  ║
╚══════════════════════════════════════════════╝
질문 예시:
  - 카자흐스탄 매출이 왜 급감했나요?
  - 우즈베키스탄 영업이익 흑자전환 배경이 뭔가요?
  - 4개 법인 중 영업이익률이 가장 높은 곳은?
  - 자본잠식 법인은 어디인가요?
/help 입력 시 명령어 목록 표시
"""


def cmd_index(reset: bool = False) -> None:
    print("BSPL 데이터 로딩 중...")
    chunks = build_all_chunks(BASE_DIR)
    print(f"청크 생성 완료: {len(chunks)}개")

    vs = VectorStore(_VECTOR_DIR)
    vs.index(chunks, reset=reset)
    print("인덱싱 완료. 이제 'python -m analysis.chatbot chat' 으로 대화를 시작하세요.")


def cmd_chat() -> None:
    vs = VectorStore(_VECTOR_DIR)
    if vs.count() == 0:
        print("인덱스가 비어있습니다. 먼저 다음 명령을 실행하세요:")
        print("  python -m analysis.chatbot index")
        return

    if not ANTHROPIC_API_KEY:
        print("오류: .env 파일에 ANTHROPIC_API_KEY가 없습니다.")
        return

    engine = RAGEngine(vs, ANTHROPIC_API_KEY)
    print(BANNER)

    corp_filter: str | None = None

    while True:
        prompt_label = f"[{corp_filter}] " if corp_filter else ""
        try:
            user_input = input(f"\n{prompt_label}질문 > ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n종료합니다.")
            break

        if not user_input:
            continue

        # 명령어 처리
        if user_input.lower() in ("/quit", "/exit", "exit", "quit"):
            print("종료합니다.")
            break
        elif user_input.lower() == "/help":
            print(HELP_TEXT)
            continue
        elif user_input.lower() == "/reset":
            engine.reset_history()
            corp_filter = None
            continue
        elif user_input.startswith("/"):
            # /법인명 필터
            name = user_input[1:].strip()
            valid = ["러시아", "우즈베키스탄", "우크라이나", "카자흐스탄", ""]
            if name in valid:
                corp_filter = name if name else None
                print(f"법인 필터 설정: {corp_filter or '전체'}")
            else:
                print(f"알 수 없는 명령어: {user_input}")
            continue

        # RAG 답변
        print("\n분석 중...\n")
        try:
            answer = engine.ask(user_input, corp_filter=corp_filter)
            print(answer)
        except Exception as e:
            print(f"오류 발생: {e}")


def main() -> None:
    args = sys.argv[1:]
    if not args or args[0] == "chat":
        cmd_chat()
    elif args[0] == "index":
        cmd_index(reset="--reset" in args)
    else:
        print(f"알 수 없는 명령: {args[0]}")
        print("사용법: python -m analysis.chatbot [index|chat] [--reset]")


if __name__ == "__main__":
    main()
