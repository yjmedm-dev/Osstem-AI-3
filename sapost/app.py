"""
해외관리실의 공간 — Streamlit 웹 앱
sapost/app.py

실행:
  streamlit run sapost/app.py
"""
import sys
import io
import logging
import queue
import threading
import time
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path

import streamlit as st

# 프로젝트 루트를 sys.path에 추가
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from sapost.fbl5n_download import (
    ACCOUNT_CORP_MAP,
    _CORP_NAME_MAP,
    _resolve_accounts_from_input,
    _parse_date_arg,
    get_customer_accounts,
    FBL5NDownloader,
)
from sapost.zqsab01_download import ZQSAB01Downloader
from sapost.src.utils import get_config


# ─────────────────────────────────────────────────────────────
# 로깅 핸들러: 로그 레코드를 Queue에 기록
# ─────────────────────────────────────────────────────────────
class _QueueHandler(logging.Handler):
    def __init__(self, q: queue.Queue):
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord):
        self.q.put(self.format(record))


# ─────────────────────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="해외관리실의 공간",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────
# 세션 상태 초기화
# ─────────────────────────────────────────────────────────────
_DEFAULTS: dict = {
    "stage":           "idle",   # idle | input | confirm | running | done
    "params":          {},
    "log_lines":       [],
    "run_error":       None,
    "_thread_started": False,
    "_log_queue":      None,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


def _reset_state():
    for _k, _v in _DEFAULTS.items():
        st.session_state[_k] = _v if not isinstance(_v, (list, dict)) else type(_v)()


# ─────────────────────────────────────────────────────────────
# 사이드바 — 메뉴 (추후 모듈 추가 시 options 목록에 항목 추가)
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    _assets = Path(__file__).parent / "assets"
    _imgs = sorted(_assets.glob("*.png")) if _assets.exists() else []
    if len(_imgs) >= 2:
        _c1, _c2 = st.columns(2)
        _c1.image(str(_imgs[0]), width=110)
        _c2.image(str(_imgs[1]), width=110)
    elif len(_imgs) == 1:
        st.image(str(_imgs[0]), width=120)
    st.markdown("## 해외관리실의 공간")
    st.divider()
    menu = st.radio(
        "메뉴",
        options=["채권명세서", "품목별 연결손익"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption("메뉴를 선택하세요.")


# ─────────────────────────────────────────────────────────────
# 메인 헤더
# ─────────────────────────────────────────────────────────────
st.title("해외관리실의 공간")


# ═════════════════════════════════════════════════════════════
# 채권명세서
# ═════════════════════════════════════════════════════════════
if menu == "채권명세서":
    st.subheader("채권명세서 업데이트")
    st.caption("SAP FBL5N 미결/반제 항목을 조회해 채권명세서 파일에 자동 기입합니다.")
    st.divider()

    # ── IDLE ─────────────────────────────────────────────────
    if st.session_state.stage == "idle":
        if st.button("실행", type="primary"):
            st.session_state.stage = "input"
            st.rerun()

    # ── INPUT ────────────────────────────────────────────────
    elif st.session_state.stage == "input":
        st.markdown("### 조건 입력")

        with st.form("input_form"):
            # 기간
            st.markdown("**기간**")
            col1, col2 = st.columns(2)
            with col1:
                budat_low_raw = st.text_input("시작일 (YYYYMMDD)", placeholder="20260301")
            with col2:
                budat_high_raw = st.text_input("종료일 (YYYYMMDD)", placeholder="20260331")

            st.markdown("---")

            # 법인
            st.markdown("**법인**")
            st.caption(
                "7자리 고객코드, 법인명(부분 일치), 여러 개는 쉼표/공백 구분 · "
                "**일체**: 경로 내 전체 · **해외법인 일체**: 전체 해외법인 · "
                "**유럽 일체**: '유럽' 포함 법인 전체 · 빈칸: 경로 내 전체"
            )
            corp_input = st.text_input(
                "법인",
                placeholder="예) 유럽  /  1700031  /  독일, 프랑스  /  해외법인 일체",
                label_visibility="collapsed",
            )

            st.markdown("---")

            # 경로
            st.markdown("**채권명세서 폴더 경로**")
            st.caption("빈칸으로 두면 config.ini의 source_dir 경로를 사용합니다.")
            source_dir_raw = st.text_input(
                "경로",
                placeholder=r"예) D:\법인서류\채권명세서",
                label_visibility="collapsed",
            )

            st.markdown("---")
            col_next, col_cancel, _ = st.columns([1, 1, 6])
            with col_next:
                submitted = st.form_submit_button("다음 →", type="primary")
            with col_cancel:
                cancelled = st.form_submit_button("취소")

        if cancelled:
            st.session_state.stage = "idle"
            st.rerun()

        if submitted:
            errors: list[str] = []
            budat_low = budat_high = None

            if not budat_low_raw.strip():
                errors.append("시작일을 입력해주세요.")
            else:
                try:
                    budat_low = _parse_date_arg(budat_low_raw.strip())
                except ValueError as e:
                    errors.append(f"시작일 오류: {e}")

            if not budat_high_raw.strip():
                errors.append("종료일을 입력해주세요.")
            else:
                try:
                    budat_high = _parse_date_arg(budat_high_raw.strip())
                except ValueError as e:
                    errors.append(f"종료일 오류: {e}")

            if errors:
                for err in errors:
                    st.error(err)
            else:
                yyyymm = budat_high_raw.strip()[:6]

                # '일체' 처리 시 print() 출력 캡처
                buf = io.StringIO()
                with redirect_stdout(buf):
                    accounts = _resolve_accounts_from_input(corp_input.strip())
                resolve_msg = buf.getvalue().strip()

                st.session_state.params = {
                    "budat_low":    budat_low,
                    "budat_high":   budat_high,
                    "yyyymm":       yyyymm,
                    "accounts":     accounts,
                    "source_dir":   source_dir_raw.strip() or None,
                    "_resolve_msg": resolve_msg,
                }
                st.session_state.stage = "confirm"
                st.rerun()

    # ── CONFIRM ──────────────────────────────────────────────
    elif st.session_state.stage == "confirm":
        p = st.session_state.params
        st.markdown("### 실행 조건 확인")

        if p["accounts"]:
            corp_display = ", ".join(ACCOUNT_CORP_MAP.get(a, a) for a in p["accounts"])
            corp_display += f"  ({len(p['accounts'])}개)"
        else:
            corp_display = "경로 내 전체"

        if p.get("_resolve_msg"):
            st.info(p["_resolve_msg"])

        st.table(
            {
                "항목": ["기간", "법인", "경로"],
                "값": [
                    f"{p['budat_low']} ~ {p['budat_high']}",
                    corp_display,
                    p["source_dir"] or "(config.ini 기본값)",
                ],
            }
        )

        col1, col2, col3, _ = st.columns([1, 1, 1, 5])
        with col1:
            if st.button("확인", type="primary"):
                st.session_state.stage           = "running"
                st.session_state.log_lines       = []
                st.session_state.run_error       = None
                st.session_state._thread_started = False
                st.session_state._log_queue      = queue.Queue()
                st.rerun()
        with col2:
            if st.button("수정"):
                st.session_state.stage = "input"
                st.rerun()
        with col3:
            if st.button("취소"):
                _reset_state()
                st.rerun()

    # ── RUNNING ───────────────────────────────────────────────
    elif st.session_state.stage == "running":
        p = st.session_state.params
        q: queue.Queue = st.session_state._log_queue  # type: ignore[assignment]

        st.markdown("### 실행 중")
        st.info("SAP GUI에서 조회가 진행 중입니다. 완료될 때까지 기다려주세요.")

        log_placeholder = st.empty()

        # 스레드 최초 기동 (재실행 시 중복 기동 방지)
        if not st.session_state._thread_started:
            st.session_state._thread_started = True

            # q를 기본 인수로 바인딩해 스레드 내 session_state 직접 접근 방지
            def _run(q: queue.Queue = q, p: dict = p) -> None:
                try:
                    config = get_config()

                    logger = logging.getLogger("sapost.fbl5n.app")
                    logger.setLevel(logging.DEBUG)
                    logger.handlers = []
                    h = _QueueHandler(q)
                    h.setFormatter(
                        logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S")
                    )
                    logger.addHandler(h)

                    if p["source_dir"]:
                        src = Path(p["source_dir"])
                        config.set("PATHS", "source_dir", str(src))
                        config.set("PATHS", "raw_dir",    str(src / "raw"))

                    source_dir = Path(config.get("PATHS", "source_dir"))

                    if p["accounts"]:
                        accounts = p["accounts"]
                        logger.info(f"지정 계정 {len(accounts)}개: {accounts}")
                    else:
                        accounts = get_customer_accounts(source_dir, logger)
                        if not accounts:
                            raise RuntimeError("고객계정을 추출할 파일이 없습니다.")
                        logger.info(f"총 {len(accounts)}개 계정 추출: {accounts}")

                    downloader = FBL5NDownloader(config, logger)
                    downloader.connect()
                    try:
                        downloader.run_all(
                            accounts, p["budat_low"], p["budat_high"], p["yyyymm"]
                        )
                    finally:
                        downloader.close()

                    q.put("__DONE__")

                except Exception as exc:
                    q.put(f"__ERROR__{exc}")

            threading.Thread(target=_run, daemon=True).start()

        # 큐 → log_lines 수집
        stage_changed = False
        while not q.empty():
            item: str = q.get_nowait()
            if item == "__DONE__":
                st.session_state.stage = "done"
                stage_changed = True
                break
            elif item.startswith("__ERROR__"):
                st.session_state.run_error = item[len("__ERROR__"):]
                st.session_state.stage = "done"
                stage_changed = True
                break
            else:
                st.session_state.log_lines.append(item)

        # 로그 표시
        if st.session_state.log_lines:
            log_placeholder.text_area(
                "로그",
                value="\n".join(st.session_state.log_lines),
                height=450,
                label_visibility="collapsed",
            )

        if stage_changed:
            st.rerun()
        else:
            time.sleep(1)
            st.rerun()

    # ── DONE ─────────────────────────────────────────────────
    elif st.session_state.stage == "done":
        if st.session_state.run_error:
            st.error(f"오류가 발생했습니다: {st.session_state.run_error}")
        else:
            st.success("채권명세서 업데이트가 완료되었습니다.")

        if st.session_state.log_lines:
            st.text_area(
                "실행 로그",
                value="\n".join(st.session_state.log_lines),
                height=450,
            )

        if st.button("처음으로", type="primary"):
            _reset_state()
            st.rerun()


# ═════════════════════════════════════════════════════════════
# 품목별 연결손익 (ZQSAB01)
# ═════════════════════════════════════════════════════════════
elif menu == "품목별 연결손익":
    st.subheader("품목별 연결손익 조회")
    st.caption("SAP ZQSAB01에서 회계연도·기간을 입력해 품목별 연결손익 데이터를 추출합니다.")
    st.divider()

    # ── IDLE ─────────────────────────────────────────────────
    if st.session_state.stage == "idle":
        if st.button("실행", type="primary"):
            st.session_state.stage = "input"
            st.rerun()

    # ── INPUT ────────────────────────────────────────────────
    elif st.session_state.stage == "input":
        st.markdown("### 조회 기간 입력")

        with st.form("zqsab01_input_form"):
            col1, col2 = st.columns(2)
            with col1:
                perio = st.text_input(
                    "작업기간 (YYYYMM)",
                    value=datetime.now().strftime("%Y%m"),
                    placeholder="202603",
                    help="SAP P_PERIO 필드에 입력되는 기간값입니다.",
                )
            with col2:
                pcode = st.text_input(
                    "품목코드 (선택)",
                    value="",
                    placeholder="생략 시 전체 조회",
                    help="특정 품목만 조회할 경우 입력하세요.",
                )

            st.markdown("---")
            col_next, col_cancel, _ = st.columns([1, 1, 6])
            with col_next:
                submitted = st.form_submit_button("다음 →", type="primary")
            with col_cancel:
                cancelled = st.form_submit_button("취소")

        if cancelled:
            st.session_state.stage = "idle"
            st.rerun()

        if submitted:
            errors: list[str] = []
            if not perio.strip():
                errors.append("작업기간을 입력해주세요.")

            if errors:
                for err in errors:
                    st.error(err)
            else:
                st.session_state.params = {
                    "perio": perio.strip(),
                    "pcode": pcode.strip(),
                }
                st.session_state.stage = "confirm"
                st.rerun()

    # ── CONFIRM ──────────────────────────────────────────────
    elif st.session_state.stage == "confirm":
        p = st.session_state.params
        st.markdown("### 실행 조건 확인")
        st.table(
            {
                "항목": ["작업기간", "품목코드"],
                "값": [
                    p["perio"],
                    p["pcode"] or "(전체 조회)",
                ],
            }
        )

        col1, col2, col3, _ = st.columns([1, 1, 1, 5])
        with col1:
            if st.button("확인", type="primary"):
                st.session_state.stage           = "running"
                st.session_state.log_lines       = []
                st.session_state.run_error       = None
                st.session_state._thread_started = False
                st.session_state._log_queue      = queue.Queue()
                st.rerun()
        with col2:
            if st.button("수정"):
                st.session_state.stage = "input"
                st.rerun()
        with col3:
            if st.button("취소"):
                _reset_state()
                st.rerun()

    # ── RUNNING ───────────────────────────────────────────────
    elif st.session_state.stage == "running":
        p = st.session_state.params
        q: queue.Queue = st.session_state._log_queue  # type: ignore[assignment]

        st.markdown("### 실행 중")
        st.info("SAP GUI에서 ZQSAB01 조회가 진행 중입니다. 완료될 때까지 기다려주세요.")

        log_placeholder = st.empty()

        if not st.session_state._thread_started:
            st.session_state._thread_started = True

            def _run_zqsab01(q: queue.Queue = q, p: dict = p) -> None:
                try:
                    config = get_config()

                    logger = logging.getLogger("sapost.zqsab01.app")
                    logger.setLevel(logging.DEBUG)
                    logger.handlers = []
                    h = _QueueHandler(q)
                    h.setFormatter(
                        logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S")
                    )
                    logger.addHandler(h)

                    dl = ZQSAB01Downloader(config, logger)
                    dl.connect()
                    try:
                        dl.navigate()
                        dl.set_params_and_execute(
                            perio = p["perio"],
                            pcode = p["pcode"],
                        )
                        raw = dl.wait_for_download(p["perio"])
                        out = dl.process_excel(raw, p["perio"])
                        logger.info(f"결과 파일: {out}")
                        q.put(f"__OUTPATH__{out}")
                    finally:
                        dl.close()

                    q.put("__DONE__")
                except Exception as exc:
                    q.put(f"__ERROR__{exc}")

            threading.Thread(target=_run_zqsab01, daemon=True).start()

        stage_changed = False
        while not q.empty():
            item: str = q.get_nowait()
            if item == "__DONE__":
                st.session_state.stage = "done"
                stage_changed = True
                break
            elif item.startswith("__ERROR__"):
                st.session_state.run_error = item[len("__ERROR__"):]
                st.session_state.stage = "done"
                stage_changed = True
                break
            elif item.startswith("__OUTPATH__"):
                st.session_state.params["_out_path"] = item[len("__OUTPATH__"):]
            else:
                st.session_state.log_lines.append(item)

        if st.session_state.log_lines:
            log_placeholder.text_area(
                "로그",
                value="\n".join(st.session_state.log_lines),
                height=450,
                label_visibility="collapsed",
            )

        if stage_changed:
            st.rerun()
        else:
            time.sleep(1)
            st.rerun()

    # ── DONE ─────────────────────────────────────────────────
    elif st.session_state.stage == "done":
        if st.session_state.run_error:
            st.error(f"오류가 발생했습니다: {st.session_state.run_error}")
        else:
            out_path = st.session_state.params.get("_out_path", "")
            st.success(f"조회가 완료되었습니다.")
            if out_path:
                st.info(f"저장 경로: {out_path}")

        if st.session_state.log_lines:
            st.text_area(
                "실행 로그",
                value="\n".join(st.session_state.log_lines),
                height=450,
            )

        if st.button("처음으로", type="primary"):
            _reset_state()
            st.rerun()
