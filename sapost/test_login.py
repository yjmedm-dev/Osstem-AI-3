"""
SAP 로그인 테스트 스크립트
실행 전: SAP GUI를 열어두세요 (로그인 화면 또는 로그인된 상태 모두 가능)

실행:
  python sapost/test_login.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from sapost.src.utils import get_config, setup_logger
from sapost.src.sap_controller import SAPController


def main():
    config = get_config()
    logger = setup_logger("sapost.test", config)

    logger.info("=" * 40)
    logger.info("SAP 로그인 테스트 시작")
    logger.info("=" * 40)

    sap = None
    try:
        sap = SAPController(config, logger)

        logger.info("[1] SAP GUI 세션 연결 중...")
        sap.connect()
        logger.info("    → 연결 성공")

        logger.info("[2] 로그인 시도...")
        sap.login()
        logger.info("    → 로그인 성공")

        # 현재 접속 정보 출력
        info = sap.session.Info
        logger.info(f"    → 사용자: {info.User}")
        logger.info(f"    → 클라이언트: {info.Client}")
        logger.info(f"    → 현재 트랜잭션: {info.Transaction}")

        logger.info("=" * 40)
        logger.info("테스트 완료 — 로그인 정상")
        logger.info("=" * 40)

    except Exception as e:
        logger.error(f"테스트 실패: {e}")
        sys.exit(1)
    finally:
        if sap:
            sap.close()


if __name__ == "__main__":
    main()
