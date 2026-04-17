"""
SAP 데이터 추출 자동화 파이프라인 진입점

사용법:
  python sapost/main.py --month 202503

  --month     기준월 (YYYYMM 형식, 생략 시 현재 월 사용)
  --skip-sap  SAP 조작을 건너뛰고 raw/ 에 이미 있는 파일로 처리
              (extract_mode = export 로 다운로드된 파일이 있을 때 사용)
"""
import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from sapost.src.utils import get_config, setup_logger
from sapost.src.sap_controller import SAPController
from sapost.src.data_processor import DataProcessor
from sapost.src.template_writer import TemplateWriter


def parse_args():
    parser = argparse.ArgumentParser(description="SAP 데이터 추출 자동화")
    parser.add_argument(
        "--month",
        default=datetime.now().strftime("%Y%m"),
        help="기준월 (예: 202503). 기본값: 현재 월",
    )
    parser.add_argument(
        "--skip-sap",
        action="store_true",
        help="SAP 조작 생략 (raw/ 폴더에 이미 다운로드된 파일이 있을 때 사용)",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    month = args.month

    # ── 초기화 ──────────────────────────────────────
    config = get_config()
    logger = setup_logger("sapost.main", config)

    logger.info("=" * 60)
    logger.info(f"SAP 데이터 추출 자동화 시작 | 기준월: {month}")
    logger.info("=" * 60)

    start_time = time.time()
    result_summary = {
        "month": month,
        "sap": None,
        "process": None,
        "write": None,
    }

    extract_mode = config.get("SAP", "extract_mode", fallback="grid")

    # ── STEP 1-2: SAP 연결 & 데이터 추출 ────────────
    raw_dir = Path(config.get("PATHS", "raw_dir"))
    raw_dir.mkdir(parents=True, exist_ok=True)

    df_raw = None  # grid 모드에서 직접 읽은 DataFrame

    if args.skip_sap:
        logger.info("[SKIP] SAP 조작 생략 — raw/ 폴더의 기존 파일 사용")
        file_paths = sorted(raw_dir.glob(f"{month}_*.xlsx"))
        if not file_paths:
            logger.error(f"raw/ 폴더에 {month}_*.xlsx 파일이 없습니다.")
            sys.exit(1)
        logger.info(f"기존 파일 {len(file_paths)}건 발견")
    else:
        sap = None
        try:
            logger.info("[STEP 1] SAP GUI 연결 및 로그인")
            sap = SAPController(config, logger)
            sap.connect()
            sap.login()

            logger.info("[STEP 2] 트랜잭션 이동 및 조회 실행")
            sap.navigate_to()
            sap.set_params_and_execute(month)
            result_summary["sap"] = "success"

            logger.info("[STEP 3] 데이터 추출")
            if extract_mode == "grid":
                df_raw = sap.get_data()
                if df_raw is None or df_raw.empty:
                    logger.error("ALV 그리드에서 데이터를 읽지 못했습니다.")
                    sys.exit(1)
                logger.info(f"ALV 그리드에서 {len(df_raw)}행 추출 완료")
                file_paths = []  # grid 모드에서는 파일 목록 불필요
            else:
                # export 모드: SAP 내보내기 → raw/ 파일 저장
                export_path = sap.export_to_file(month)
                file_paths = [export_path]
                logger.info(f"내보내기 완료: {export_path}")

        except Exception as e:
            logger.critical(f"SAP/추출 단계 오류: {e}")
            result_summary["sap"] = f"failed: {e}"
            sys.exit(1)
        finally:
            if sap:
                sap.close()

    # ── STEP 4: 데이터 정제 ──────────────────────────
    try:
        logger.info("[STEP 4] 데이터 정제")
        processor = DataProcessor(config, logger)

        if df_raw is not None:
            # grid 모드: 메모리 DataFrame 직접 정제
            df = processor.process_dataframe(df_raw, month)
        else:
            # export / skip-sap 모드: 엑셀 파일 읽어서 정제
            df = processor.process(file_paths, month)

        result_summary["process"] = f"{len(df)}행"
        logger.info(f"정제 완료: {len(df)}행")
    except Exception as e:
        logger.critical(f"데이터 정제 오류: {e}")
        result_summary["process"] = f"failed: {e}"
        sys.exit(1)

    # ── STEP 5: 양식 붙여넣기 ────────────────────────
    try:
        logger.info("[STEP 5] 양식 붙여넣기")
        writer = TemplateWriter(config, logger)
        output_path = writer.write(df, month)
        result_summary["write"] = str(output_path)
    except Exception as e:
        logger.critical(f"양식 기입 오류: {e}")
        result_summary["write"] = f"failed: {e}"
        sys.exit(1)

    # ── 완료 요약 ────────────────────────────────────
    elapsed = time.time() - start_time
    logger.info("=" * 60)
    logger.info("자동화 완료 요약")
    logger.info(f"  기준월    : {result_summary['month']}")
    logger.info(f"  SAP 추출  : {result_summary['sap']}")
    logger.info(f"  정제 결과 : {result_summary['process']}")
    logger.info(f"  산출 파일 : {result_summary['write']}")
    logger.info(f"  소요 시간 : {elapsed:.1f}초")
    logger.info("=" * 60)
    logger.info("※ 최종 결과 파일을 열어 이상값 여부를 검수하세요.")


if __name__ == "__main__":
    main()
