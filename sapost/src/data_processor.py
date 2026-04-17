"""
데이터 정제/가공 모듈
SAP에서 추출된 데이터를 읽어 단일 DataFrame으로 정제 후 반환
"""
import pandas as pd
import configparser
import logging
from pathlib import Path


class DataProcessor:
    def __init__(self, config: configparser.ConfigParser, logger: logging.Logger):
        self.config = config
        self.logger = logger

        self.header_row = config.getint("EXCEL_MAPPING", "header_row", fallback=1)
        self.total_keyword = config.get("EXCEL_MAPPING", "total_row_keyword", fallback="합계")

        skip_raw = config.get("EXCEL_MAPPING", "skip_columns", fallback="")
        self.skip_columns = [c.strip() for c in skip_raw.split(",") if c.strip()]

        intermediate_dir = config.get("PATHS", "intermediate_dir", fallback="sapost/data/intermediate")
        self.intermediate_dir = Path(intermediate_dir)
        self.intermediate_dir.mkdir(parents=True, exist_ok=True)

    def process(self, file_paths: list[Path], month: str) -> pd.DataFrame:
        """
        file_paths: 추출된 엑셀 파일 경로 목록
        month: 기준월 (예: "202503")
        반환: 정제된 전체 데이터 DataFrame
        """
        checkpoint = self.intermediate_dir / f"cleaned_{month}.pkl"
        if checkpoint.exists():
            self.logger.info(f"중간 결과 파일 재사용: {checkpoint}")
            return pd.read_pickle(checkpoint)

        all_frames = []
        failed = []

        for i, path in enumerate(file_paths, start=1):
            self.logger.info(f"[{i}/{len(file_paths)}] 처리 중: {path.name}")
            try:
                df = self._read_single_file(path)
                df["_source_file"] = path.name
                all_frames.append(df)
            except Exception as e:
                self.logger.error(f"파일 읽기 실패 ({path.name}): {e}")
                failed.append(path.name)

        if failed:
            self.logger.warning(f"처리 실패 파일 {len(failed)}건: {failed}")

        if not all_frames:
            raise ValueError("정제 가능한 파일이 없습니다.")

        result = pd.concat(all_frames, ignore_index=True)
        self.logger.info(f"전체 정제 완료: {len(result)}행")

        result.to_pickle(checkpoint)
        self.logger.info(f"중간 결과 저장: {checkpoint}")

        return result

    def process_dataframe(self, df: pd.DataFrame, month: str) -> pd.DataFrame:
        """
        ALV 직접 읽기 방식: 이미 메모리에 있는 DataFrame을 정제
        extract_mode = grid 일 때 sap_controller가 반환한 DataFrame을 받아 처리
        """
        checkpoint = self.intermediate_dir / f"cleaned_{month}.pkl"
        if checkpoint.exists():
            self.logger.info(f"중간 결과 파일 재사용: {checkpoint}")
            return pd.read_pickle(checkpoint)

        df = self._clean_dataframe(df)
        self.logger.info(f"정제 완료: {len(df)}행")

        df.to_pickle(checkpoint)
        self.logger.info(f"중간 결과 저장: {checkpoint}")

        return df

    def _read_single_file(self, path: Path) -> pd.DataFrame:
        """단일 엑셀 파일을 읽고 정제하여 반환"""
        df = pd.read_excel(path, header=self.header_row - 1, dtype=str)
        return self._clean_dataframe(df)

    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """공통 정제 로직"""
        df.columns = [str(c).strip() for c in df.columns]

        if self.skip_columns:
            existing_skip = [c for c in self.skip_columns if c in df.columns]
            df = df.drop(columns=existing_skip)

        if len(df.columns) > 0:
            first_col = df.columns[0]
            df = df[~df[first_col].astype(str).str.contains(self.total_keyword, na=False)]

        df = df.dropna(how="all")
        df = self._convert_numeric_columns(df)

        return df.reset_index(drop=True)

    def _convert_numeric_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """쉼표가 있는 숫자 문자열을 float으로 변환 시도"""
        for col in df.columns:
            cleaned = df[col].astype(str).str.replace(",", "").str.strip()
            try:
                converted = pd.to_numeric(cleaned, errors="raise")
                df[col] = converted
            except (ValueError, TypeError):
                pass
        return df
