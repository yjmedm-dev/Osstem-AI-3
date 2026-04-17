"""
SAP GUI 조작 모듈 (설치형 SAP GUI, COM 기반)

SAP GUI 스크립팅 API(win32com)를 사용하여 SAP 화면을 자동 조작합니다.

사용 전 준비:
  1. SAP GUI 실행 후 로그온 상태 유지
  2. SAP GUI 옵션 → 접근성 → 스크립팅 활성화 체크
     (또는 SAP GUI 상단 메뉴 → Customize Local Layout → Options → Accessibility & Scripting)
  3. SAP GUI 레코더로 실제 트랜잭션 녹화 → config.ini 컨트롤 ID 확인
     (SAP GUI 상단 메뉴 → Help → SAP GUI Scripting → Record and Playback)
"""
import os
import time
import shutil
import logging
import configparser
from pathlib import Path
from dotenv import load_dotenv

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

import pandas as pd


class SAPController:
    """
    SAP GUI 자동 조작 클래스.
    win32com COM 인터페이스를 통해 SAP GUI Scripting API를 호출합니다.

    기본 흐름:
        connect() → login() → navigate_to() → set_params_and_execute()
        → get_data() or export_to_file() → close()
    """

    def __init__(self, config: configparser.ConfigParser, logger: logging.Logger):
        if not WIN32COM_AVAILABLE:
            raise ImportError(
                "pywin32가 설치되어 있지 않습니다.\n"
                "pip install pywin32==306  을 실행하세요."
            )

        self.config = config
        self.logger = logger
        self.session = None

        # .env 에서 로그인 정보 로드
        env_path = Path(__file__).parent.parent / "config" / ".env"
        load_dotenv(dotenv_path=env_path)
        self.user_id   = os.getenv("SAP_USER_ID", "")
        self.password  = os.getenv("SAP_PASSWORD", "")
        self.client    = os.getenv("SAP_CLIENT", "")
        self.language  = os.getenv("SAP_LANGUAGE", "KO")

        if not self.user_id or not self.password:
            raise ValueError(".env 파일에 SAP_USER_ID / SAP_PASSWORD가 설정되지 않았습니다.")

        # config.ini 에서 SAP 설정 로드
        self.transaction    = config.get("SAP", "transaction")
        self.grid_id        = config.get("SAP", "grid_id")
        self.month_field_id = config.get("SAP", "month_field_id")
        self.execute_vkey   = config.getint("SAP", "execute_vkey", fallback=8)
        self.extract_mode   = config.get("SAP", "extract_mode", fallback="grid")

        self.download_dir  = Path(config.get("PATHS", "download_dir",
                                             fallback="C:/Users/Osstem/Downloads"))
        self.raw_dir       = Path(config.get("PATHS", "raw_dir", fallback="sapost/data/raw"))
        self.raw_dir.mkdir(parents=True, exist_ok=True)

    # ──────────────────────────────────────────────
    # 공개 메서드
    # ──────────────────────────────────────────────

    def connect(self):
        """
        실행 중인 SAP GUI 세션에 연결합니다.
        SAP GUI가 실행 중이고 로그온 상태여야 합니다.
        """
        self.logger.info("SAP GUI 세션 연결 중...")
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application  = sap_gui_auto.GetScriptingEngine
            connection   = application.Children(0)   # 첫 번째 연결
            self.session = connection.Children(0)    # 첫 번째 세션
            self.logger.info("SAP GUI 세션 연결 완료")
        except Exception as e:
            raise ConnectionError(
                f"SAP GUI 세션에 연결할 수 없습니다: {e}\n"
                "SAP GUI가 실행 중이고 스크립팅이 활성화되어 있는지 확인하세요."
            )

    def login(self):
        """
        SAP 로그인 화면이 표시된 경우 자격증명을 입력하고 로그인합니다.
        이미 로그인되어 있으면 건너뜁니다.
        """
        self.logger.info("SAP 로그인 상태 확인")
        try:
            # 현재 화면 타입 확인 — 로그인 화면은 transaction = "SLOGIN" 또는 빈 화면
            current_transaction = self.session.Info.Transaction
            if current_transaction.upper() in ("SLOGIN", ""):
                self.logger.info("로그인 화면 감지 — 자격증명 입력 중")
                self._fill_login_screen()
            else:
                self.logger.info(f"이미 로그인됨 (현재 트랜잭션: {current_transaction})")
        except Exception as e:
            self.logger.warning(f"로그인 상태 확인 실패, 로그인 시도: {e}")
            self._fill_login_screen()

    def navigate_to(self):
        """config.ini 에 설정된 트랜잭션으로 이동합니다."""
        self.logger.info(f"트랜잭션 이동: {self.transaction}")
        try:
            cmd_field = self.session.findById("wnd[0]/tbar[0]/okcd")
            cmd_field.text = f"/n{self.transaction}"
            self.session.findById("wnd[0]").sendVKey(0)  # Enter
            time.sleep(2)
            self.logger.info("트랜잭션 이동 완료")
        except Exception as e:
            raise RuntimeError(f"트랜잭션 이동 실패 ({self.transaction}): {e}")

    def set_params_and_execute(self, month: str):
        """
        기준월 등 조회 파라미터를 입력하고 실행합니다.
        month: 'YYYYMM' 형식 (예: '202503')
        """
        self.logger.info(f"기준월 입력: {month}")
        try:
            month_field = self.session.findById(self.month_field_id)
            month_field.text = month
            self.session.findById("wnd[0]").sendVKey(self.execute_vkey)  # F8
            time.sleep(3)
            self.logger.info("조회 실행 완료")
        except Exception as e:
            raise RuntimeError(f"파라미터 입력/실행 실패: {e}")

    def get_data(self) -> pd.DataFrame:
        """
        extract_mode 에 따라 데이터를 추출합니다.
          - grid  : ALV 그리드를 직접 읽어 DataFrame 반환
          - export: SAP 내보내기 버튼 → 다운로드 파일 경로 반환 (파일은 raw/ 에 이동)
        반환: DataFrame (grid 모드) 또는 None (export 모드 — raw/ 파일 사용)
        """
        if self.extract_mode == "grid":
            return self._read_alv_grid()
        elif self.extract_mode == "export":
            return None  # main.py 에서 raw/ 파일 목록으로 처리
        else:
            raise ValueError(f"알 수 없는 extract_mode: {self.extract_mode}")

    def export_to_file(self, month: str) -> Path:
        """
        extract_mode = export 일 때:
        SAP ALV 툴바의 내보내기 버튼을 클릭하고 다운로드된 파일을 raw/ 로 이동합니다.
        반환: raw/ 에 저장된 파일 경로
        """
        from .utils import wait_for_file
        self.logger.info("SAP 내보내기 버튼 클릭")
        try:
            export_btn_id = self.config.get("SAP", "export_btn_id",
                                            fallback="wnd[0]/tbar[1]/btn[45]")
            self.session.findById(export_btn_id).press()
            time.sleep(1)

            # 내보내기 대화상자 처리 — "스프레드시트" 선택 후 확인
            self._handle_export_dialog()

            # 다운로드 완료 대기
            downloaded = wait_for_file(self.download_dir, timeout=30.0)

            # raw/ 로 이동
            dest = self.raw_dir / f"{month}_001.xlsx"
            shutil.move(str(downloaded), str(dest))
            self.logger.info(f"파일 이동 완료: {dest}")
            return dest

        except Exception as e:
            raise RuntimeError(f"내보내기 실패: {e}")

    def close(self):
        """세션을 초기 화면으로 복귀시킵니다 (SAP 종료 아님)."""
        try:
            if self.session:
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                self.session.findById("wnd[0]").sendVKey(0)
                self.logger.info("SAP 초기 화면으로 복귀 완료")
        except Exception:
            pass

    # ──────────────────────────────────────────────
    # 내부 메서드
    # ──────────────────────────────────────────────

    def _fill_login_screen(self):
        """SAP 로그인 화면에 자격증명 입력"""
        try:
            if self.client:
                self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = self.client
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text  = self.user_id
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text  = self.password
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text  = self.language
            self.session.findById("wnd[0]").sendVKey(0)  # Enter
            time.sleep(3)
            self.logger.info("SAP 로그인 완료")
        except Exception as e:
            raise RuntimeError(f"SAP 로그인 실패: {e}")

    def _read_alv_grid(self) -> pd.DataFrame:
        """
        ALV 그리드 컨트롤에서 모든 행/컬럼 데이터를 읽어 DataFrame으로 반환합니다.

        ※ SAP GUI 레코더로 grid_id를 확인 후 config.ini 에 입력하세요.
           일반적인 ALV 그리드 ID: wnd[0]/usr/cntlGRID1/shellcont/shell
        """
        self.logger.info("ALV 그리드 데이터 읽기 시작")
        try:
            grid = self.session.findById(self.grid_id)
        except Exception as e:
            raise RuntimeError(
                f"ALV 그리드를 찾을 수 없습니다 (ID: {self.grid_id}): {e}\n"
                "config.ini 의 grid_id 를 SAP GUI 레코더로 확인하세요."
            )

        row_count = grid.RowCount
        self.logger.info(f"ALV 그리드 행 수: {row_count}")

        if row_count == 0:
            self.logger.warning("ALV 그리드에 데이터가 없습니다.")
            return pd.DataFrame()

        # 컬럼 목록 수집
        columns = self._get_grid_columns(grid)
        self.logger.info(f"컬럼 수: {len(columns)}")

        # 행 데이터 수집
        records = []
        for row in range(row_count):
            record = {}
            for col_name in columns:
                try:
                    record[col_name] = grid.GetCellValue(row, col_name)
                except Exception:
                    record[col_name] = None
            records.append(record)

        df = pd.DataFrame(records, columns=columns)
        self.logger.info(f"ALV 그리드 읽기 완료: {len(df)}행 × {len(df.columns)}열")
        return df

    def _get_grid_columns(self, grid) -> list[str]:
        """ALV 그리드의 컬럼명 목록을 반환합니다."""
        columns = []
        try:
            col_order = grid.ColumnOrder
            for i in range(len(col_order)):
                columns.append(col_order.ElementAt(i))
        except Exception:
            # ColumnOrder 를 지원하지 않는 경우 — 컬럼 수를 직접 순회
            try:
                col_count = grid.ColumnCount
                for i in range(col_count):
                    columns.append(grid.GetColumnKeyName(i))
            except Exception as e:
                raise RuntimeError(
                    f"ALV 그리드 컬럼 목록을 읽을 수 없습니다: {e}\n"
                    "SAP GUI 레코더로 컬럼명을 직접 확인 후 코드를 수정하세요."
                )
        return columns

    def _handle_export_dialog(self):
        """SAP 내보내기 대화상자 처리 — '스프레드시트' 항목 선택 후 확인"""
        time.sleep(1)
        try:
            # 대화상자가 열린 경우 "스프레드시트" 항목 선택
            # (SAP 버전에 따라 대화상자 ID가 다를 수 있음 — 레코더로 확인 필요)
            dialog = self.session.findById("wnd[1]")
            # 스프레드시트 항목 더블클릭 (일반적인 위치)
            try:
                dialog.findById("usr/cmbG_LISTBOX").key = "spreadsheetml"
            except Exception:
                pass  # 이미 기본값이 스프레드시트인 경우
            dialog.findById("tbar[0]/btn[0]").press()  # 확인
            time.sleep(2)
        except Exception:
            # 대화상자가 없으면 바로 다운로드됨
            pass
