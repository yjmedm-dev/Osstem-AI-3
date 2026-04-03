"""1C:Enterprise 8.3 OData API 클라이언트.

사용법:
    client = OneCODataClient(
        base_url="http://서버주소/베이스명",
        username="사용자명",
        password="비밀번호",
    )
    tb = client.fetch_trial_balance("UZ01", "2025-03")
"""

import requests
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

from models.trial_balance import TrialBalance, TrialBalanceRow
from utils.currency import convert
from utils.exceptions import OsstemBaseError
from config.settings import ANTHROPIC_API_KEY  # noqa — settings 로드 트리거


class OneCConnectionError(OsstemBaseError):
    """1C OData 연결 실패."""


class OneCODataClient:
    """1C:Enterprise 8.3 OData API 클라이언트."""

    def __init__(
        self,
        base_url: str,
        username: str,
        password: str,
        verify_ssl: bool = True,
        timeout: int = 30,
    ):
        """
        Args:
            base_url: 1C 웹 게시 URL (예: "http://192.168.1.10/UZ_Base")
            username: 1C 사용자명
            password: 1C 비밀번호
            verify_ssl: SSL 인증서 검증 여부 (내부망은 False 가능)
            timeout: 요청 타임아웃 (초)
        """
        self.odata_url = base_url.rstrip("/") + "/odata/standard.odata"
        self.auth = (username, password)
        self.verify_ssl = verify_ssl
        self.timeout = timeout
        self._session = requests.Session()
        self._session.auth = self.auth
        self._session.verify = self.verify_ssl

    # ──────────────────────────────────────────────────────────────
    # 공개 메서드
    # ──────────────────────────────────────────────────────────────

    def test_connection(self) -> bool:
        """연결 테스트. True 반환 시 정상."""
        try:
            resp = self._get("", params={"$format": "json", "$top": "1"})
            return resp.status_code == 200
        except Exception:
            return False

    def fetch_trial_balance(
        self,
        subsidiary_code: str,
        period: str,
        org_name: str | None = None,
        exchange_rate: float = 1.0,
        currency: str = "UZS",
    ) -> TrialBalance:
        """시산표 데이터를 가져와 TrialBalance 객체로 반환.

        Args:
            subsidiary_code: 법인 코드 (예: "UZ01")
            period: 기간 문자열 (예: "2025-03")
            org_name: 1C 내 조직명 (None이면 전체 조직 합산)
            exchange_rate: 원화 환산 환율 (UZS→KRW)
            currency: 원본 통화 코드
        """
        year, month = map(int, period.split("-"))
        date_from = date(year, month, 1)
        date_to = date_from + relativedelta(months=1) - relativedelta(days=1)

        raw_df = self._fetch_balance_turnover(date_from, date_to, org_name)

        if raw_df.empty:
            raise OneCConnectionError(
                f"[{subsidiary_code}] {period} 데이터가 없습니다. "
                "조직명(org_name) 또는 기간을 확인하세요."
            )

        rows = self._to_trial_balance_rows(
            raw_df, subsidiary_code, period, exchange_rate, currency
        )

        return TrialBalance(
            subsidiary_code=subsidiary_code,
            period=period,
            rows=rows,
            source_file=f"1C OData ({self.odata_url})",
        )

    def fetch_receivables(self) -> pd.DataFrame:
        """매출채권 잔액 조회 (대손충당금 Aging 분석용)."""
        params = {
            "$format": "json",
            "$select": ",".join([
                "Контрагент_Key",
                "Контрагент",
                "ДоговорКонтрагента",
                "СуммаОстаток",
            ]),
        }
        resp = self._get("РасчетыСКонтрагентамиОстатки", params=params)
        return pd.DataFrame(resp.json().get("value", []))

    def fetch_inventory(self) -> pd.DataFrame:
        """재고 잔액 조회 (재고자산평가충당금·단품충당금 검증용)."""
        params = {
            "$format": "json",
            "$select": ",".join([
                "Номенклатура_Key",
                "Номенклатура",
                "Склад",
                "КоличествоОстаток",
                "СуммаОстаток",
            ]),
        }
        resp = self._get("ТоварыНаСкладахОстатки", params=params)
        return pd.DataFrame(resp.json().get("value", []))

    def list_organizations(self) -> list[str]:
        """1C에 등록된 조직 목록 반환 (org_name 확인용)."""
        params = {
            "$format": "json",
            "$select": "Description",
        }
        resp = self._get("Организация", params=params)
        return [r["Description"] for r in resp.json().get("value", [])]

    # ──────────────────────────────────────────────────────────────
    # 내부 메서드
    # ──────────────────────────────────────────────────────────────

    def _fetch_balance_turnover(
        self,
        date_from: date,
        date_to: date,
        org_name: str | None,
    ) -> pd.DataFrame:
        """ХозрасчетныйОстаткиИОбороты (시산표) 조회."""

        # OData datetime 형식
        dt_from = date_from.strftime("%Y-%m-%dT00:00:00")
        dt_to   = date_to.strftime("%Y-%m-%dT23:59:59")

        filters = [
            f"Период ge datetime'{dt_from}'",
            f"Период le datetime'{dt_to}'",
        ]
        if org_name:
            filters.append(f"Организация eq '{org_name}'")

        params = {
            "$format": "json",
            "$filter": " and ".join(filters),
            "$select": ",".join([
                "Счет_Key",
                "Счет",
                "СуммаНачальныйОстатокДт",
                "СуммаНачальныйОстатокКт",
                "СуммаОборотДт",
                "СуммаОборотКт",
                "СуммаКонечныйОстатокДт",
                "СуммаКонечныйОстатокКт",
            ]),
        }

        resp = self._get("ХозрасчетныйОстаткиИОбороты", params=params)
        return pd.DataFrame(resp.json().get("value", []))

    def _to_trial_balance_rows(
        self,
        df: pd.DataFrame,
        subsidiary_code: str,
        period: str,
        exchange_rate: float,
        currency: str,
    ) -> list[TrialBalanceRow]:
        """1C DataFrame → TrialBalanceRow 리스트 변환."""
        rows = []
        for _, r in df.iterrows():
            # 기말 잔액 기준 (차변/대변)
            debit_orig  = float(r.get("СуммаКонечныйОстатокДт") or 0)
            credit_orig = float(r.get("СуммаКонечныйОстатокКт") or 0)

            debit_krw  = convert(debit_orig,  currency, exchange_rate)
            credit_krw = convert(credit_orig, currency, exchange_rate)

            rows.append(TrialBalanceRow(
                subsidiary_code=subsidiary_code,
                period=period,
                account_code=str(r.get("Счет_Key", "")).strip(),
                account_name=str(r.get("Счет", "")).strip(),
                debit=debit_krw,
                credit=credit_krw,
                original_amount=debit_orig - credit_orig,
                original_currency=currency,
                exchange_rate=exchange_rate,
            ))
        return rows

    def _get(self, endpoint: str, params: dict | None = None) -> requests.Response:
        """GET 요청 공통 처리."""
        url = f"{self.odata_url}/{endpoint}" if endpoint else self.odata_url
        try:
            resp = self._session.get(url, params=params, timeout=self.timeout)
            resp.raise_for_status()
            return resp
        except requests.exceptions.ConnectionError as e:
            raise OneCConnectionError(f"1C 서버 연결 실패: {e}") from e
        except requests.exceptions.HTTPError as e:
            raise OneCConnectionError(
                f"1C API 오류 [{resp.status_code}]: {resp.text[:200]}"
            ) from e
        except requests.exceptions.Timeout as e:
            raise OneCConnectionError(
                f"1C 서버 응답 타임아웃 ({self.timeout}초)"
            ) from e
