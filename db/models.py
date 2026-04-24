from datetime import datetime

from sqlalchemy import (
    BigInteger, Boolean, Column, DateTime, Integer,
    Numeric, String, Text,
)
from sqlalchemy.orm import DeclarativeBase


class Base(DeclarativeBase):
    pass


class AccountMaster(Base):
    """3개 시스템(현지회계/네트라/Confinas) 계정과목 매핑 마스터"""
    __tablename__ = "account_master"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    subsidiary_code = Column(String(10),  nullable=False, index=True)
    local_code      = Column(String(50),  nullable=True)
    local_name      = Column(String(200), nullable=True)
    # 네트라는 개별 계정코드 없이 5개 항목 합계만 제공
    # 값: 매출채권 | 선수금 | 원가 | 재고자산 | 매출액 | NULL(비대상)
    netra_category  = Column(String(20),  nullable=True)
    confinas_code   = Column(String(50),  nullable=True)
    confinas_name   = Column(String(200), nullable=True)
    standard_code   = Column(String(50),  nullable=True)   # FP/PL 신계정
    account_type    = Column(String(20),  nullable=True)   # asset/liability/…
    created_at      = Column(DateTime, default=datetime.utcnow)
    updated_at      = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class FinancialLocal(Base):
    """현지회계프로그램에서 업로드된 재무 데이터"""
    __tablename__ = "financial_local"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    subsidiary_code = Column(String(10),  nullable=False, index=True)
    period          = Column(String(7),   nullable=False, index=True)   # '2025-03'
    account_code    = Column(String(50),  nullable=False)
    account_name    = Column(String(200), nullable=True)
    debit           = Column(Numeric(20, 2), default=0)
    credit          = Column(Numeric(20, 2), default=0)
    balance         = Column(Numeric(20, 2), default=0)    # debit - credit
    currency        = Column(String(10),  nullable=True)
    exchange_rate   = Column(Numeric(15, 6), nullable=True)
    amount_krw      = Column(Numeric(20, 2), default=0)
    # 계정 레벨: 1=주계정(끝2자리 00), 2=보조계정(4자리 비00), 3=세부계정(소수점)
    local_level     = Column(Integer, nullable=True)
    uploaded_at     = Column(DateTime, default=datetime.utcnow)


class FinancialNetra(Base):
    """네트라(ERP) 5개 항목 합계 데이터
    category: 매출채권 | 선수금 | 원가 | 재고자산 | 매출액
    """
    __tablename__ = "financial_netra"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    subsidiary_code = Column(String(10),  nullable=False, index=True)
    period          = Column(String(7),   nullable=False, index=True)
    category        = Column(String(20),  nullable=False)  # 5개 항목명
    amount          = Column(Numeric(20, 2), default=0)    # 현지통화 금액
    currency        = Column(String(10),  nullable=True)
    exchange_rate   = Column(Numeric(15, 6), nullable=True)
    amount_krw      = Column(Numeric(20, 2), default=0)    # 원화 환산액
    uploaded_at     = Column(DateTime, default=datetime.utcnow)


class UploadLog(Base):
    """각 시스템별 업로드 이력 및 검증 결과"""
    __tablename__ = "upload_log"

    id              = Column(Integer, primary_key=True, autoincrement=True)
    system_name     = Column(String(20),  nullable=False)  # 'local' | 'netra' | 'confinas'
    subsidiary_code = Column(String(10),  nullable=False, index=True)
    period          = Column(String(7),   nullable=False)
    status          = Column(String(20),  nullable=False)  # 'success' | 'error'
    row_count       = Column(Integer,     default=0)
    total_debit     = Column(Numeric(20, 2), nullable=True)
    total_credit    = Column(Numeric(20, 2), nullable=True)
    is_balanced     = Column(Boolean,     nullable=True)
    message         = Column(Text,        nullable=True)
    uploaded_at     = Column(DateTime, default=datetime.utcnow)
