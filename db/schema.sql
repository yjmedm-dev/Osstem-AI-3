-- Osstem AI — 3-System Reconciliation Schema
-- DB: yeji  |  MySQL 8.0+
-- 실행: mysql -h 10.190.6.99 -P 5010 -u yeji -p yeji < db/schema.sql

USE yeji;

-- 1. 3개 시스템 계정과목 매핑 마스터
CREATE TABLE IF NOT EXISTS account_master (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    subsidiary_code VARCHAR(10)  NOT NULL COMMENT '법인 코드 (UZ01, RU01 등)',
    local_code      VARCHAR(50)  COMMENT '현지회계 계정코드',
    local_name      VARCHAR(200) COMMENT '현지회계 계정명',
    netra_category  VARCHAR(20)  COMMENT '네트라 대사 항목: 매출채권|선수금|원가|재고자산|매출액|NULL',
    confinas_code   VARCHAR(50)  COMMENT 'Confinas 계정코드',
    confinas_name   VARCHAR(200) COMMENT 'Confinas 계정명',
    standard_code   VARCHAR(50)  COMMENT '내부 신계정 (FP/PL 체계)',
    account_type    VARCHAR(20)  COMMENT 'asset/liability/equity/revenue/expense',
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_subsidiary (subsidiary_code),
    INDEX idx_local     (subsidiary_code, local_code),
    INDEX idx_netra_cat (subsidiary_code, netra_category)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 2. 현지회계프로그램 데이터
CREATE TABLE IF NOT EXISTS financial_local (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    subsidiary_code VARCHAR(10)     NOT NULL,
    period          VARCHAR(7)      NOT NULL COMMENT '예: 2025-03',
    account_code    VARCHAR(50)     NOT NULL,
    account_name    VARCHAR(200),
    debit           DECIMAL(20, 2)  DEFAULT 0,
    credit          DECIMAL(20, 2)  DEFAULT 0,
    balance         DECIMAL(20, 2)  DEFAULT 0 COMMENT 'debit - credit',
    currency        VARCHAR(10),
    exchange_rate   DECIMAL(15, 6),
    amount_krw      DECIMAL(20, 2)  DEFAULT 0,
    local_level     TINYINT         COMMENT '1=주계정(끝00) 2=보조계정(4자리) 3=세부계정(소수점)',
    uploaded_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_corp_period (subsidiary_code, period)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 3. 네트라 5개 항목 합계 데이터
CREATE TABLE IF NOT EXISTS financial_netra (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    subsidiary_code VARCHAR(10)     NOT NULL,
    period          VARCHAR(7)      NOT NULL,
    category        VARCHAR(20)     NOT NULL COMMENT '매출채권|선수금|원가|재고자산|매출액',
    amount          DECIMAL(20, 2)  DEFAULT 0 COMMENT '현지통화 금액',
    currency        VARCHAR(10),
    exchange_rate   DECIMAL(15, 6),
    amount_krw      DECIMAL(20, 2)  DEFAULT 0 COMMENT '원화 환산액',
    uploaded_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_corp_period (subsidiary_code, period)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 4. 업로드 이력 및 검증 결과
CREATE TABLE IF NOT EXISTS upload_log (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    system_name     VARCHAR(20)  NOT NULL COMMENT 'local | netra | confinas',
    subsidiary_code VARCHAR(10)  NOT NULL,
    period          VARCHAR(7)   NOT NULL,
    status          VARCHAR(20)  NOT NULL COMMENT 'success | error',
    row_count       INT          DEFAULT 0,
    total_debit     DECIMAL(20, 2),
    total_credit    DECIMAL(20, 2),
    is_balanced     TINYINT(1)   COMMENT '대차균형 여부',
    message         TEXT,
    uploaded_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_corp_period (subsidiary_code, period)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
