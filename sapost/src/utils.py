"""
공통 유틸리티: 로깅 설정, 재시도 데코레이터, 파일 대기 함수
"""
import logging
import time
import functools
import configparser
from pathlib import Path
from datetime import datetime


def get_config(config_path: str = None) -> configparser.ConfigParser:
    """config.ini 파일을 읽어 반환"""
    if config_path is None:
        base = Path(__file__).parent.parent
        config_path = base / "config" / "config.ini"

    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")
    return config


def setup_logger(name: str, config: configparser.ConfigParser) -> logging.Logger:
    """날짜별 로그 파일 + 콘솔 동시 출력 로거 생성"""
    log_level = config.get("LOGGING", "level", fallback="INFO")
    log_dir = Path(config.get("LOGGING", "log_dir", fallback="sapost/logs"))
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"{datetime.now().strftime('%Y%m%d')}.log"

    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, log_level))

    if not logger.handlers:
        fmt = logging.Formatter(
            "[%(asctime)s] %(levelname)-8s %(name)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        logger.addHandler(ch)

    return logger


def retry(max_attempts: int = 3, delay: float = 2.0, exceptions: tuple = (Exception,)):
    """
    재시도 데코레이터.
    max_attempts 횟수만큼 재시도하고 모두 실패하면 마지막 예외를 발생시킴.
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exc = e
                    if attempt < max_attempts:
                        time.sleep(delay)
            raise last_exc
        return wrapper
    return decorator


def wait_for_file(directory: Path, timeout: float = 30.0, poll: float = 0.5) -> Path:
    """
    directory에 새 .xlsx 파일이 생성될 때까지 대기 후 경로 반환.
    .tmp / .crdownload 등 임시 파일이 사라진 뒤 완성된 파일만 반환.
    timeout 초 안에 파일이 없으면 TimeoutError 발생.
    """
    deadline = time.time() + timeout
    before = set(directory.glob("*.xlsx"))

    while time.time() < deadline:
        current = set(directory.glob("*.xlsx"))
        new_files = current - before
        temp_files = list(directory.glob("*.tmp")) + list(directory.glob("*.crdownload"))
        if new_files and not temp_files:
            return new_files.pop()
        time.sleep(poll)

    raise TimeoutError(f"다운로드 완료 파일을 {timeout}초 안에 찾지 못했습니다: {directory}")
