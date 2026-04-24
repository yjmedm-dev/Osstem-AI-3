from contextlib import contextmanager

from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker, Session

from config.settings import MYSQL_HOST, MYSQL_PORT, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DB

_ENGINE = None


def get_engine():
    global _ENGINE
    if _ENGINE is None:
        url = (
            f"mysql+pymysql://{MYSQL_USER}:{MYSQL_PASSWORD}"
            f"@{MYSQL_HOST}:{MYSQL_PORT}/{MYSQL_DB}"
            f"?charset=utf8mb4"
        )
        _ENGINE = create_engine(url, pool_pre_ping=True, echo=False)
    return _ENGINE


@contextmanager
def get_session() -> Session:
    factory = sessionmaker(bind=get_engine(), expire_on_commit=False)
    session = factory()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def test_connection() -> bool:
    try:
        with get_engine().connect() as conn:
            conn.execute(text("SELECT 1"))
        return True
    except Exception:
        return False
