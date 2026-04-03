class OsstemBaseError(Exception):
    """프로젝트 공통 기반 예외."""


class FileNotSubmittedError(OsstemBaseError):
    """법인이 파일을 제출하지 않은 경우."""


class InvalidFileFormatError(OsstemBaseError):
    """파일 형식 또는 시트 구조가 맞지 않는 경우."""


class SchemaMapError(OsstemBaseError):
    """계정 코드 매핑 실패."""


class CurrencyConversionError(OsstemBaseError):
    """환율 변환 실패."""


class ValidationCriticalError(OsstemBaseError):
    """CRITICAL 검증 오류 발생으로 처리를 중단해야 하는 경우."""
