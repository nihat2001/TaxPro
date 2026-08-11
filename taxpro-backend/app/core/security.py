from cryptography.fernet import Fernet

from app.core.config import get_settings

_PREFIX = "enc:"


class EncryptionService:
    def __init__(self, key: str) -> None:
        if not key:
            raise ValueError("ENCRYPTION_KEY mühit dəyişəni təyin edilməyib.")
        self._fernet = Fernet(key.encode("utf-8"))

    def encrypt(self, value: str) -> str:
        token = self._fernet.encrypt(value.encode("utf-8"))
        return _PREFIX + token.decode("utf-8")

    def decrypt(self, value: str) -> str:
        if not value.startswith(_PREFIX):
            return value
        token = value[len(_PREFIX):].encode("utf-8")
        return self._fernet.decrypt(token).decode("utf-8")


_encryption_service: EncryptionService | None = None


def get_encryption_service() -> EncryptionService:
    global _encryption_service
    if _encryption_service is None:
        _encryption_service = EncryptionService(get_settings().encryption_key)
    return _encryption_service


def encrypt_value(value: str) -> str:
    return get_encryption_service().encrypt(value)


def decrypt_value(value: str) -> str:
    return get_encryption_service().decrypt(value)


def mask_value(value: str | None) -> str:
    return "****" if value else ""
