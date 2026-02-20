"""Tests for password encryption/decryption functions."""
import os
import tempfile
import pytest


def _setup_env(tmp_path):
    """Set up isolated config/key environment."""
    import mailgui
    mailgui.CONFIG_FILE = str(tmp_path / "config.json")
    mailgui.KEY_FILE = str(tmp_path / ".mailgui.key")
    # Remove existing key if any
    if os.path.exists(mailgui.KEY_FILE):
        os.remove(mailgui.KEY_FILE)


class TestGetOrCreateKey:
    def test_creates_key_file(self, tmp_path):
        _setup_env(tmp_path)
        import mailgui
        key = mailgui._get_or_create_key()
        assert os.path.exists(mailgui.KEY_FILE)
        assert len(key) > 0

    def test_returns_same_key(self, tmp_path):
        _setup_env(tmp_path)
        import mailgui
        key1 = mailgui._get_or_create_key()
        key2 = mailgui._get_or_create_key()
        assert key1 == key2

    def test_key_file_permissions(self, tmp_path):
        _setup_env(tmp_path)
        import mailgui
        mailgui._get_or_create_key()
        stat = os.stat(mailgui.KEY_FILE)
        assert oct(stat.st_mode & 0o777) == "0o600"


class TestEncryptPassword:
    def test_encrypt_decrypt_roundtrip(self, tmp_path):
        _setup_env(tmp_path)
        import mailgui
        password = "my-secret-password"
        encrypted = mailgui._encrypt_password(password)
        assert encrypted != password
        decrypted = mailgui._decrypt_password(encrypted)
        assert decrypted == password

    def test_empty_password(self, tmp_path):
        _setup_env(tmp_path)
        import mailgui
        assert mailgui._encrypt_password("") == ""
        assert mailgui._decrypt_password("") == ""

    def test_unicode_password(self, tmp_path):
        _setup_env(tmp_path)
        import mailgui
        password = "密碼測試🔐"
        encrypted = mailgui._encrypt_password(password)
        decrypted = mailgui._decrypt_password(encrypted)
        assert decrypted == password

    def test_different_encryptions_differ(self, tmp_path):
        """Fernet uses random IV, so same plaintext produces different ciphertext."""
        _setup_env(tmp_path)
        import mailgui
        password = "test123"
        enc1 = mailgui._encrypt_password(password)
        enc2 = mailgui._encrypt_password(password)
        assert enc1 != enc2  # Different IVs
        assert mailgui._decrypt_password(enc1) == password
        assert mailgui._decrypt_password(enc2) == password


class TestDecryptPasswordBackwardCompat:
    def test_plaintext_fallback(self, tmp_path):
        """If decryption fails, return as-is (backward compatibility)."""
        _setup_env(tmp_path)
        import mailgui
        plaintext = "old-plaintext-password"
        result = mailgui._decrypt_password(plaintext)
        assert result == plaintext

    def test_invalid_token_fallback(self, tmp_path):
        _setup_env(tmp_path)
        import mailgui
        result = mailgui._decrypt_password("not-a-valid-fernet-token")
        assert result == "not-a-valid-fernet-token"
