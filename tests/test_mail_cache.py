"""Tests for MailCache SQLite persistent cache."""
import os
import sqlite3
import tempfile
import pytest


def _make_cache(tmp_path):
    """Create a MailCache pointing to a temp directory."""
    import mailgui
    mailgui.CONFIG_FILE = str(tmp_path / "config.json")
    cache = mailgui.MailCache()
    return cache


class TestMailCacheInit:
    def test_creates_db(self, tmp_path):
        cache = _make_cache(tmp_path)
        db_path = os.path.join(str(tmp_path), "mail_cache.db")
        assert os.path.exists(db_path)
        cache.close()

    def test_creates_table(self, tmp_path):
        cache = _make_cache(tmp_path)
        with cache.lock:
            cur = cache.conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='emails'"
            )
            assert cur.fetchone() is not None
        cache.close()


class TestMailCacheOperations:
    def test_store_and_load_list(self, tmp_path):
        cache = _make_cache(tmp_path)
        messages = [
            ("uid1", "\\Seen", "alice@test.com", "Hello", "2026-02-20", None),
            ("uid2", "", "bob@test.com", "World", "2026-02-20", None),
        ]
        cache.store_batch("acct1", "INBOX", messages)
        result = cache.load_list("acct1", "INBOX")
        assert len(result) == 2
        # Check fields (uid, flags, from_addr, subject, date_str)
        uids = {r[0] for r in result}
        assert "uid1" in uids
        assert "uid2" in uids
        cache.close()

    def test_load_raw(self, tmp_path):
        cache = _make_cache(tmp_path)
        raw_bytes = b"From: test@test.com\r\nSubject: Test\r\n\r\nBody"
        messages = [("uid1", "", "test@test.com", "Test", "2026-02-20", raw_bytes)]
        cache.store_batch("acct1", "INBOX", messages)
        loaded = cache.load_raw("acct1", "INBOX", "uid1")
        assert loaded == raw_bytes
        cache.close()

    def test_load_raw_missing(self, tmp_path):
        cache = _make_cache(tmp_path)
        result = cache.load_raw("acct1", "INBOX", "nonexistent")
        assert result is None
        cache.close()

    def test_get_uids(self, tmp_path):
        cache = _make_cache(tmp_path)
        messages = [
            ("uid1", "", "a@b.com", "S1", "2026-01-01", None),
            ("uid2", "", "c@d.com", "S2", "2026-01-02", None),
        ]
        cache.store_batch("acct1", "INBOX", messages)
        uids = cache.get_uids("acct1", "INBOX")
        assert uids == {"uid1", "uid2"}
        cache.close()

    def test_delete(self, tmp_path):
        cache = _make_cache(tmp_path)
        messages = [("uid1", "", "a@b.com", "S1", "2026-01-01", None)]
        cache.store_batch("acct1", "INBOX", messages)
        cache.delete("acct1", "INBOX", "uid1")
        uids = cache.get_uids("acct1", "INBOX")
        assert "uid1" not in uids
        cache.close()

    def test_store_batch_ignore_duplicate(self, tmp_path):
        cache = _make_cache(tmp_path)
        messages = [("uid1", "", "a@b.com", "S1", "2026-01-01", None)]
        cache.store_batch("acct1", "INBOX", messages)
        # Store same uid again — should not raise
        cache.store_batch("acct1", "INBOX", messages)
        uids = cache.get_uids("acct1", "INBOX")
        assert len(uids) == 1
        cache.close()

    def test_separate_accounts(self, tmp_path):
        cache = _make_cache(tmp_path)
        cache.store_batch("acct1", "INBOX", [("uid1", "", "a@b.com", "S1", "d1", None)])
        cache.store_batch("acct2", "INBOX", [("uid1", "", "c@d.com", "S2", "d2", None)])
        uids1 = cache.get_uids("acct1", "INBOX")
        uids2 = cache.get_uids("acct2", "INBOX")
        assert uids1 == {"uid1"}
        assert uids2 == {"uid1"}
        list1 = cache.load_list("acct1", "INBOX")
        list2 = cache.load_list("acct2", "INBOX")
        assert list1[0][3] == "S1"  # subject is at index 3
        assert list2[0][3] == "S2"
        cache.close()

    def test_separate_folders(self, tmp_path):
        cache = _make_cache(tmp_path)
        cache.store_batch("acct1", "INBOX", [("uid1", "", "a@b.com", "Inbox", "d1", None)])
        cache.store_batch("acct1", "Sent", [("uid1", "", "a@b.com", "Sent", "d1", None)])
        inbox_list = cache.load_list("acct1", "INBOX")
        sent_list = cache.load_list("acct1", "Sent")
        assert inbox_list[0][3] == "Inbox"
        assert sent_list[0][3] == "Sent"
        cache.close()
