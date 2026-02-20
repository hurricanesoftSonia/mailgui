"""Tests for MsgTool HTTP Client."""
import json
import pytest
from unittest.mock import patch, MagicMock
from msgtool_client import MsgClient


class TestMsgClientInit:
    def test_default_server(self):
        with patch.dict("os.environ", {}, clear=True):
            client = MsgClient()
            assert client.server_url == "http://localhost:8900"

    def test_custom_server(self):
        client = MsgClient(server_url="http://custom:9000")
        assert client.server_url == "http://custom:9000"

    def test_trailing_slash_stripped(self):
        client = MsgClient(server_url="http://custom:9000/")
        assert client.server_url == "http://custom:9000"

    def test_env_server(self):
        with patch.dict("os.environ", {"MSGTOOL_SERVER": "http://env:8800"}):
            client = MsgClient()
            assert client.server_url == "http://env:8800"

    def test_credentials(self):
        client = MsgClient(user="luna", password="secret")
        assert client.user == "luna"
        assert client.password == "secret"

    def test_env_credentials(self):
        with patch.dict("os.environ", {"MSG_USER": "envuser", "MSG_PASSWORD": "envpass"}):
            client = MsgClient()
            assert client.user == "envuser"
            assert client.password == "envpass"


class TestMsgClientHeaders:
    def test_headers_contain_credentials(self):
        client = MsgClient(user="luna", password="secret")
        headers = client._headers()
        assert headers["X-User"] == "luna"
        assert headers["X-Password"] == "secret"
        assert headers["Content-Type"] == "application/json"


class TestMsgClientGet:
    @patch("msgtool_client.urllib.request.urlopen")
    def test_get_success(self, mock_urlopen):
        resp_data = {"status": "ok"}
        mock_resp = MagicMock()
        mock_resp.read.return_value = json.dumps(resp_data).encode()
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_urlopen.return_value = mock_resp

        client = MsgClient(server_url="http://test:8900")
        result = client._get("/health")
        assert result == resp_data

    @patch("msgtool_client.urllib.request.urlopen")
    def test_get_with_params(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.read.return_value = b'{"items": []}'
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_urlopen.return_value = mock_resp

        client = MsgClient(server_url="http://test:8900")
        client._get("/inbox", {"limit": "10", "unread": "1"})
        call_args = mock_urlopen.call_args
        url = call_args[0][0].full_url
        assert "limit=10" in url
        assert "unread=1" in url

    @patch("msgtool_client.urllib.request.urlopen")
    def test_get_http_error_json(self, mock_urlopen):
        import urllib.error
        error_body = json.dumps({"error": "not found"}).encode()
        mock_fp = MagicMock()
        mock_fp.read.return_value = error_body
        error = urllib.error.HTTPError("http://test", 404, "Not Found", {}, mock_fp)
        error.fp = mock_fp
        mock_urlopen.side_effect = error

        client = MsgClient(server_url="http://test:8900")
        result = client._get("/missing")
        assert result == {"error": "not found"}

    @patch("msgtool_client.urllib.request.urlopen")
    def test_get_connection_error(self, mock_urlopen):
        mock_urlopen.side_effect = ConnectionError("refused")

        client = MsgClient(server_url="http://test:8900")
        result = client._get("/health")
        assert "error" in result


class TestMsgClientPost:
    @patch("msgtool_client.urllib.request.urlopen")
    def test_post_success(self, mock_urlopen):
        resp_data = {"ok": True}
        mock_resp = MagicMock()
        mock_resp.read.return_value = json.dumps(resp_data).encode()
        mock_resp.__enter__ = lambda s: s
        mock_resp.__exit__ = MagicMock(return_value=False)
        mock_urlopen.return_value = mock_resp

        client = MsgClient(server_url="http://test:8900")
        result = client._post("/send", {"to": "bob", "msg": "hi"})
        assert result == {"ok": True}


class TestMsgClientMethods:
    def setup_method(self):
        self.client = MsgClient(server_url="http://test:8900", user="luna", password="pw")

    @patch.object(MsgClient, "_get")
    def test_health(self, mock_get):
        mock_get.return_value = {"status": "ok"}
        result = self.client.health()
        mock_get.assert_called_once_with("/health")
        assert result == {"status": "ok"}

    @patch.object(MsgClient, "_get")
    def test_inbox(self, mock_get):
        mock_get.return_value = {"messages": []}
        self.client.inbox(unread=True, limit=10)
        mock_get.assert_called_once_with("/inbox", {"limit": "10", "unread": "1"})

    @patch.object(MsgClient, "_get")
    def test_sent(self, mock_get):
        mock_get.return_value = {"messages": []}
        self.client.sent(limit=5)
        mock_get.assert_called_once_with("/sent", {"limit": "5"})

    @patch.object(MsgClient, "_get")
    def test_read(self, mock_get):
        mock_get.return_value = {"body": "hello"}
        self.client.read("msg123")
        mock_get.assert_called_once_with("/read/msg123")

    @patch.object(MsgClient, "_post")
    def test_send(self, mock_post):
        mock_post.return_value = {"ok": True}
        self.client.send("bob", "hi")
        mock_post.assert_called_once_with("/send", {"to": "bob", "msg": "hi"})

    @patch.object(MsgClient, "_post")
    def test_send_with_reply(self, mock_post):
        mock_post.return_value = {"ok": True}
        self.client.send("bob", "hi", reply_to="msg1")
        mock_post.assert_called_once_with("/send", {"to": "bob", "msg": "hi", "reply_to": "msg1"})

    @patch.object(MsgClient, "_post")
    def test_reply(self, mock_post):
        mock_post.return_value = {"ok": True}
        self.client.reply("msg1", "thanks")
        mock_post.assert_called_once_with("/reply", {"id": "msg1", "msg": "thanks"})

    @patch.object(MsgClient, "_post")
    def test_register(self, mock_post):
        mock_post.return_value = {"ok": True}
        self.client.register("newuser", "pass123", "New User")
        mock_post.assert_called_once_with("/register", {
            "username": "newuser", "password": "pass123", "display_name": "New User"
        })
