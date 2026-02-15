"""MsgTool HTTP Client — connects to MsgTool server."""
import json
import urllib.request
import urllib.error
import os


class MsgClient:
    def __init__(self, server_url=None, user=None, password=None):
        self.server_url = (server_url or os.environ.get('MSGTOOL_SERVER', 
                           'http://localhost:8900')).rstrip('/')
        self.user = user or os.environ.get('MSG_USER', '')
        self.password = password or os.environ.get('MSG_PASSWORD', '')

    def _headers(self):
        return {
            'X-User': self.user,
            'X-Password': self.password,
            'Content-Type': 'application/json',
        }

    def _get(self, path, params=None):
        url = f"{self.server_url}{path}"
        if params:
            from urllib.parse import urlencode
            url += '?' + urlencode(params)
        req = urllib.request.Request(url, headers=self._headers())
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                return json.loads(resp.read())
        except urllib.error.HTTPError as e:
            body = e.read().decode() if e.fp else ''
            try:
                return json.loads(body)
            except:
                return {'error': f'HTTP {e.code}: {body}'}
        except Exception as e:
            return {'error': str(e)}

    def _post(self, path, data):
        url = f"{self.server_url}{path}"
        body = json.dumps(data).encode()
        req = urllib.request.Request(url, data=body, headers=self._headers())
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                return json.loads(resp.read())
        except urllib.error.HTTPError as e:
            body = e.read().decode() if e.fp else ''
            try:
                return json.loads(body)
            except:
                return {'error': f'HTTP {e.code}: {body}'}
        except Exception as e:
            return {'error': str(e)}

    def health(self):
        return self._get('/health')

    def inbox(self, unread=False, limit=50):
        params = {'limit': str(limit)}
        if unread:
            params['unread'] = '1'
        return self._get('/inbox', params)

    def sent(self, limit=50):
        return self._get('/sent', {'limit': str(limit)})

    def read(self, msg_id):
        return self._get(f'/read/{msg_id}')

    def send(self, to, msg, reply_to=None):
        data = {'to': to, 'msg': msg}
        if reply_to:
            data['reply_to'] = reply_to
        return self._post('/send', data)

    def reply(self, msg_id, msg):
        return self._post('/reply', {'id': msg_id, 'msg': msg})

    def mentions(self, limit=20):
        return self._get('/mentions', {'limit': str(limit)})

    def notify(self):
        return self._get('/notify')

    def users(self):
        return self._get('/users')

    def register(self, username, password, display_name=''):
        return self._post('/register', {
            'username': username, 'password': password, 'display_name': display_name
        })
