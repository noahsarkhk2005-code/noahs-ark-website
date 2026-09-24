#!/usr/bin/env python3
"""Local-only mock for Noah's Ark member API (IT狗 contract + google_auth)."""
from http.server import BaseHTTPRequestHandler, HTTPServer
import json, uuid

HOST, PORT = '127.0.0.1', 8799

ACCOUNTS = {
    'NA000042': {
        'password': None, 'must_change': False, 'approved': True, 'locked': False,
        'profile': {'member_no': 'NA000042', 'name': '阿甲', 'phone': '91234567', 'email': 'a@b.co'},
        'google_sub': None,
    },
    'NA000099': {
        'password': 'ReadyPass9', 'must_change': False, 'approved': True, 'locked': False,
        'profile': {'member_no': 'NA000099', 'name': '阿乙', 'phone': '98765432', 'email': 'b@c.co'},
        'google_sub': 'linked-sub',
    },
    'NA000088': {
        'password': 'ForceChg1', 'must_change': True, 'approved': True, 'locked': False,
        'profile': {'member_no': 'NA000088', 'name': '阿丙', 'phone': '91112222', 'email': 'c@d.co'},
        'google_sub': None,
    },
    'NA000001': {
        'password': None, 'must_change': False, 'approved': False, 'locked': False,
        'profile': {'member_no': 'NA000001', 'name': '待核', 'phone': '90000001', 'email': 'p@c.co'},
        'google_sub': 'pending-sub',
    },
    'NA000066': {
        'password': 'LockedPw1', 'must_change': False, 'approved': True, 'locked': True,
        'profile': {'member_no': 'NA000066', 'name': '鎖住', 'phone': '90000066', 'email': 'l@c.co'},
        'google_sub': 'locked-sub',
    },
}
# google_sub -> member_no for linked accounts
GOOGLE_LINKS = {a['google_sub']: no for no, a in ACCOUNTS.items() if a.get('google_sub')}

SETUP_TOKENS = {'setup-valid-042': 'NA000042', 'setup-expired': None}
RESET_TOKENS = {'reset-valid-099': 'NA000099', 'reset-expired': None}
TOKENS = {}
CORS = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
}

def parse_mock_token(id_token):
    """mock:<sub>:<name>:<email>  or bare invalid."""
    if not id_token or not isinstance(id_token, str):
        return None
    if not id_token.startswith('mock:'):
        return None
    parts = id_token.split(':', 3)
    if len(parts) < 2:
        return {'sub': 'bad', 'name': '', 'email': ''}
    sub = parts[1]
    name = parts[2] if len(parts) > 2 else ''
    email = parts[3] if len(parts) > 3 else ''
    return {'sub': sub, 'name': name, 'email': email}

class H(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print('[mock]', fmt % args)

    def _send(self, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        for k, v in CORS.items():
            self.send_header(k, v)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        for k, v in CORS.items():
            self.send_header(k, v)
        self.end_headers()

    def do_POST(self):
        n = int(self.headers.get('Content-Length') or 0)
        raw = self.rfile.read(n).decode('utf-8') if n else '{}'
        try:
            data = json.loads(raw)
        except Exception:
            return self._send({'status': 'invalid'})
        action = data.get('action')
        print('[mock] action=', action, data)

        if action == 'google_auth':
            parsed = parse_mock_token(data.get('id_token'))
            if not parsed or parsed['sub'] in ('bad', ''):
                return self._send({'status': 'invalid'})
            sub = parsed['sub']
            if sub == 'locked-sub' or sub in GOOGLE_LINKS and ACCOUNTS[GOOGLE_LINKS[sub]].get('locked'):
                return self._send({'status': 'locked'})
            if sub == 'pending-sub' or (sub in GOOGLE_LINKS and not ACCOUNTS[GOOGLE_LINKS[sub]].get('approved')):
                return self._send({'status': 'pending'})
            if sub in GOOGLE_LINKS:
                no = GOOGLE_LINKS[sub]
                acc = ACCOUNTS[no]
                tok = str(uuid.uuid4())
                TOKENS[tok] = no
                return self._send({
                    'status': 'ok',
                    'token': tok,
                    'must_change_password': bool(acc.get('must_change')),
                    'profile': dict(acc['profile']),
                })
            # unknown sub → need_register
            return self._send({
                'status': 'need_register',
                'name': parsed.get('name') or 'Google User',
                'email': parsed.get('email') or 'user@gmail.com',
            })

        if action == 'register':
            # With id_token: Google registration (name/email from token)
            if data.get('id_token'):
                parsed = parse_mock_token(data.get('id_token'))
                if not parsed:
                    return self._send({'status': 'invalid'})
                # Link on pending for demo — just accept
                return self._send({'status': 'pending'})
            return self._send({'status': 'pending'})

        if action == 'login':
            no = str(data.get('member_no') or '').upper().replace(' ', '')
            pw = str(data.get('password') or '')
            acc = ACCOUNTS.get(no)
            if not acc:
                return self._send({'status': 'invalid'})
            if acc.get('locked'):
                return self._send({'status': 'locked'})
            if (not acc.get('approved')) or (acc.get('password') is None):
                return self._send({'status': 'pending'})
            if acc['password'] != pw:
                return self._send({'status': 'invalid'})
            tok = str(uuid.uuid4())
            TOKENS[tok] = no
            return self._send({
                'status': 'ok',
                'token': tok,
                'must_change_password': bool(acc.get('must_change')),
                'profile': dict(acc['profile']),
            })

        if action == 'setup_password':
            st = data.get('setup_token')
            if st not in SETUP_TOKENS or SETUP_TOKENS[st] is None:
                return self._send({'status': 'expired'})
            no = SETUP_TOKENS[st]
            new_pw = str(data.get('new_password') or '')
            if len(new_pw) < 8:
                return self._send({'status': 'invalid'})
            ACCOUNTS[no]['password'] = new_pw
            ACCOUNTS[no]['must_change'] = False
            del SETUP_TOKENS[st]
            return self._send({'status': 'ok'})

        if action == 'change_password':
            tok = data.get('token')
            no = TOKENS.get(tok)
            if not no:
                return self._send({'status': 'expired'})
            acc = ACCOUNTS[no]
            if acc['password'] != data.get('old_password'):
                return self._send({'status': 'invalid'})
            new_pw = str(data.get('new_password') or '')
            if len(new_pw) < 8:
                return self._send({'status': 'invalid'})
            acc['password'] = new_pw
            acc['must_change'] = False
            return self._send({'status': 'ok'})

        if action == 'reset_request':
            no = str(data.get('member_no') or '').upper().replace(' ', '')
            if no in ACCOUNTS and ACCOUNTS[no].get('password'):
                rt = 'reset-valid-' + no[-3:]
                RESET_TOKENS[rt] = no
                print('[mock] issued reset_token', rt)
            return self._send({'status': 'pending'})

        if action == 'reset_password':
            rt = data.get('reset_token')
            if rt not in RESET_TOKENS or RESET_TOKENS[rt] is None:
                return self._send({'status': 'expired'})
            no = RESET_TOKENS[rt]
            new_pw = str(data.get('new_password') or '')
            if len(new_pw) < 8:
                return self._send({'status': 'invalid'})
            ACCOUNTS[no]['password'] = new_pw
            ACCOUNTS[no]['must_change'] = False
            del RESET_TOKENS[rt]
            return self._send({'status': 'ok'})

        if action == 'profile':
            tok = data.get('token')
            no = TOKENS.get(tok)
            if not no:
                return self._send({'status': 'expired'})
            return self._send({'status': 'ok', 'profile': dict(ACCOUNTS[no]['profile'])})

        if action == 'logout':
            TOKENS.pop(data.get('token'), None)
            return self._send({'status': 'ok'})

        return self._send({'status': 'invalid'})

if __name__ == '__main__':
    print(f'Member mock http://{HOST}:{PORT}/')
    print('google: mock:linked-sub:… → ok; mock:new-sub:… → need_register')
    print('login: NA000099 / ReadyPass9')
    HTTPServer((HOST, PORT), H).serve_forever()
