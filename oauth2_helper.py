import os
import json
import webbrowser
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import requests
import time

class OAuth2TokenStorage:
    def __init__(self, token_file='token.json'):
        self.token_file = token_file
        self.token_data = None
        self.load_token()

    def load_token(self):
        if os.path.exists(self.token_file):
            with open(self.token_file, 'r') as f:
                self.token_data = json.load(f)
        else:
            self.token_data = None

    def save_token(self, token_data):
        self.token_data = token_data
        with open(self.token_file, 'w') as f:
            json.dump(token_data, f)

    def get_access_token(self):
        if not self.token_data:
            return None
        expires_at = self.token_data.get('expires_at', 0)
        if time.time() > expires_at:
            return None
        return self.token_data.get('access_token')

    def get_refresh_token(self):
        if not self.token_data:
            return None
        return self.token_data.get('refresh_token')

class OAuth2Client:
    def __init__(self, client_id, client_secret, auth_uri, token_uri, redirect_uri, scope):
        self.client_id = client_id
        self.client_secret = client_secret
        self.auth_uri = auth_uri
        self.token_uri = token_uri
        self.redirect_uri = redirect_uri
        self.scope = scope
        self.token_storage = OAuth2TokenStorage()

    def get_authorization_url(self, state):
        params = {
            'client_id': self.client_id,
            'response_type': 'code',
            'redirect_uri': self.redirect_uri,
            'scope': self.scope,
            'state': state,
            'access_type': 'offline',
            'prompt': 'consent'
        }
        from urllib.parse import urlencode
        return f"{self.auth_uri}?{urlencode(params)}"

    def exchange_code_for_token(self, code):
        data = {
            'code': code,
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'redirect_uri': self.redirect_uri,
            'grant_type': 'authorization_code'
        }
        response = requests.post(self.token_uri, data=data)
        response.raise_for_status()
        token_data = response.json()
        # Calculate expires_at timestamp
        token_data['expires_at'] = time.time() + token_data.get('expires_in', 3600)
        self.token_storage.save_token(token_data)
        return token_data

    def refresh_access_token(self):
        refresh_token = self.token_storage.get_refresh_token()
        if not refresh_token:
            return None
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token'
        }
        response = requests.post(self.token_uri, data=data)
        response.raise_for_status()
        token_data = response.json()
        token_data['refresh_token'] = refresh_token  # keep the same refresh token
        token_data['expires_at'] = time.time() + token_data.get('expires_in', 3600)
        self.token_storage.save_token(token_data)
        return token_data

    def get_access_token(self):
        access_token = self.token_storage.get_access_token()
        if access_token:
            return access_token
        # Try to refresh
        token_data = self.refresh_access_token()
        if token_data:
            return token_data.get('access_token')
        return None

class OAuth2CallbackHandler(BaseHTTPRequestHandler):
    server_version = "OAuth2CallbackHandler/0.1"

    def do_GET(self):
        parsed_path = urlparse(self.path)
        query = parse_qs(parsed_path.query)
        if 'code' in query:
            self.server.auth_code = query['code'][0]
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b"<html><body><h1>Authentication successful. You can close this window.</h1></body></html>")
        else:
            self.send_response(400)
            self.end_headers()

def run_local_server(port=8080):
    server_address = ('', port)
    httpd = HTTPServer(server_address, OAuth2CallbackHandler)
    httpd.auth_code = None
    thread = threading.Thread(target=httpd.serve_forever)
    thread.daemon = True
    thread.start()
    return httpd

def authenticate(oauth_client):
    state = 'state123'  # In production, generate a random state
    auth_url = oauth_client.get_authorization_url(state)
    httpd = run_local_server()
    webbrowser.open(auth_url)
    print("Please complete the authentication in the browser.")
    while httpd.auth_code is None:
        time.sleep(1)
    code = httpd.auth_code
    httpd.shutdown()
    token_data = oauth_client.exchange_code_for_token(code)
    return token_data
