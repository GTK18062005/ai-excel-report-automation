import urllib.request
import json
from urllib.error import HTTPError, URLError

BASE='http://localhost:8081'

def post(path, data):
    url=BASE+path
    req=urllib.request.Request(url, data=json.dumps(data).encode('utf-8'), headers={'Content-Type':'application/json'}, method='POST')
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode('utf-8'))

def get(path, token):
    url=BASE+path
    headers={'Authorization': f'Bearer {token}'}
    req=urllib.request.Request(url, headers=headers, method='GET')
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode('utf-8'))

if __name__=='__main__':
    try:
        try:
            print('Registering test user...')
            post('/api/auth/register', {"name":"AITest","email":"aitest+1@example.com","password":"Password1!"})
            print('Register OK (or already exists)')
        except HTTPError as e:
            print('Register HTTPError', e.code, e.reason)
        except Exception as e:
            print('Register exception', e)

        print('Logging in...')
        login = post('/api/auth/login', {"email":"aitest+1@example.com","password":"Password1!"})
        token = login.get('token')
        print('TOKEN:', 'REDACTED' if token else None)

        print('\n---UNDERSTANDING---')
        try:
            u = get('/api/files/2/understanding', token)
            print(json.dumps(u, indent=2))
        except HTTPError as e:
            print('UNDERSTANDING HTTPError', e.code)
            print(e.read().decode('utf-8'))
        except Exception as e:
            print('UNDERSTANDING exception', e)

        print('\n---RECOMMENDATIONS---')
        try:
            r = get('/api/files/2/recommendations', token)
            print(json.dumps(r, indent=2))
        except HTTPError as e:
            print('RECOMMENDATIONS HTTPError', e.code)
            try:
                print(e.read().decode('utf-8'))
            except Exception:
                pass
        except Exception as e:
            print('RECOMMENDATIONS exception', e)

    except Exception as outer:
        print('Fatal error', outer)
