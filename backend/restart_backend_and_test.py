import json, os, subprocess, time, urllib.request
from urllib.error import HTTPError, URLError

BASE='http://localhost:8082'
BACKEND = r'E:\tharu\claudefree\ai-excel-report-platform\backend'
MVN = r'E:\tharu\claudefree\tools\apache-maven-3.9.9\bin\mvn.cmd'



def start_backend():
    env = os.environ.copy()
    env['GROQ_API_KEY'] = KEY
    env['GROQ_BASE_URL'] = 'https://api.groq.com/openai/v1'
    env['GROQ_MODEL'] = 'openai/gpt-oss-20b'
    env['SPRING_AI_OPENAI_API_KEY'] = KEY
    env['SPRING_AI_OPENAI_BASE_URL'] = 'https://api.groq.com/openai/v1'
    env['SPRING_AI_OPENAI_CHAT_OPTIONS_MODEL'] = 'openai/gpt-oss-20b'
    env['SPRING_AI_OPENAI_CHAT_OPTIONS_MAX_TOKENS'] = '4000'
    env['SERVER_PORT'] = '8082'

    try:
        out = subprocess.check_output([
            'powershell', '-NoProfile', '-Command',
            '(Get-NetTCPConnection -LocalPort 8082 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -ErrorAction SilentlyContinue) 2>$null'
        ], text=True)
        for pid in sorted({p.strip() for p in out.split() if p.strip()}):
            try:
                subprocess.run(['cmd', '/c', f'taskkill /PID {pid} /F'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                print('KILLED', pid)
            except Exception as e:
                print('KILL_FAILED', pid, e)
    except Exception as e:
        print('PORT_CHECK_ERROR', e)

    log_path = os.path.join(BACKEND, 'spring-boot-startup.log')
    with open(log_path, 'wb') as log_file:
        proc = subprocess.Popen(['cmd', '/c', f'"{MVN}" spring-boot:run'], cwd=BACKEND, env=env, stdout=log_file, stderr=subprocess.STDOUT)
        print('STARTED_PID', proc.pid)
        deadline = time.time() + 120
        while time.time() < deadline:
            if os.path.exists(log_path):
                with open(log_path, 'r', encoding='utf-8', errors='replace') as fh:
                    content = fh.read()
                if 'Tomcat started on port 8082' in content or 'Started AiExcelReportPlatformApplication' in content:
                    print('SERVER_STARTED')
                    return proc
            time.sleep(1)
        print('SERVER_START_TIMEOUT')
        return proc


def post_json(path, data):
    req = urllib.request.Request(BASE + path, data=json.dumps(data).encode('utf-8'), headers={'Content-Type': 'application/json'}, method='POST')
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode('utf-8'))


def get_json(path, token):
    req = urllib.request.Request(BASE + path, headers={'Authorization': f'Bearer {token}'}, method='GET')
    with urllib.request.urlopen(req, timeout=120) as resp:
        content = resp.read().decode('utf-8')
        return json.loads(content)


def upload_sample(token):
    sample = os.path.join(BACKEND, 'uploads', 'fff6b23d-bcf5-4e59-bb7f-59f288cd6108.xlsx')
    if not os.path.exists(sample):
        raise FileNotFoundError(sample)
    boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
    with open(sample, 'rb') as f:
        file_bytes = f.read()
    body = []
    body.append(('--' + boundary).encode('utf-8'))
    body.append(f'Content-Disposition: form-data; name="file"; filename="{os.path.basename(sample)}"'.encode('utf-8'))
    body.append(b'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    body.append(b'')
    body.append(file_bytes)
    body.append(('--' + boundary + '--').encode('utf-8'))
    payload = b'\r\n'.join(body)
    headers = {'Content-Type': f'multipart/form-data; boundary={boundary}', 'Authorization': f'Bearer {token}'}
    req = urllib.request.Request(BASE + '/api/files', data=payload, headers=headers, method='POST')
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode('utf-8'))


if __name__ == '__main__':
    print('RESTARTING_BACKEND...')
    proc = start_backend()
    print('WAITING_FOR_APP...')
    time.sleep(5)

    print('REGISTERING_USER...')
    try:
        reg = post_json('/api/auth/register', {
            'name': 'AITest',
            'email': 'aitest+1@example.com',
            'password': 'Password1!'
        })
        print('REGISTER_RESULT', reg)
    except HTTPError as e:
        print('REGISTER_HTTP', e.code, e.read().decode('utf-8', 'replace'))
    except Exception as e:
        print('REGISTER_EXCEPTION', type(e).__name__, e)

    print('LOGGING_IN...')
    try:
        login = post_json('/api/auth/login', {
            'email': 'aitest+1@example.com',
            'password': 'Password1!'
        })
        token = login.get('token')
        print('LOGIN_RESULT', {'userId': login.get('userId'), 'name': login.get('name'), 'email': login.get('email')})
    except HTTPError as e:
        print('LOGIN_HTTP', e.code, e.read().decode('utf-8', 'replace'))
        raise SystemExit(1)

    print('UPLOADING_SAMPLE...')
    upload = upload_sample(token)
    print('UPLOAD_RESULT', upload)
    file_id = upload['id']

    print('\n---UNDERSTANDING---')
    try:
        print(get_json(f'/api/files/{file_id}/understanding', token))
    except HTTPError as e:
        print('UNDERSTANDING_HTTP', e.code, e.read().decode('utf-8', 'replace'))

    print('\n---RECOMMENDATIONS---')
    try:
        print(get_json(f'/api/files/{file_id}/recommendations', token))
    except HTTPError as e:
        print('RECOMMENDATIONS_HTTP', e.code, e.read().decode('utf-8', 'replace'))

    print('DONE')
    if proc and proc.poll() is None:
        print('KEEP_RUNNING')
