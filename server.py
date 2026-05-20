import os, requests
from flask import Flask, request, jsonify, redirect
from flask_cors import CORS
from msal import ConfidentialClientApplication, SerializableTokenCache
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID', 'consumers')
REDIRECT_URI = os.getenv('REDIRECT_URI', 'http://localhost:8080/callback')
SCOPES = ['Mail.Read', 'Calendars.Read', 'User.Read']
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
CACHE_FILE = 'token_cache.json'

def load_cache():
    cache = SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE).read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        open(CACHE_FILE, 'w').write(cache.serialize())

def get_msal_app(cache=None):
    return ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache
    )

def get_token():
    cache = load_cache()
    msal_app = get_msal_app(cache)
    accounts = msal_app.get_accounts()
    if accounts:
        result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
        save_cache(cache)
        if result and 'access_token' in result:
            return result['access_token']
    return None

@app.route('/')
def index():
    token = get_token()
    if token:
        return jsonify({'status': 'authenticated', 'message': 'MCP server is running'})
    return jsonify({'status': 'unauthenticated', 'message': 'Visit /login to authenticate'})

@app.route('/login')
def login():
    cache = load_cache()
    msal_app = get_msal_app(cache)
    auth_url = msal_app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route('/callback')
def callback():
    code = request.args.get('code')
    if not code:
        return jsonify({'error': 'No code received'}), 400
    cache = load_cache()
    msal_app = get_msal_app(cache)
    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri=REDIRECT_URI)
    save_cache(cache)
    if 'access_token' in result:
        return jsonify({'status': 'success', 'message': 'Authenticated! Token saved.'})
    return jsonify({'error': result.get('error_description', 'Unknown error')}), 400

@app.route('/emails/jobs')
def job_emails():
    token = get_token()
    if not token:
        return jsonify({'error': 'Not authenticated'}), 401
    after = request.args.get('after', '2026-01-01T00:00:00Z')
    keywords = ['thank you for applying', 'thanks for applying', 'your application',
                'application received', 'interview invitation', 'application update',
                'we regret', 'not moving forward', 'not selected', 'position has been filled',
                'thanks for your interest', 'thank you for your interest',
                'submitted your application', 'congratulations your application']
    all_results = []
    seen_ids = set()
    headers = {'Authorization': f'Bearer {token}'}
    for kw in keywords:
        url = f"https://graph.microsoft.com/v1.0/me/messages?$search=\"{kw}\"&$top=50&$select=subject,from,receivedDateTime,bodyPreview,isRead"
        page = 0
        while url and page < 5:
            resp = requests.get(url, headers=headers)
            if resp.status_code != 200:
                break
            data = resp.json()
            for msg in data.get('value', []):
                if msg['id'] not in seen_ids:
                    if msg.get('receivedDateTime', '') >= after:
                        seen_ids.add(msg['id'])
                        all_results.append(msg)
            url = data.get('@odata.nextLink')
            page += 1
    all_results.sort(key=lambda x: x.get('receivedDateTime', ''), reverse=True)
    return jsonify({'value': all_results, 'count': len(all_results)})

@app.route('/emails/rejections')
def rejection_emails():
    token = get_token()
    if not token:
        return jsonify({'error': 'Not authenticated'}), 401
    rejection_phrases = [
        'not moving forward', 'will not be moving forward',
        'after careful consideration, we will not', 'after careful review, we have decided',
        'we have decided not to move forward', 'chosen to pursue other candidates',
        'decided to move forward with other', 'not selected for',
        'position has been filled', 'position has now been filled',
        'unfortunately we will not', 'after thoughtful review',
        'not be proceeding', 'unable to move forward'
    ]
    all_results = []
    seen_ids = set()
    headers = {'Authorization': f'Bearer {token}'}
    for phrase in rejection_phrases:
        url = f"https://graph.microsoft.com/v1.0/me/messages?$search=\"{phrase}\"&$top=50&$select=subject,from,receivedDateTime,bodyPreview,isRead"
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            continue
        data = resp.json()
        for msg in data.get('value', []):
            if msg['id'] not in seen_ids:
                if msg.get('receivedDateTime', '') >= '2026-01-01T00:00:00Z':
                    seen_ids.add(msg['id'])
                    all_results.append(msg)
    all_results.sort(key=lambda x: x.get('receivedDateTime', ''), reverse=True)
    return jsonify({'value': all_results, 'count': len(all_results)})

@app.route('/emails/search')
def search_emails():
    token = get_token()
    if not token:
        return jsonify({'error': 'Not authenticated'}), 401
    query = request.args.get('q', '')
    top = request.args.get('top', 25)
    url = f"https://graph.microsoft.com/v1.0/me/messages?$search=\"{query}\"&$top={top}&$select=subject,from,receivedDateTime,bodyPreview,isRead"
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers)
    return jsonify(response.json())

@app.route('/emails/recent')
def recent_emails():
    token = get_token()
    if not token:
        return jsonify({'error': 'Not authenticated'}), 401
    top = request.args.get('top', 25)
    url = f"https://graph.microsoft.com/v1.0/me/messages?$orderby=receivedDateTime desc&$top={top}&$select=subject,from,receivedDateTime,bodyPreview,isRead"
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers)
    return jsonify(response.json())

@app.route('/emails/read/<message_id>')
def read_email(message_id):
    token = get_token()
    if not token:
        return jsonify({'error': 'Not authenticated'}), 401
    url = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}'
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers)
    return jsonify(response.json())

@app.route('/emails/batch-details')
def batch_details():
    token = get_token()
    if not token:
        return jsonify({'error': 'Not authenticated'}), 401
    ids = request.args.get('ids', '').split(',')
    ids = [i.strip() for i in ids if i.strip()][:20]
    headers = {'Authorization': f'Bearer {token}'}
    results = []
    for msg_id in ids:
        url = f'https://graph.microsoft.com/v1.0/me/messages/{msg_id}?$select=subject,from,receivedDateTime,bodyPreview,body'
        resp = requests.get(url, headers=headers)
        if resp.status_code == 200:
            results.append(resp.json())
    return jsonify({'value': results})

@app.route('/calendar/events')
def calendar_events():
    token = get_token()
    if not token:
        return jsonify({'error': 'Not authenticated'}), 401
    url = "https://graph.microsoft.com/v1.0/me/events?$orderby=start/dateTime&$top=20&$select=subject,start,end,organizer,location"
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers)
    return jsonify(response.json())

if __name__ == '__main__':
    app.run(port=8080, debug=True)
