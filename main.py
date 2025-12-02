import os
import json
import datetime
import gspread
from slack_sdk import WebClient
# Google Vertex AI (Gemini)
import vertexai
from vertexai.generative_models import GenerativeModel
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# --- 設定と認証 ---
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
GOOGLE_SHEET_ID = os.environ["GOOGLE_SHEET_ID"]
DRIVE_FOLDER_ID = os.environ["GOOGLE_DRIVE_FOLDER_ID"]
# SecretsからJSONキーを読み込む
GOOGLE_JSON_KEY = json.loads(os.environ["GOOGLE_JSON_KEY"])
# JSONキーからプロジェクトIDを自動取得
PROJECT_ID = GOOGLE_JSON_KEY["project_id"]

# Slackクライアント
slack = WebClient(token=SLACK_BOT_TOKEN)

# Google認証 (スプレッドシート, Docs, Drive, Gemini全て共通)
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/cloud-platform'
]
creds = Credentials.from_service_account_info(GOOGLE_JSON_KEY, scopes=SCOPES)

# Geminiの初期化
vertexai.init(project=PROJECT_ID, location="us-central1", credentials=creds)
model = GenerativeModel("gemini-1.5-flash") # 高速・低コストなモデル

# Googleサービス初期化
gc_sheets = gspread.authorize(creds)
docs_service = build('docs', 'v1', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

def get_configs():
    """スプレッドシートから設定(Activeなもの)を取得"""
    print("Reading config from Google Sheet...")
    worksheet = gc_sheets.open_by_key(GOOGLE_SHEET_ID).sheet1
    rows = worksheet.get_all_records()
    return [row for row in rows if row.get('Status') == 'Active']

def get_yesterday_messages(channel_id):
    """Slackから昨日のメッセージを取得"""
    today = datetime.datetime.now()
    yesterday = today - datetime.timedelta(days=1)
    
    oldest = yesterday.replace(hour=0, minute=0, second=0).timestamp()
    latest = yesterday.replace(hour=23, minute=59, second=59).timestamp()

    try:
        result = slack.conversations_history(
            channel=channel_id, oldest=str(oldest), latest=str(latest)
        )
        messages = result["messages"]
        if not messages:
            return None

        text_data = []
        for m in messages[::-1]:
            user = m.get('user', 'Unknown')
            text = m.get('text', '')
            if text:
                text_data.append(f"User({user}): {text}")
        return "\n".join(text_data)
    except Exception as e:
        print(f"Error fetching Slack: {e}")
        return None

def summarize_text(text, project_name):
    """Geminiで要約を作成"""
    if not text:
        return "昨日のメッセージはありませんでした。"

    prompt = f"""
    あなたは優秀なプロジェクトマネージャーです。
    以下のSlackログ（プロジェクト: {project_name}）から、昨日の動きを把握できる議事録サマリを作成してください。
    
    【要件】
    1. 「決定事項」と「その経緯（誰がなぜそうしたか）」を明確に。
    2. ネクストアクションや課題があれば記載。
    3. 雑談は除外。
    4. Googleドキュメント用のため、Markdownは使わず、見出しは【見出し】のように隅付き括弧で表現。
    
    【ログ】
    {text[:20000]}
    """
    
    # Gemini生成実行
    response = model.generate_content(prompt)
    return response.text

def create_google_doc(project_name, summary):
    """Googleドキュメント作成・保存"""
    yesterday_str = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    doc_title = f"{yesterday_str}_{project_name}_日報"

    # ドキュメント作成
    doc = docs_service.documents().create(body={'title': doc_title}).execute()
    doc_id = doc['documentId']
    print(f"Created Doc: {doc_title}")

    # 書き込み
    requests = [{'insertText': {'location': {'index': 1}, 'text': summary}}]
    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

    # フォルダ移動
    file = drive_service.files().get(fileId=doc_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    drive_service.files().update(
        fileId=doc_id, addParents=DRIVE_FOLDER_ID, removeParents=previous_parents
    ).execute()
    print(f"Moved to folder.")

def main():
    configs = get_configs()
    print(f"Found {len(configs)} active projects.")

    for config in configs:
        project_name = config['プロジェクト名']
        channel_id = config['SlackチャンネルID']
        print(f"--- Processing: {project_name} ---")
        
        text = get_yesterday_messages(channel_id)
        if text:
            print("Summarizing with Gemini...")
            summary = summarize_text(text, project_name)
            create_google_doc(project_name, summary)
        else:
            print("No messages.")

if __name__ == "__main__":
    main()
