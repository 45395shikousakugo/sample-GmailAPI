import os.path
import sys
import base64
import re
import datetime
import pytz

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# 使用するGmailAPI
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# メール取得条件（meは認証情報のGmail）
USERID = "me"
# 検索条件
# 「in:anywhere」は迷惑メール、ゴミ箱を含む（全メール）
# Gmailの検索条件「https://support.google.com/mail/answer/7190?hl=ja」は参照
Q = "in:anywhere 「gmailの検索条件を記載」"
# 取得件数
NUM = 500


# GmailAPIトークン取得
def getGmailToken():
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  return creds


# GmailAPI サービス起動
def buildGmailService(creds):

  service = build("gmail", "v1", credentials=creds)
  return service


# GmailAPI メッセージリスト取得
# Method: users.messages.list の仕様
# https://developers.google.com/gmail/api/reference/rest/v1/users.messages/list?hl=ja
def getGmailMsgList(service,USERID,Q,NUM):
  try:
    # GmailAPIメール一覧取得
    results = service.users().messages().list(userId=USERID,q=Q,maxResults=NUM).execute()
    messages = results.get("messages", [])

  except HttpError as error:
    print(f"An error occurred: {error}")
    sys.exit()

  return messages


# GmailAPI メッセージ詳細取得
# Method: users.messages.get の仕様
# https://developers.google.com/gmail/api/reference/rest/v1/users.messages/get?hl=ja
def getGmailMsgDetail(service,msgId):
  try:
    detail = service.users().messages().get(userId=USERID,id=msgId).execute()
  except HttpError as error:
    print(f"An error occurred: {error}")
    sys.exit()

  return detail


# GmailAPI 件名取得
def getGmailSubject(headers):
  subject = ''
  for h in headers:
    if h['name'] == 'Subject':
      subject = h['value']
  return subject


# GmailAPI メールアドレス取得
def getGmailAddr(headers,label):
  subject = ''
  for h in headers:
    if h['name'] == label:
      mail = getMailAddr(h['value'])
  return mail


# GmailAPI 送受信日時（日本時間に変更）
def getGmailInternalDate(unixtime):
  utc_datetime = datetime.datetime.fromtimestamp(int(unixtime)/1000, datetime.timezone.utc)
  jst_timezone = pytz.timezone('Asia/Tokyo')
  internaldate = utc_datetime.replace(tzinfo=pytz.utc).astimezone(jst_timezone)
  return internaldate


# GmailAPI 本文取得
def getGamilBody(txt):
  message = ''
  # textメール
  if 'data' in txt['payload']['body']:
    message = txt['payload']['body']['data']
  # htmlメール
  elif 'parts' in txt['payload']:
    if 'parts' in txt['payload']['parts'][0]:
      message = txt['payload']['parts'][0]['parts'][0]['body']['data']
    elif 'body' in txt['payload']['parts'][0]:
      message = txt['payload']['parts'][0]['body']['data']

  return base64ToUtf8(message)


# メールアドレス抽出
def getMailAddr(str):
  mail = ''
  pattern = r'[\w\.-]+@[\w\.-]+'
  match = re.search(pattern, str)
  if match:
    # 抽出結果があればメールアドレスを出力
    mail = match.group()
  return mail


# base64をUTF-8に変換
def base64ToUtf8(str):
  return base64.urlsafe_b64decode(str).decode()


def main():

  # gmailトークン発行
  creds = getGmailToken()
  # service起動
  service = buildGmailService(creds)
  # 対象メールのメッセージ詳細を取得
  messages = getGmailMsgList(service,USERID,Q,NUM)

  # print(messages)

  # メール詳細情報取得
  l = []
  d = {}
  for message in messages:
    # メール詳細取得
    detail = getGmailMsgDetail(service,message["id"])
    #メールID
    d['mailId'] = message['id']
    # スレッドID
    d['threadId'] = message["threadId"]
    # 件名取得
    d['subject'] = getGmailSubject(detail['payload']['headers'])
    # From取得
    d['fromMail'] = getGmailAddr(detail['payload']['headers'],"From")
    # To取得
    d['toMail'] = getGmailAddr(detail['payload']['headers'],"To")
    # 送受信日時取得
    d['unixTime'] = detail['internalDate']
    d['internalDate'] = getGmailInternalDate(detail['internalDate'])
    # 本文取得
    d['body'] = getGamilBody(detail)

    l.append(d)

  print(l)


if __name__ == "__main__":
  main()