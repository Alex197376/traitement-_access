import os
import base64
import mimetypes
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QProgressBar
from PyQt5.QtCore import Qt, QTimer

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
TOKEN_PATH = "token.json"
CREDENTIALS_PATH = "credentials.json"
ATTACHMENTS_DIR = "gmail_ddt_pieces_jointes"

def authentifier_gmail():
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, 'w') as token:
            token.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

class ProgressDialog(QDialog):
    def __init__(self, total_steps, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Scan en cours")
        self.setWindowModality(Qt.ApplicationModal)
        self.setMinimumWidth(300)

        layout = QVBoxLayout(self)
        self.label = QLabel("Analyse des messages Gmail en cours…")
        self.progress = QProgressBar()
        self.progress.setMaximum(total_steps)
        self.progress.setValue(0)

        layout.addWidget(self.label)
        layout.addWidget(self.progress)

        self.setLayout(layout)

    def update_progress(self, value):
        self.progress.setValue(value)


def telecharger_pieces_jointes(parent=None):
    service = authentifier_gmail()
    results = service.users().messages().list(userId='me', labelIds=["SENT"], q="has:attachment filename:pdf").execute()
    messages = results.get("messages", [])

    historique_path = "historique_scan.json"
    if os.path.exists(historique_path):
        with open(historique_path, "r", encoding="utf-8") as f:
            historique = json.load(f)
    else:
        historique = {"messages": [], "fichiers": []}

    fichiers_trouves = []
    total = len(messages)
    dialog = ProgressDialog(total_steps=total, parent=parent)
    dialog.show()

    for i, msg in enumerate(messages):
        QApplication.processEvents()
        if msg["id"] in historique["messages"]:
            dialog.update_progress(i + 1)
            continue

        message = service.users().messages().get(userId='me', id=msg['id']).execute()
        parts = message['payload'].get('parts', [])
        for part in parts:
            filename = part.get("filename")
            if filename and filename.lower().endswith(".pdf"):
                if filename not in historique["fichiers"]:
                    fichiers_trouves.append(filename)
                    historique["fichiers"].append(filename)

        historique["messages"].append(msg["id"])
        dialog.update_progress(i + 1)

    with open(historique_path, "w", encoding="utf-8") as f:
        json.dump(historique, f, indent=2, ensure_ascii=False)

    QTimer.singleShot(300, dialog.close)
    return fichiers_trouves
