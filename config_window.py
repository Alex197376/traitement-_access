from PyQt5.QtWidgets import (
    QWidget, QLabel, QPushButton, QLineEdit, QFileDialog,
    QVBoxLayout, QHBoxLayout, QDialog, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView
)
import json
import os
import re
import unicodedata
import pyodbc
from datetime import datetime

CONFIG_FILE = "config_suiviclientpro.json"
MANUAL_STATE_FILE = "manual_states.json"

class ConfigWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Paramétrage de SuiviClientPro")
        self.setFixedWidth(500)
        self.config = self.load_config()

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.access_label = QLabel("Chemin vers la base LICIEL (.mdb) :")
        self.access_input = QLineEdit(self.config.get("access_path", ""))
        self.access_btn = QPushButton("📂 Parcourir")
        self.access_btn.clicked.connect(self.select_access_path)

        access_layout = QHBoxLayout()
        access_layout.addWidget(self.access_input)
        access_layout.addWidget(self.access_btn)

        self.folder_label = QLabel("Dossier parent contenant tous les dossiers clients :")
        self.folder_input = QLineEdit(self.config.get("clients_parent_folder", ""))
        self.folder_btn = QPushButton("📂 Parcourir")
        self.folder_btn.clicked.connect(self.select_client_folder)

        folder_layout = QHBoxLayout()
        folder_layout.addWidget(self.folder_input)
        folder_layout.addWidget(self.folder_btn)

        self.email_label = QLabel("Adresse Gmail utilisée pour les envois :")
        self.email_input = QLineEdit(self.config.get("email_address", ""))

        self.ddt_label = QLabel("Ajouter une colonne DDT envoyé : (automatique si Gmail configuré)")

        self.client_dirs_label = QLabel("Liste des dossiers clients chargés depuis le dossier parent :")
        self.client_dirs_display = QLabel(self.get_all_client_subfolders_display())
        self.client_dirs_display.setStyleSheet("font-size: 10px; color: gray;")
        self.client_dirs_display.setWordWrap(True)

        self.save_btn = QPushButton("📅 Enregistrer le paramétrage")
        self.save_btn.clicked.connect(self.save_config)

        layout.addWidget(self.access_label)
        layout.addLayout(access_layout)
        layout.addWidget(self.folder_label)
        layout.addLayout(folder_layout)
        layout.addWidget(self.email_label)
        layout.addWidget(self.email_input)
        layout.addWidget(self.ddt_label)
        layout.addWidget(self.client_dirs_label)
        layout.addWidget(self.client_dirs_display)
        layout.addWidget(self.save_btn)

        self.setLayout(layout)

    def select_access_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choisir le fichier .mdb", "", "Base Access (*.mdb)")
        if path:
            self.access_input.setText(path)

    def select_client_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choisir le dossier contenant les sous-dossiers clients")
        if folder:
            self.folder_input.setText(folder)
            self.client_dirs_display.setText(self.get_all_client_subfolders_display(folder))

    def get_all_client_subfolders_display(self, base_path=None):
        base = base_path or self.folder_input.text()
        if not base or not os.path.isdir(base):
            return "Aucun dossier valide sélectionné."
        try:
            all_subdirs = []
            for subfolder in os.listdir(base):
                full_path = os.path.join(base, subfolder)
                if os.path.isdir(full_path) and subfolder.lower().startswith("dossiers_"):
                    year_dirs = [d for d in os.listdir(full_path) if os.path.isdir(os.path.join(full_path, d))]
                    all_subdirs.extend(year_dirs)

            access_path = self.access_input.text()
            if not os.path.exists(access_path):
                return "Base Access introuvable."

            access_names = []
            conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={access_path};"
            try:
                with pyodbc.connect(conn_str) as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT nom_dossier FROM Dossiers")
                    access_names = [normalize_folder_name(row[0]) for row in cursor.fetchall()]
            except Exception as e:
                return f"Erreur lecture base Access : {e}"

            matching = [d for d in all_subdirs if normalize_folder_name(d) in access_names]
            self.config["all_client_folders"] = [normalize_folder_name(d) for d in matching]
            return ", ".join(sorted(matching)) if matching else "Aucune correspondance entre disque et base Access."
        except Exception as e:
            return f"Erreur : {e}"

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_config(self):
        data = {
            "access_path": self.access_input.text(),
            "clients_parent_folder": self.folder_input.text(),
            "email_address": self.email_input.text(),
            "all_client_folders": self.config.get("all_client_folders", [])
        }

        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
            QMessageBox.information(self, "Succès", "Paramétrage enregistré avec succès.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Échec lors de l'enregistrement : {e}")

def normalize_folder_name(name: str) -> str:
    name = name.replace("/", "_")
    name = name.replace("\\", "_")
    name = name.replace(":", "")
    name = re.sub(r"\s+", "_", name)
    name = unicodedata.normalize("NFD", name)
    name = name.encode("ascii", "ignore").decode("utf-8")
    name = re.sub(r"[^a-zA-Z0-9_]+", "", name)
    return name.strip().lower()

def load_clients_for_main_table():
    if not os.path.exists(CONFIG_FILE):
        return []
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)

        access_path = config.get("access_path", "")
        client_ids = config.get("all_client_folders", [])
        if not access_path or not os.path.exists(access_path):
            return []

        manual_states = {}
        if os.path.exists(MANUAL_STATE_FILE):
            with open(MANUAL_STATE_FILE, "r", encoding="utf-8") as f:
                manual_states = json.load(f)

        conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={access_path};"
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT nom_dossier, type_mission, date_rdv, statut_paiement FROM Dossiers")
            rows = cursor.fetchall()

        result = []
        for row in rows:
            norm = normalize_folder_name(row[0])
            if norm in client_ids:
                dossier = row[0]
                mission = row[1] or ""
                date = row[2].strftime("%Y-%m-%d") if isinstance(row[2], datetime) else ""
                paiement = row[3] or ""
                state = manual_states.get(norm, {})
                commentaire = state.get("commentaires", "")
                ddt_envoye = state.get("ddt_envoye", False)
                result.append((dossier, mission, date, paiement, commentaire, "✅" if ddt_envoye else "❌"))

        return result
    except Exception as e:
        print("Erreur chargement clients:", e)
        return []

class MainClientTable(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tableau principal des dossiers SuiviClientPro")
        self.setGeometry(100, 100, 1200, 600)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Nom du dossier", "Type de mission", "Date RDV",
            "Statut Paiement", "Commentaires", "DDT Envoyé"
        ])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        self.setLayout(layout)

        self.load_data()

    def load_data(self):
        rows = load_clients_for_main_table()
        self.table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(val))