import sys
import os
import json
import pyodbc
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QSplitter, QMessageBox, QLineEdit, QComboBox, QFileDialog, QDialog, QTextEdit, QTabWidget, QGroupBox, QFormLayout
)
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt
from config_window import ConfigWindow

CONFIG_PATH = "config_suiviclientpro.json"
MANUAL_STATES_PATH = "manual_states.json"

class FicheClientWindow(QDialog):
    def __init__(self, dossier_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Fiche client")
        self.setMinimumSize(1000, 600)

        layout = QVBoxLayout()
        tabs = QTabWidget()

        # Onglet Informations générales
        general_tab = QWidget()
        general_layout = QVBoxLayout()

        # Encadré Données dossier
        box_dossier = QGroupBox("Données du dossier")
        form_dossier = QFormLayout()
        form_dossier.addRow("Nom du dossier:", QLabel(dossier_data.get('nom_du_dossier', '')))
        form_dossier.addRow("Type de mission:", QLabel(dossier_data.get('type_de_mission', '')))
        form_dossier.addRow("Date & Heure:", QLabel(dossier_data.get('date_&_heure', '')))
        form_dossier.addRow("Statut de paiement:", QLabel(dossier_data.get('statut_paiement', '')))
        form_dossier.addRow("Assainissement:", QLabel(dossier_data.get('assainissement', '')))
        form_dossier.addRow("Statut dossier:", QLabel(dossier_data.get('dossier', '')))
        box_dossier.setLayout(form_dossier)

        # Encadré Informations client
        box_client = QGroupBox("Informations du client")
        form_client = QFormLayout()
        form_client.addRow("Nom:", QLabel(dossier_data.get('client_nom', '')))
        form_client.addRow("Prénom:", QLabel(dossier_data.get('client_prenom', '')))
        form_client.addRow("Adresse:", QLabel(dossier_data.get('client_adresse', '')))
        form_client.addRow("Code postal:", QLabel(dossier_data.get('client_cp', '')))
        form_client.addRow("Ville:", QLabel(dossier_data.get('client_ville', '')))
        form_client.addRow("Email:", QLabel(dossier_data.get('client_email', '')))
        form_client.addRow("Téléphone:", QLabel(dossier_data.get('client_tel', '')))
        box_client.setLayout(form_client)

        # Encadré Bien concerné
        box_bien = QGroupBox("Adresse du bien")
        form_bien = QFormLayout()
        form_bien.addRow("Adresse:", QLabel(dossier_data.get('bien_adresse', '')))
        form_bien.addRow("Code postal:", QLabel(dossier_data.get('bien_cp', '')))
        form_bien.addRow("Ville:", QLabel(dossier_data.get('bien_ville', '')))
        box_bien.setLayout(form_bien)

        # Encadré Donneur d'ordre
        box_donneur = QGroupBox("Donneur d'ordre")
        form_donneur = QFormLayout()
        form_donneur.addRow("Nom donneur d'ordre:", QLabel(dossier_data.get('donneur_ordre', '')))
        box_donneur.setLayout(form_donneur)

        # Encadré Facturation
        box_facturation = QGroupBox("Facturation")
        form_facturation = QFormLayout()
        form_facturation.addRow("Montant TTC:", QLabel(dossier_data.get('montant_ttc', '')))
        form_facturation.addRow("Montant payé:", QLabel(dossier_data.get('montant_paye', '')))
        form_facturation.addRow("Reste à payer:", QLabel(dossier_data.get('reste_a_payer', '')))
        box_facturation.setLayout(form_facturation)

        # Photo si présente
        if dossier_data.get("photo") and os.path.exists(dossier_data["photo"]):
            pixmap = QPixmap(dossier_data["photo"]).scaled(120, 120, Qt.KeepAspectRatio)
            image_label = QLabel()
            image_label.setPixmap(pixmap)
            general_layout.addWidget(image_label)

        general_layout.addWidget(box_dossier)
        general_layout.addWidget(box_client)
        general_layout.addWidget(box_bien)
        general_layout.addWidget(box_donneur)
        general_layout.addWidget(box_facturation)

        # Boutons
        open_folder_btn = QPushButton("📂 Ouvrir le dossier")
        open_folder_btn.clicked.connect(lambda: os.startfile(dossier_data.get("chemin", "")))
        general_layout.addWidget(open_folder_btn)

        open_email_btn = QPushButton("📧 Contacter par mail")
        open_email_btn.clicked.connect(lambda: os.startfile(f"mailto:?subject={dossier_data.get('nom_du_dossier', 'Dossier')}"))
        general_layout.addWidget(open_email_btn)

        export_pdf_btn = QPushButton("📄 Exporter la fiche en PDF")
        export_pdf_btn.clicked.connect(self.export_pdf_placeholder)
        general_layout.addWidget(export_pdf_btn)

        general_tab.setLayout(general_layout)
        tabs.addTab(general_tab, "Informations générales")

        layout.addWidget(tabs)
        self.setLayout(layout)

    def export_pdf_placeholder(self):
        QMessageBox.information(self, "Export PDF", "Fonction à implémenter pour exporter la fiche client en PDF.")




class SuiviClientPro(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SuiviClientPro - Diagnostic Immobilier")
        self.setMinimumSize(1200, 700)
        self.setWindowIcon(QIcon("icons/app_icon.png"))

        self.manual_states = {}
        self.dossiers = []
        self.filtered_dossiers = []
        self.load_manual_states()
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        self.btn_param = QPushButton("⚙️ Paramétrage")
        self.btn_actualiser = QPushButton("🔄 Actualiser")
        self.btn_statistiques = QPushButton("📊 Statistiques")
        self.btn_relances = QPushButton("📧 Relances")

        self.btn_param.clicked.connect(self.open_config)
        self.btn_actualiser.clicked.connect(self.refresh_data)
        

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Rechercher par nom...")
        self.search_input.textChanged.connect(self.apply_filters)

        self.combo_type = QComboBox()
        self.combo_type.addItem("Tous les types")
        self.combo_type.currentTextChanged.connect(self.apply_filters)

        self.combo_paiement = QComboBox()
        self.combo_paiement.addItem("Tous les paiements")
        self.combo_paiement.addItems(["Payé", "En attente"])
        self.combo_paiement.currentTextChanged.connect(self.apply_filters)

        filter_layout = QVBoxLayout()
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("Type de mission"))
        filter_layout.addWidget(self.combo_type)
        filter_layout.addWidget(QLabel("Statut paiement"))
        filter_layout.addWidget(self.combo_paiement)

        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("📁 Menu"))
        left_layout.addWidget(self.btn_param)
        left_layout.addWidget(self.btn_actualiser)
        left_layout.addWidget(self.btn_statistiques)
        left_layout.addWidget(self.btn_relances)
        left_layout.addStretch()
        left_layout.addLayout(filter_layout)

        menu_widget = QWidget()
        menu_widget.setLayout(left_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "Nom du client", "Type de mission", "Date & Heure", "Paiement",
            "Assainissement", "Dossier", "Commentaires", "Photo", "DDT envoyé"
        ])

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QTableWidget.AllEditTriggers)
        self.table.itemChanged.connect(self.save_manual_states)
        self.table.itemDoubleClicked.connect(self.open_fiche_client)

        split = QSplitter()
        split.addWidget(menu_widget)
        split.addWidget(self.table)
        split.setSizes([300, 900])

        layout = QHBoxLayout()
        layout.addWidget(split)
        central_widget.setLayout(layout)

        self.refresh_data()

    def open_fiche_client(self, item):
        if item.column() != 0:  # Colonne 0 = "Nom du dossier"
            return
        row = item.row()
        dossier_data = self.get_dossier_data_from_row(row)
        if dossier_data:
            fiche_window = FicheClientWindow(dossier_data, self)
            fiche_window.exec_()


    def load_manual_states(self):
        if os.path.exists(MANUAL_STATES_PATH):
            with open(MANUAL_STATES_PATH, 'r', encoding='utf-8') as f:
                self.manual_states = json.load(f)

    def load_config(self):
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}

    def get_dossiers_from_access(self):
        config = self.load_config()
        db_path = config.get("access_path", "")
        dossiers = []

        if not os.path.exists(db_path):
            QMessageBox.critical(self, "Erreur", f"Fichier introuvable : {db_path}")
            return dossiers

        try:
            conn_str = (
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                f"DBQ={db_path};"
            )
            with pyodbc.connect(conn_str) as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT Num_dossier, type_de_dossier, Date_RDV, dossier_etat_paie, photo_de_presentation, dossier_Acces
                    FROM Donnees_Dossiers
                """)
                for row in cursor.fetchall():
                    try:
                        raw_date = row.Date_RDV
                        if isinstance(raw_date, datetime):
                            date_str = raw_date.strftime("%d/%m/%Y")
                        elif isinstance(raw_date, str):
                            date_str = raw_date
                        elif isinstance(raw_date, int):
                            try:
                                date_str = datetime.strptime(str(raw_date), "%Y%m%d").strftime("%d/%m/%Y")
                            except:
                                date_str = str(raw_date)
                        else:
                            date_str = ""

                        dossiers.append({
                            "nom": str(row.Num_dossier),
                            "type": str(row.type_de_dossier or ""),
                            "date": date_str,
                            "paiement": str(row.dossier_etat_paie or ""),
                            "photo": str(row.photo_de_presentation or ""),
                            "chemin": str(row.dossier_Acces or "")
                        })
                    except Exception as e:
                        print(f"[Erreur ligne] {row}: {e}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur Base Access", str(e))

        return dossiers

    def refresh_data(self):
        self.dossiers = self.get_dossiers_from_access()
        self.update_filter_options()
        self.apply_filters()

    def apply_filters(self):
        search_text = self.search_input.text().lower()
        selected_type = self.combo_type.currentText()
        selected_paiement = self.combo_paiement.currentText()

        self.filtered_dossiers = []
        for dossier in self.dossiers:
            if search_text and search_text not in dossier["nom"].lower():
                continue
            if selected_type != "Tous les types" and dossier["type"] != selected_type:
                continue
            if selected_paiement != "Tous les paiements" and dossier["paiement"] != selected_paiement:
                continue
            self.filtered_dossiers.append(dossier)

        self.update_table()

    def update_filter_options(self):
        types = sorted(set(d["type"] for d in self.dossiers if d["type"]))
        self.combo_type.clear()
        self.combo_type.addItem("Tous les types")
        self.combo_type.addItems(types)

    def update_table(self):
        self.table.setRowCount(len(self.filtered_dossiers))

        for row, dossier in enumerate(self.filtered_dossiers):
            self.table.setItem(row, 0, QTableWidgetItem(dossier["nom"]))
            self.table.setItem(row, 1, QTableWidgetItem(dossier["type"]))
            self.table.setItem(row, 2, QTableWidgetItem(dossier["date"]))
            self.table.setItem(row, 3, QTableWidgetItem(dossier["paiement"]))

            assainissement = self.manual_states.get(dossier["nom"], {}).get("assainissement", "")
            dossier_statut = self.manual_states.get(dossier["nom"], {}).get("dossier", "")
            commentaire = self.manual_states.get(dossier["nom"], {}).get("commentaire", "")

            self.table.setItem(row, 4, QTableWidgetItem(assainissement))
            self.table.setItem(row, 5, QTableWidgetItem(dossier_statut))
            self.table.setItem(row, 6, QTableWidgetItem(commentaire))

    def save_manual_states(self, item):
        row = item.row()
        col = item.column()
        if col >= 4:
            dossier = self.table.item(row, 0).text()
            if dossier not in self.manual_states:
                self.manual_states[dossier] = {}
            if col == 4:
                self.manual_states[dossier]["assainissement"] = item.text()
            elif col == 5:
                self.manual_states[dossier]["dossier"] = item.text()
            elif col == 6:
                self.manual_states[dossier]["commentaire"] = item.text()

            with open(MANUAL_STATES_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.manual_states, f, indent=2, ensure_ascii=False)

    def open_config(self):
        config_dialog = ConfigWindow(self)
        if config_dialog.exec_():
            self.refresh_data()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SuiviClientPro()
    window.show()
    sys.exit(app.exec_())
