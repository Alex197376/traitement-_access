import sys
import os
import json
import pyodbc
from fiche_client_window import FicheClientWindow
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QSplitter, QMessageBox, QLineEdit, QComboBox, QFileDialog
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from config_window import ConfigWindow



CONFIG_PATH = "config_suiviclientpro.json"
MANUAL_STATES_PATH = "manual_states.json"

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def validate_json_file(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            json.load(f)
        return True
    except json.JSONDecodeError as e:
        QMessageBox.critical(None, "Erreur JSON", f"Erreur dans le fichier JSON :\n{e}")
        return False
    
CONFIG_PATH = "config_suiviclientpro.json"
MANUAL_STATES_PATH = "manual_states.json"

class SuiviClientPro(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SuiviClientPro - Diagnostic Immobilier")
        self.setMinimumSize(1200, 700)
        self.setWindowIcon(QIcon("icons/app_icon.png"))

        self.manual_states = {}
        self.dossiers = []
        self.filtered_dossiers = []
        self.sorted_column = -1
        self.sort_order = Qt.AscendingOrder

        self.load_manual_states()
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        self.btn_param = QPushButton("⚙️ Paramétrage")
        self.btn_actualiser = QPushButton("🔄 Actualiser")
        self.btn_statistiques = QPushButton("📊 Statistiques")
        self.btn_relances = QPushButton("📧 Relances")
        self.btn_reset_sort = QPushButton("🔁 Réinitialiser tri")
       

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
        left_layout.addWidget(self.btn_reset_sort)
        left_layout.addStretch()
        left_layout.addLayout(filter_layout)

        menu_widget = QWidget()
        menu_widget.setLayout(left_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Nom du dossier", "Type de mission", "Date & Heure", "Statut paiement",
            "Assainissement", "Dossier", "Commentaires"])
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().sectionClicked.connect(self.handle_sorting)

        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QHeaderView::section {
                background-color: #f2f2f2;
                font-weight: bold;
                padding: 6px;
                border-bottom: 1px solid #aaa;
            }
            QTableWidget {
                alternate-background-color: #fafafa;
                background-color: #ffffff;
                border-radius: 4px;
                border: 1px solid #ccc;
            }
            QTableWidget::item:hover {
                background-color: #e0f0ff;
            }
        """)

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(
            QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed
)
        self.table.itemDoubleClicked.connect(self.handle_double_click)

        self.table.itemChanged.connect(self.save_manual_states)

        split = QSplitter()
        split.addWidget(menu_widget)
        split.addWidget(self.table)
        split.setSizes([300, 900])

        layout = QHBoxLayout()
        layout.addWidget(split)
        central_widget.setLayout(layout)

        self.refresh_data()

    def load_manual_states(self):
        if not os.path.exists(MANUAL_STATES_PATH):
            self.manual_states = {}
            return

        if not validate_json_file(MANUAL_STATES_PATH):
            self.manual_states = {}
            return

        with open(MANUAL_STATES_PATH, 'r', encoding='utf-8') as f:
            self.manual_states = json.load(f)

    def load_config(self):
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            QMessageBox.warning(self, "Configuration manquante", "Le fichier de configuration est introuvable.")
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
                    SELECT Num_dossier, type_de_dossier, rdv_date, rdv_heure, dossier_etat_paie, photo_de_presentation, dossier_Acces
                    FROM Donnees_Dossiers
                """)

                for row in cursor.fetchall():
                    try:
                        raw_date = row.rdv_date
                        raw_time = row.rdv_heure

                        # Traitement de la date
                        if isinstance(raw_date, datetime):
                            date_str = raw_date.strftime("%d/%m/%Y")
                        elif isinstance(raw_date, str):
                            date_str = raw_date
                        else:
                            date_str = ""

                        # Traitement de l'heure (déjà en texte normalement)
                        heure_str = str(raw_time or "")

                        # Combinaison propre : "31/03/2021 09 h 00"
                        date_heure = f"{date_str} {heure_str}".strip()


                        dossiers.append({
                            "nom": str(row.Num_dossier),
                            "type": str(row.type_de_dossier or ""),
                            "date": date_heure,
                            "paiement": str(row.dossier_etat_paie or ""),
                            "photo": str(row.photo_de_presentation or ""),
                            "chemin": str(row.dossier_Acces or "")
                        })
                    except Exception as e:
                        print(f"[Erreur ligne] {row}: {e}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur Base Access", str(e))

        return dossiers
    
    def handle_double_click(self, item):
        if item.column() == 0:  # Colonne "Nom du client"
            row = item.row()
            dossier_data = self.get_dossier_data_from_row(row)
            print("Contenu du dossier sélectionné :", dossier_data)  # Déplacer ici
            if dossier_data:
                from fiche_client_window import FicheClientWindow
                fiche = FicheClientWindow(dossier_data, self)
                fiche.exec_()

    def get_dossier_data_from_row(self, row):
        nom_dossier = self.table.item(row, 0).text()

        config = self.load_config()
        chemin_base = config.get("access_path")
        if not chemin_base or not os.path.exists(chemin_base):
            QMessageBox.warning(self, "Erreur", "Chemin de la base de données Access non défini ou introuvable.")
            return {}

        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={chemin_base};'
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            # Requête pour récupérer toutes les données du dossier
            cursor.execute("""
                SELECT *
                FROM Dossiers
                WHERE nom_dossier = ?
            """, (nom_dossier,))

            row = cursor.fetchone()
            if not row:
                QMessageBox.warning(self, "Introuvable", f"Aucun dossier trouvé dans la base Access pour {nom_dossier}")
                return {}

            # Construction du dictionnaire de données
            dossier_data = {
                "nom_du_dossier": row.nom_dossier,
                "type_de_mission": row.type_mission,
                "date_&_heure": row.date_rdv.strftime("%d/%m/%Y %H h %M") if row.date_rdv else "",
                "statut_paiement": row.statut_paiement or "",
                "assainissement": row.assainissement or "",
                "dossier": row.statut_dossier or "",
                "commentaires": row.commentaires or "",
                "montant_ttc": f"{row.facturation_ttc:.2f} €" if row.facturation_ttc else "",
                "montant_paye": f"{row.facturation_paye:.2f} €" if row.facturation_paye else "",
                "reste_a_payer": f"{row.facturation_restante:.2f} €" if row.facturation_restante else "",
                "client_nom": row.client_nom or "",
                "client_prenom": row.client_prenom or "",
                "client_adresse": row.client_adresse or "",
                "client_cp": row.client_cp or "",
                "client_ville": row.client_ville or "",
                "client_email": row.client_email or "",
                "client_tel": row.client_tel or "",
                "bien_adresse": row.bien_adresse or "",
                "bien_cp": row.bien_cp or "",
                "bien_ville": row.bien_ville or "",
                "donneur_ordre": row.donneur_ordre or "",
                "chemin": row.chemin_dossier or "",
                "photo": "",  # à compléter dans une étape suivante
            }

            conn.close()
            return dossier_data

        except Exception as e:
            QMessageBox.critical(self, "Erreur Base Access", str(e))
            return {}


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

        if self.sorted_column >= 0:
            self.filtered_dossiers.sort(
                key=lambda x: list(x.values())[self.sorted_column],
                reverse=self.sort_order == Qt.DescendingOrder
            )

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

    def handle_sorting(self, column):
        header = self.table.horizontalHeader()
        order = header.sortIndicatorOrder()
        self.sorted_column = column
        self.sort_order = order
        self.apply_filters()

    def reset_sort(self):
        self.sorted_column = -1
        self.sort_order = Qt.AscendingOrder
        self.apply_filters()

    def open_fiche_client(self, item):
        row = item.row()
        nom_dossier = self.table.item(row, 0).text()

        dossier = next((d for d in self.dossiers if d["nom"] == nom_dossier), None)
        if dossier:
            fiche = FicheClientWindow(dossier)
            fiche.exec_()



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SuiviClientPro()
    window.show()
    sys.exit(app.exec_())


