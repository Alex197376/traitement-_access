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
from PyQt5.QtGui import QIcon, QColor, QBrush
from PyQt5.QtCore import Qt
from config_window import ConfigWindow
from scan_ddt_envoyes import telecharger_pieces_jointes




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
        self.btn_ddt = QPushButton("📤 DDT envoyés")
        self.btn_ddt.clicked.connect(self.actualiser_ddt_envoyes)
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
        left_layout.addWidget(self.btn_ddt)
        left_layout.addWidget(self.btn_reset_sort)
        left_layout.addStretch()
        left_layout.addLayout(filter_layout)

        menu_widget = QWidget()
        menu_widget.setLayout(left_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "Nom du dossier", "Type de mission", "Date & Heure", "Statut paiement",
            "Assainissement", "Dossier", "Commentaires", "DDT envoyé"
        ])

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
            cursor.execute(
                """
                SELECT *
                FROM Donnees_Dossiers
                WHERE Num_dossier = ?
                """,
                (nom_dossier,),
            )

            row = cursor.fetchone()
            if not row:
                QMessageBox.warning(self, "Introuvable", f"Aucun dossier trouvé dans la base Access pour {nom_dossier}")
                return {}

            # Construction du dictionnaire de données
            date_part = getattr(row, "rdv_date", None)
            heure_part = getattr(row, "rdv_heure", "")
            date_str = date_part.strftime("%d/%m/%Y") if isinstance(date_part, datetime) else ""
            date_heure = f"{date_str} {heure_part}".strip()

            dossier_data = {
                "nom_du_dossier": getattr(row, "Num_dossier", ""),
                "type_de_mission": getattr(row, "type_de_dossier", ""),
                "date_&_heure": date_heure,
                "statut_paiement": getattr(row, "dossier_etat_paie", ""),
                "assainissement": getattr(row, "assainissement", ""),
                "dossier": getattr(row, "statut_dossier", ""),
                "commentaires": getattr(row, "commentaires", ""),
                "montant_ttc": f"{getattr(row, 'facturation_ttc', 0):.2f} €" if getattr(row, 'facturation_ttc', None) else "",
                "montant_paye": f"{getattr(row, 'facturation_paye', 0):.2f} €" if getattr(row, 'facturation_paye', None) else "",
                "reste_a_payer": f"{getattr(row, 'facturation_restante', 0):.2f} €" if getattr(row, 'facturation_restante', None) else "",
                "client_nom": getattr(row, "client_nom", ""),
                "client_prenom": getattr(row, "client_prenom", ""),
                "client_adresse": getattr(row, "client_adresse", ""),
                "client_cp": getattr(row, "client_cp", ""),
                "client_ville": getattr(row, "client_ville", ""),
                "client_email": getattr(row, "client_email", ""),
                "client_tel": getattr(row, "client_tel", ""),
                "bien_adresse": getattr(row, "bien_adresse", ""),
                "bien_cp": getattr(row, "bien_cp", ""),
                "bien_ville": getattr(row, "bien_ville", ""),
                "donneur_ordre": getattr(row, "donneur_ordre", ""),
                "chemin": getattr(row, "chemin_dossier", ""),
                "photo": "",
            }

            conn.close()
            return dossier_data

        except Exception as e:
            QMessageBox.critical(self, "Erreur Base Access", str(e))
            return {}
        

    def verifier_ddt_local(self, nom_dossier):
        config = self.load_config()
        base_path = config.get("dossiers_path", "")
        if not base_path:
            return False

        # Dossier complet à analyser
        dossier_path = os.path.join(base_path, nom_dossier)
        if not os.path.exists(dossier_path):
            return False

        # Vérifie si un PDF DDT est présent
        for fichier in os.listdir(dossier_path):
            if fichier.lower().endswith(".pdf") and any(mot in fichier.lower() for mot in ["dpe", "amiante", "ddt", "rapport"]):
                return True
        return False



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

        self.table.blockSignals(True)
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
            ddt_present = self.verifier_ddt_local(dossier["nom"])
            ddt_item = QTableWidgetItem("Oui" if ddt_present else "Non")
            ddt_item.setForeground(QBrush(QColor("green" if ddt_present else "red")))
            self.table.setItem(row, 7, ddt_item)
           
            ddt_item = QTableWidgetItem("Non")
            ddt_item.setForeground(QBrush(QColor("red")))
            self.table.setItem(row, 7, ddt_item)


        self.table.blockSignals(False)

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

            # Mise à jour dans la base Access
            config = self.load_config()
            db_path = config.get("access_path", "")
            if not os.path.exists(db_path):
                QMessageBox.warning(self, "Erreur", f"Base Access introuvable : {db_path}")
                return

            conn_str = (
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                f"DBQ={db_path};"
            )
            try:
                with pyodbc.connect(conn_str) as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        """
                        UPDATE Donnees_Dossiers
                        SET assainissement = ?, statut_dossier = ?, commentaires = ?
                        WHERE Num_dossier = ?
                        """,
                        (
                            self.manual_states[dossier].get("assainissement", ""),
                            self.manual_states[dossier].get("dossier", ""),
                            self.manual_states[dossier].get("commentaire", ""),
                            dossier,
                        ),
                    )
                    conn.commit()
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Erreur",
                    f"Échec de l'écriture dans la base Access : {e}",
                )

    def actualiser_ddt_envoyes(self):
        try:
            fichiers_gmail = telecharger_pieces_jointes()
        except Exception as e:
            QMessageBox.critical(self, "Erreur Gmail", f"Erreur lors du scan Gmail : {e}")
            return

        fichiers_gmail = [f.lower() for f in fichiers_gmail]

        for row in range(self.table.rowCount()):
            nom_dossier = self.table.item(row, 0).text().lower()
            ddt_local = self.verifier_ddt_local(nom_dossier)

            # Correspondance approximative entre nom_dossier et fichier Gmail
            trouve_gmail = any(nom_dossier in f for f in fichiers_gmail)

            if ddt_local or trouve_gmail:
                statut = "Oui"
                couleur = "green"
            else:
                statut = "Non"
                couleur = "red"

            ddt_item = QTableWidgetItem(statut)
            ddt_item.setForeground(QBrush(QColor(couleur)))
            self.table.setItem(row, 7, ddt_item)

        QMessageBox.information(self, "Scan terminé", "La mise à jour des DDT envoyés est terminée.")




    def open_config(self):
        config_dialog = ConfigWindow(self)
        if config_dialog.exec_():
            self.refresh_data()
        from config_window import GmailConfigDialog
        gmail_dialog = GmailConfigDialog(self)
        gmail_dialog.exec_()


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

    def actualiser_ddt_envoyes(self):
        QMessageBox.information(self, "DDT envoyés", "Fonction à venir : scan Gmail + mise à jour DDT.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SuiviClientPro()
    window.show()
    sys.exit(app.exec_())


