import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout,
    QLabel, QDialog, QTabWidget, QGroupBox, QFormLayout, QMessageBox
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

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





if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FicheClientWindow({})
    window.show()
    sys.exit(app.exec_())
