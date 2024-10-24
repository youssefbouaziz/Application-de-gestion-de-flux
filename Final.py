import sys
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, 
                             QTableWidget, QTableWidgetItem, QLineEdit, QFileDialog, QMessageBox)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import QRect
from PyQt5.QtCore import Qt

dico = {}

# Charger le fichier Excel
df = pd.read_excel(r"C:\Users\youss\Downloads\pf.xlsx")

# Initialiser l'index et la liste des indices
i = 0
L = []
cle = []
# Boucle pour parcourir les lignes jusqu'à l'avant-dernière ligne
while i < len(df):
    cellule_A1 = df.iloc[i, 0]
    if pd.isnull(cellule_A1):
        i += 1
    else:
        cle.append(cellule_A1)
        L.append(i + 1)
        i += 1
x = df['ref composant.0'].dropna().shape[0]
L.append(x + 2)
k = 0
for ele in cle:
    dico[ele] = (L[k], L[k + 1] - 1)
    k = k + 1

print(dico)

col = [(1, 5), (11, 12), (12, 13), (13, 17), (17, 19),(19,23),(23,25),(25,29),(7,11)]

class ExcelWindow(QWidget):
    def __init__(self, dataframe, p):
        super().__init__()
        self.dataframe = dataframe
        self.p = p

        # Initialisation de la fenêtre secondaire
        self.setWindowTitle("Contenu Excel")
        self.setGeometry(150, 150, 600, 400)

        # Création du tableau pour afficher les données
        self.table = QTableWidget()
        self.load_data()

        # Création du bouton de sauvegarde
        self.save_button = QPushButton('Sauvegarder', self)
        self.save_button.clicked.connect(self.save_data)

        # Création du layout
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.save_button)
        self.setLayout(layout)

    def load_data(self):
        # Filtrer les deux premières colonnes
        t = dico[float(m)]
        z = col[self.p]
        dataframe = self.dataframe.iloc[t[0] - 1:t[1], [5] + list(range(z[0], z[1]))]

        # Configuration du tableau avec les dimensions et les données
        self.table.setRowCount(dataframe.shape[0])
        self.table.setColumnCount(dataframe.shape[1])
        self.table.setHorizontalHeaderLabels(dataframe.columns)

        for i in range(dataframe.shape[0]):
            for j in range(dataframe.shape[1]):
                item = QTableWidgetItem(str(dataframe.iat[i, j]))
                self.table.setItem(i, j, item)

    def save_data(self):
        # Sauvegarde des modifications dans le DataFrame et fichier Excel
        t = dico[float(m)]
        z = col[self.p]
        updated_df = self.dataframe.copy()
        updated_df.iloc[t[0] - 1:t[1], [5] + list(range(z[0], z[1]))] = [
            [self.table.item(i, j).text() if self.table.item(i, j) else '' for j in range(self.table.columnCount())]
            for i in range(self.table.rowCount())
        ]
        updated_df.to_excel(r"C:\Users\youss\Downloads\pf.xlsx", index=False)

        QMessageBox.information(self, 'Sauvegarde', 'Les données ont été sauvegardées avec succès !')

class MainWindow2(QMainWindow):
    def __init__(self):
        super().__init__()

        # Créer le widget principal
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Créer un layout sans gestion de placement automatique
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\en-stock (4).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(30, 10, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\fleche-droite (1).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(160, 10, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\lavage-des-mains (1).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(270, 10, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\tri.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(520, 10, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\fleche-droite (1).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(400, 10, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\fleches-vers-le-bas.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(520, 200, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\robotique (2).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(520, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\fleche-droite (1).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(650, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\four-a-arc-electrique.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(760, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\fleche-droite (1).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(900, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\industriel.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(1040, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\fleche-droite (1).png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(1200, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\machine.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(1280, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\entretien.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(1530, 380, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\fleches-vers-le-bas.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(1530, 600, self.image_pixmap.width(), self.image_pixmap.height()))
        self.image_label = QLabel(self)
        self.image_pixmap = QPixmap(r"C:\Users\youss\Downloads\distribution.png")  # Remplace par le chemin de ton image
        self.image_label.setPixmap(self.image_pixmap)
        self.image_label.setGeometry(QRect(1530, 750, self.image_pixmap.width(), self.image_pixmap.height()))
        # Créer les boutons
        self.buttons = [
            ('magasin import', 0, QRect(0, 115, 200, 50)),
            ('lavage', 1, QRect(230, 115, 200, 50)),
            ('tri', 2, QRect(480, 115, 200, 50)),
            ('assemblage', 3, QRect(480, 500, 200, 50)),
            ('brasage et four', 4, QRect(730, 500, 200, 50)),
            ('sertissage', 5, QRect(1000, 500, 200, 50)),
            ('pliage', 6, QRect(1250, 500, 200, 50)),
            ('encombrement et ctr final',7, QRect(1500,500,200,50)),
            ('magasin export',8, QRect(1500, 872,200,50))
        ]
        self.button_widgets = []
        for text, p, geometry in self.buttons:
            button = QPushButton(text, self)
            button.setGeometry(geometry)
            button.clicked.connect(lambda _, p=p: self.open_excel_window(p))
            self.button_widgets.append(button)

        # Définir les propriétés de la fenêtre principale
        self.setWindowTitle('process')
        self.setGeometry(200, 200, 800, 600)  # Taille de la fenêtre

    def open_excel_window(self, p):
        # Lire le fichier Excel
        filepath = r"C:\Users\youss\Downloads\pf.xlsx"
        df = pd.read_excel(filepath, sheet_name='Sheet1')

        # Créer et afficher la fenêtre Excel
        self.excel_window = ExcelWindow(df, p)
        self.excel_window.show()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Créer le widget principal
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.setGeometry(100, 100, 400, 300)  # Taille moyenne

        # Ajouter une image
        self.image_label = QLabel(self)
        pixmap = QPixmap(r"C:\Users\youss\Downloads\téléchargement.png")  # Remplacer par le chemin de ton image
        self.image_label.setPixmap(pixmap)
        self.image_label.setAlignment(Qt.AlignCenter)  # Centrer l'image
        self.layout.addWidget(self.image_label)

        # Créer un label
        self.label = QLabel('Entrez ref :')
        self.layout.addWidget(self.label)


        # Créer un champ de saisie
        self.line_edit = QLineEdit()
        self.layout.addWidget(self.line_edit)

        # Création du bouton
        self.button = QPushButton('Afficher le process', self)
        self.button.clicked.connect(self.open_mainwindow2)

        # Ajouter le bouton au layout
        self.layout.addWidget(self.button)

    def open_mainwindow2(self):
        global m
        m = float(self.line_edit.text())
        print(m)

        # Lire la colonne 'ref pf' du fichier Excel
        df = pd.read_excel(r"C:\Users\youss\Downloads\pf.xlsx")

        column_values = df['ref pf'].dropna().tolist()

        if m not in column_values:
            self.label.setText("réference non trouvé")
        else:
            # Créer et afficher la fenêtre MainWindow2
            self.main_window2 = MainWindow2()
            self.main_window2.show()
            self.main_window2.setGeometry(0, 30, 1920, 1080)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
