# history_dialog.py

from PyQt5.QtWidgets import (QDialog, QHBoxLayout, QVBoxLayout, QGroupBox, 
                             QListWidget, QTableWidget, QTableWidgetItem, 
                             QPushButton, QMessageBox)

# A ÚNICA importação de outro módulo do nosso projeto deve ser esta:
from history_manager import HistoryManager

class HistoryDialog(QDialog):
    def __init__(self, history_manager: HistoryManager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Histórico de Projetos")
        self.history_manager = history_manager
        self.loaded_project_data = None
        self.setMinimumSize(800, 600)

        layout = QHBoxLayout(self)

        # Painel esquerdo com a lista de projetos
        left_panel = QGroupBox("Projetos Salvos")
        left_layout = QVBoxLayout()
        self.project_list_widget = QListWidget()
        left_layout.addWidget(self.project_list_widget)
        left_panel.setLayout(left_layout)

        # Painel direito com os detalhes das peças
        right_panel = QGroupBox("Peças do Projeto")
        right_layout = QVBoxLayout()
        self.pieces_table_widget = QTableWidget()
        right_layout.addWidget(self.pieces_table_widget)

        button_layout = QHBoxLayout()
        self.load_btn = QPushButton("Carregar Projeto na Lista")
        self.delete_btn = QPushButton("Excluir Projeto do Histórico")
        close_btn = QPushButton("Fechar")
        
        button_layout.addWidget(self.load_btn)
        button_layout.addWidget(self.delete_btn)
        button_layout.addStretch()
        button_layout.addWidget(close_btn)
        right_layout.addLayout(button_layout)
        right_panel.setLayout(right_layout)

        layout.addWidget(left_panel, 1)
        layout.addWidget(right_panel, 3)
        
        # Conexões
        self.project_list_widget.currentItemChanged.connect(self.display_project_details)
        self.load_btn.clicked.connect(self.load_project)
        self.delete_btn.clicked.connect(self.delete_project)
        close_btn.clicked.connect(self.reject)
        
        # Inicialização
        self.populate_project_list()
        self.update_buttons_state()

    def populate_project_list(self):
        self.project_list_widget.clear()
        projects = self.history_manager.get_projects()
        self.project_list_widget.addItems(projects)

    def display_project_details(self, current, previous):
        self.pieces_table_widget.setRowCount(0)
        if not current:
            self.update_buttons_state()
            return

        project_number = current.text()
        pieces = self.history_manager.get_project_data(project_number)

        if pieces:
            # Pega as chaves do primeiro dicionário como cabeçalhos
            headers = list(pieces[0].keys()) if pieces else []
            self.pieces_table_widget.setColumnCount(len(headers))
            self.pieces_table_widget.setHorizontalHeaderLabels([h.replace('_', ' ').title() for h in headers])
            self.pieces_table_widget.setRowCount(len(pieces))

            for i, piece in enumerate(pieces):
                for j, key in enumerate(headers):
                    value = piece.get(key)
                    if key == 'furos' and isinstance(value, list):
                        display_text = f"{len(value)} Furo(s)"
                    else:
                        display_text = str(value if value is not None else '-')
                    self.pieces_table_widget.setItem(i, j, QTableWidgetItem(display_text))
        
        self.pieces_table_widget.resizeColumnsToContents()
        self.update_buttons_state()

    def update_buttons_state(self):
        has_selection = self.project_list_widget.currentItem() is not None
        self.load_btn.setEnabled(has_selection)
        self.delete_btn.setEnabled(has_selection)

    def load_project(self):
        if self.project_list_widget.currentItem():
            project_number = self.project_list_widget.currentItem().text()
            self.loaded_project_data = self.history_manager.get_project_data(project_number)
            self.accept()

    def delete_project(self):
        if self.project_list_widget.currentItem():
            project_number = self.project_list_widget.currentItem().text()
            reply = QMessageBox.question(self, 'Excluir Projeto', 
                                         f"Tem certeza que deseja excluir o projeto '{project_number}' do histórico?", 
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.history_manager.delete_project(project_number)
                self.populate_project_list()
                self.pieces_table_widget.setRowCount(0)