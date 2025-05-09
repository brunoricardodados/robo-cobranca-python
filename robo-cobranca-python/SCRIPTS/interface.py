import sys
import subprocess
import os
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel
from PySide6.QtGui import QPixmap, QPalette, QBrush
from PySide6.QtCore import Qt, QSize

class TelaPrincipal(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Order to Cash - Sistema Contas a Receber Automatizado")

        # Definir um tamanho fixo para a janela
        self.fixed_width = 600
        self.fixed_height = 450
        self.setFixedSize(QSize(self.fixed_width, self.fixed_height))

        # Caminho base do projeto (diretório do script)
        self.pasta_base = os.path.dirname(os.path.abspath(__file__))

        # Background
        palette = QPalette()
        fundo_path = os.path.join(self.pasta_base, "background.png")
        fundo = QPixmap(fundo_path)
        palette.setBrush(QPalette.Window, QBrush(fundo))
        self.setPalette(palette)
        self.setAutoFillBackground(True)

        layout = QVBoxLayout()

        # Logo
        self.logo = QLabel(self)
        logo_path = os.path.join(self.pasta_base, "logo_waters.png")
        logo_pixmap = QPixmap(logo_path)
        self.logo.setPixmap(logo_pixmap)
        self.logo.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        layout.addWidget(self.logo)

        # Mensagem de Boas-Vindas
        self.labelBoasVindas = QLabel("Bem vindo ao Oder to Cash - Sistema Automatizado de Cobrança e Consulta", self)
        self.labelBoasVindas.setAlignment(Qt.AlignCenter)
        self.labelBoasVindas.setStyleSheet("color: white; font-size: 18px; font-weight: bold; margin-top: 20px;")
        layout.addWidget(self.labelBoasVindas)

        # Layout para os botões (vertical)
        layout_botoes = QVBoxLayout()
        layout_botoes.setSpacing(10)

        # Estilo dos botões
        botao_estilo = """
            QPushButton {
                background-color: #2c3e50;
                color: white;
                border-radius: 5px;
                padding: 10px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #34495e;
                color: yellow;
            }
        """

        # Botões na ordem correta e com os nomes certos
        self.btnAtualizarAging = QPushButton("Atualizar Aging")
        self.btnCobrancaProativa = QPushButton("Cobrança Proativa")
        self.btnCobrancaReativa = QPushButton("Cobrança Reativa")
        self.btnRelatorioPowerBI = QPushButton("Gerar Relatório Power BI")

        self.btnAtualizarAging.setStyleSheet(botao_estilo)
        self.btnCobrancaProativa.setStyleSheet(botao_estilo)
        self.btnCobrancaReativa.setStyleSheet(botao_estilo)
        self.btnRelatorioPowerBI.setStyleSheet(botao_estilo)

        layout_botoes.addWidget(self.btnAtualizarAging)
        layout_botoes.addWidget(self.btnCobrancaProativa)
        layout_botoes.addWidget(self.btnCobrancaReativa)
        layout_botoes.addWidget(self.btnRelatorioPowerBI)

        layout.addLayout(layout_botoes)
        layout.setAlignment(Qt.AlignCenter)

        self.setLayout(layout)
        self.show()

        # Conexões dos botões (na ordem correta)
        self.btnAtualizarAging.clicked.connect(self.executar_atualizar_aging)
        self.btnCobrancaProativa.clicked.connect(self.executar_cobranca_proativa)
        self.btnCobrancaReativa.clicked.connect(self.executar_cobranca_reativa)
        self.btnRelatorioPowerBI.clicked.connect(self.executar_power_bi_direto)

    def executar_atualizar_aging(self):
        script = os.path.join(self.pasta_base, "abrir_aging.py")
        subprocess.run(["python", script])

    def executar_cobranca_proativa(self):
        script = os.path.join(self.pasta_base, "enviar_cobranca.py")
        subprocess.run(["python", script])

    def executar_cobranca_reativa(self):
        script = os.path.join(self.pasta_base, "reativa_cobranca.py")
        subprocess.run(["python", script])

    def executar_power_bi_direto(self):
        script = os.path.join(self.pasta_base, "automatizar_power_bi.py")
        subprocess.run(["python", script])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    tela = TelaPrincipal()
    sys.exit(app.exec())
