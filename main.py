"""
Sistema de Cálculo Excel Profissional
Arquivo principal da aplicação

Substitui funcionalidades VBA por Python moderno com interface PyQt6
Desenvolvido para processamento avançado de planilhas Excel
"""

import sys
import os

# Adiciona o diretório src ao path para imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QPalette, QColor
import traceback

# Importa o gerenciador de licença
from src.license_manager import check_license_and_run
from src.auto_updater import check_updates_on_startup


def setup_application_style(app):
    """Configura estilo global da aplicação"""
    app.setStyle('Fusion')  # Estilo moderno multiplataforma
    
    # Define paleta de cores personalizada minimalista
    palette = QPalette()
    
    # Cores principais mais sutis para não interferir com a tabela
    palette.setColor(QPalette.ColorRole.Window, QColor(30, 30, 30))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.Base, QColor(40, 40, 40))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(50, 50, 50))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(44, 62, 80))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.Text, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.Button, QColor(52, 152, 219))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(231, 76, 60))
    palette.setColor(QPalette.ColorRole.Link, QColor(52, 152, 219))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(52, 152, 219))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    
    app.setPalette(palette)

def handle_exception(exc_type, exc_value, exc_traceback):
    """Manipulador global de exceções"""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
        
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    
    # Log do erro (em uma aplicação real, salvaria em arquivo)
    print(f"Erro não tratado: {error_msg}")
    
    # Mostra mensagem para o usuário
    try:
        app = QApplication.instance()
        if app:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Icon.Critical)
            msg_box.setWindowTitle("Erro do Sistema")
            msg_box.setText("Ocorreu um erro inesperado na aplicação.")
            msg_box.setDetailedText(error_msg)
            msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg_box.exec()
    except:
        pass  # Se não conseguir mostrar a mensagem, só imprime no console

def main():
    """Função principal da aplicação"""
    try:
        # Configura o manipulador de exceções
        sys.excepthook = handle_exception
        
        # Cria aplicação Qt
        app = QApplication(sys.argv)
        
        # Configura propriedades da aplicação
        app.setApplicationName("Sistema de Cálculo Excel Profissional")
        app.setApplicationVersion("1.0.0")
        app.setOrganizationName("Mega Systems")
        app.setOrganizationDomain("megasystems.com")
        
        # Define ícone da aplicação
        icon_path = os.path.join(os.path.dirname(__file__), "..", "icone.ico")
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
        
        # Configura estilo
        setup_application_style(app)
        
        # Importa e cria janela principal
        try:
            from src.interface import MainWindow
        except ImportError as e:
            QMessageBox.critical(
                None, 
                "Erro de Importação", 
                f"Erro ao importar módulos necessários:\n{str(e)}\n\n"
                f"Verifique se todas as dependências estão instaladas:\n"
                f"pip install PyQt6 pandas openpyxl"
            )
            return 1
            
        # Cria e mostra janela principal
        window = MainWindow()
        window.show()
        
        # Centraliza janela na tela
        screen = app.primaryScreen().geometry()
        size = window.geometry()
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        window.move(x, y)
        
        # Verifica atualizações automaticamente (em segundo plano)
        try:
            check_updates_on_startup(window)
        except Exception as e:
            print(f"Erro ao verificar atualizações: {e}")
        
        # Inicia loop da aplicação
        return app.exec()
        
    except Exception as e:
        error_msg = f"Erro crítico ao iniciar aplicação: {str(e)}"
        print(error_msg)
        
        try:
            QMessageBox.critical(
                None,
                "Erro Crítico",
                error_msg
            )
        except:
            print("Não foi possível exibir mensagem de erro")
            
        return 1

if __name__ == "__main__":
    # Verifica versão do Python
    if sys.version_info < (3, 7):
        print("ERRO: Este sistema requer Python 3.7 ou superior")
        print(f"Versão atual: {sys.version}")
        sys.exit(1)
        
    # Verifica se está executando no diretório correto
    if not os.path.exists(os.path.join(os.path.dirname(__file__), 'src')):
        print("ERRO: Diretório 'src' não encontrado")
        print("Execute o script do diretório correto do projeto")
        sys.exit(1)
    
    # Verifica licença antes de executar o aplicativo
    if not check_license_and_run():
        print("Licença não válida. Aplicação encerrada.")
        sys.exit(1)
        
    # Se chegou até aqui, a licença é válida - continua execução
    exit_code = main()
    sys.exit(exit_code)