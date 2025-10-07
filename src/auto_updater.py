"""
Sistema de Atualização Automática
Verifica e baixa atualizações do aplicativo automaticamente
"""

import os
import sys
import json
import requests
import zipfile
import shutil
import subprocess
from pathlib import Path
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QProgressBar, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont
import tempfile

# Configurações do servidor de atualização
UPDATE_SERVER_URL = "https://github.com/seu-usuario/excel-calculator/releases"  # Substitua pela sua URL
VERSION_CHECK_URL = "https://api.github.com/repos/seu-usuario/excel-calculator/releases/latest"
CURRENT_VERSION = "1.0.0"  # Versão atual do aplicativo

class UpdateChecker(QThread):
    """Thread para verificar atualizações em segundo plano"""
    update_available = pyqtSignal(dict)
    no_update = pyqtSignal()
    error = pyqtSignal(str)
    
    def run(self):
        try:
            # Verifica versão mais recente
            response = requests.get(VERSION_CHECK_URL, timeout=10)
            if response.status_code == 200:
                release_info = response.json()
                latest_version = release_info['tag_name'].replace('v', '')
                
                if self.is_newer_version(latest_version, CURRENT_VERSION):
                    # Procura pelo arquivo executável
                    download_url = None
                    for asset in release_info['assets']:
                        if asset['name'].endswith('.exe'):
                            download_url = asset['browser_download_url']
                            break
                    
                    if download_url:
                        update_info = {
                            'version': latest_version,
                            'download_url': download_url,
                            'release_notes': release_info['body'],
                            'file_size': asset['size']
                        }
                        self.update_available.emit(update_info)
                    else:
                        self.error.emit("Arquivo de atualização não encontrado")
                else:
                    self.no_update.emit()
            else:
                self.error.emit(f"Erro ao verificar atualizações: {response.status_code}")
                
        except requests.exceptions.RequestException as e:
            self.error.emit(f"Erro de conexão: {str(e)}")
        except Exception as e:
            self.error.emit(f"Erro inesperado: {str(e)}")
    
    def is_newer_version(self, latest, current):
        """Compara versões usando versionamento semântico"""
        try:
            latest_parts = [int(x) for x in latest.split('.')]
            current_parts = [int(x) for x in current.split('.')]
            
            # Padroniza com zeros se necessário
            max_len = max(len(latest_parts), len(current_parts))
            latest_parts.extend([0] * (max_len - len(latest_parts)))
            current_parts.extend([0] * (max_len - len(current_parts)))
            
            return latest_parts > current_parts
        except:
            return False

class UpdateDownloader(QThread):
    """Thread para download da atualização"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)  # Caminho do arquivo baixado
    error = pyqtSignal(str)
    
    def __init__(self, download_url, file_size):
        super().__init__()
        self.download_url = download_url
        self.file_size = file_size
        
    def run(self):
        try:
            # Cria arquivo temporário
            temp_dir = tempfile.mkdtemp()
            temp_file = os.path.join(temp_dir, "ExcelCalculator_Update.exe")
            
            # Download com progress
            response = requests.get(self.download_url, stream=True, timeout=30)
            response.raise_for_status()
            
            downloaded = 0
            with open(temp_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if self.file_size > 0:
                            progress = int((downloaded / self.file_size) * 100)
                            self.progress.emit(progress)
            
            self.finished.emit(temp_file)
            
        except Exception as e:
            self.error.emit(f"Erro no download: {str(e)}")

class UpdateDialog(QDialog):
    """Dialog para mostrar informações da atualização"""
    
    def __init__(self, update_info, parent=None):
        super().__init__(parent)
        self.update_info = update_info
        self.temp_file = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Atualização Disponível")
        self.setFixedSize(500, 400)
        self.setModal(True)
        
        layout = QVBoxLayout()
        
        # Título
        title = QLabel(f"Nova versão disponível: v{self.update_info['version']}")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # Notas da versão
        notes_label = QLabel("Novidades:")
        notes_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        layout.addWidget(notes_label)
        
        notes = QLabel(self.update_info['release_notes'][:300] + "...")
        notes.setWordWrap(True)
        notes.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(notes)
        
        # Tamanho do arquivo
        size_mb = self.update_info['file_size'] / (1024 * 1024)
        size_label = QLabel(f"Tamanho: {size_mb:.1f} MB")
        layout.addWidget(size_label)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Botões
        button_layout = QVBoxLayout()
        
        self.download_btn = QPushButton("⬇️ Baixar e Instalar")
        self.download_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                font-size: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.download_btn.clicked.connect(self.start_download)
        button_layout.addWidget(self.download_btn)
        
        later_btn = QPushButton("⏰ Lembrar Mais Tarde")
        later_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px;
                font-size: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)
        later_btn.clicked.connect(self.reject)
        button_layout.addWidget(later_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
    def start_download(self):
        """Inicia o download da atualização"""
        self.download_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        
        # Inicia thread de download
        self.downloader = UpdateDownloader(
            self.update_info['download_url'],
            self.update_info['file_size']
        )
        self.downloader.progress.connect(self.progress_bar.setValue)
        self.downloader.finished.connect(self.download_finished)
        self.downloader.error.connect(self.download_error)
        self.downloader.start()
        
    def download_finished(self, temp_file):
        """Callback quando download termina"""
        self.temp_file = temp_file
        
        msg = QMessageBox(self)
        msg.setWindowTitle("Download Concluído")
        msg.setText("Atualização baixada com sucesso!")
        msg.setInformativeText("O aplicativo será reiniciado para aplicar a atualização.")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()
        
        self.apply_update()
        
    def download_error(self, error_msg):
        """Callback quando há erro no download"""
        QMessageBox.critical(self, "Erro no Download", f"Falha ao baixar atualização:\n{error_msg}")
        self.download_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
    def apply_update(self):
        """Aplica a atualização"""
        if not self.temp_file:
            return
            
        try:
            # Caminho do executável atual
            current_exe = sys.executable if hasattr(sys, 'frozen') else __file__
            current_dir = os.path.dirname(current_exe)
            backup_exe = os.path.join(current_dir, "ExcelCalculator_backup.exe")
            
            # Cria script de atualização
            update_script = self.create_update_script(current_exe, self.temp_file, backup_exe)
            
            # Executa script e fecha aplicativo
            subprocess.Popen([sys.executable, update_script], 
                           creationflags=subprocess.CREATE_NEW_CONSOLE if os.name == 'nt' else 0)
            
            # Fecha o aplicativo atual
            sys.exit(0)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro na Atualização", f"Falha ao aplicar atualização:\n{str(e)}")
            
    def create_update_script(self, current_exe, new_exe, backup_exe):
        """Cria script Python para aplicar a atualização"""
        script_content = f'''
import os
import sys
import time
import shutil
import subprocess

def update():
    try:
        # Aguarda o aplicativo fechar
        time.sleep(2)
        
        # Faz backup do executável atual
        if os.path.exists("{current_exe}"):
            if os.path.exists("{backup_exe}"):
                os.remove("{backup_exe}")
            shutil.move("{current_exe}", "{backup_exe}")
        
        # Copia novo executável
        shutil.move("{new_exe}", "{current_exe}")
        
        # Reinicia o aplicativo
        subprocess.Popen(["{current_exe}"])
        
        print("Atualização aplicada com sucesso!")
        
    except Exception as e:
        print(f"Erro na atualização: {{e}}")
        # Restaura backup se houver erro
        if os.path.exists("{backup_exe}"):
            shutil.move("{backup_exe}", "{current_exe}")
        input("Pressione Enter para continuar...")

if __name__ == "__main__":
    update()
'''
        
        script_path = os.path.join(tempfile.gettempdir(), "excel_calculator_update.py")
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        return script_path

class AutoUpdater:
    """Gerenciador principal de atualizações"""
    
    def __init__(self, parent_window=None):
        self.parent_window = parent_window
        self.checker = None
        
    def check_for_updates(self, silent=False):
        """Verifica atualizações (silent=True não mostra mensagens se não houver)"""
        if self.checker and self.checker.isRunning():
            return
            
        self.checker = UpdateChecker()
        self.checker.update_available.connect(lambda info: self.show_update_dialog(info))
        
        if not silent:
            self.checker.no_update.connect(lambda: QMessageBox.information(
                self.parent_window, 
                "Verificar Atualizações", 
                "Você já está usando a versão mais recente!"
            ))
            self.checker.error.connect(lambda err: QMessageBox.warning(
                self.parent_window,
                "Erro na Verificação",
                f"Não foi possível verificar atualizações:\n{err}"
            ))
        
        self.checker.start()
        
    def show_update_dialog(self, update_info):
        """Mostra dialog de atualização"""
        dialog = UpdateDialog(update_info, self.parent_window)
        dialog.exec()
        
    def get_current_version(self):
        """Retorna versão atual"""
        return CURRENT_VERSION

# Função para verificação automática no startup
def check_updates_on_startup(parent_window=None):
    """Verifica atualizações automaticamente no início"""
    updater = AutoUpdater(parent_window)
    updater.check_for_updates(silent=True)
    return updater