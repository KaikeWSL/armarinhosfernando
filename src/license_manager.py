import hashlib
import platform
import uuid
import subprocess
import psycopg2
import pyperclip
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QPixmap, QPainter, QColor

class LicenseManager:
    def __init__(self):
        self.db_url = "postgresql://neondb_owner:npg_Chj1aBTeA4pR@ep-quiet-forest-acam7hat-pooler.sa-east-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require"
        self.table_name = "armarinhos_fernando"
        
    def get_machine_id(self):
        """Gera um ID único baseado no hardware da máquina"""
        try:
            # Combina várias informações do sistema para criar ID único
            machine_info = []
            
            # Nome da máquina
            machine_info.append(platform.node())
            
            # Processador
            machine_info.append(platform.processor())
            
            # Sistema operacional
            machine_info.append(platform.system() + platform.release())
            
            # MAC Address da placa de rede
            try:
                mac = ':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff) for ele in range(0,8*6,8)][::-1])
                machine_info.append(mac)
            except:
                pass
            
            # UUID da máquina (Windows)
            try:
                if platform.system() == "Windows":
                    result = subprocess.check_output("wmic csproduct get uuid", shell=True)
                    uuid_line = result.decode().split('\n')[1].strip()
                    machine_info.append(uuid_line)
            except:
                pass
            
            # Combina tudo e gera hash
            combined = "|".join(machine_info)
            machine_hash = hashlib.sha256(combined.encode()).hexdigest()
            
            # Retorna os primeiros 10 caracteres para chave mais amigável
            return machine_hash[:10].upper()
            
        except Exception as e:
            # Fallback: usa informações básicas
            fallback = platform.node() + platform.system()
            return hashlib.md5(fallback.encode()).hexdigest()[:10].upper()
    
    def check_license(self):
        """Verifica se a licença é válida no banco de dados"""
        try:
            machine_id = self.get_machine_id()
            
            # Conecta ao banco
            conn = psycopg2.connect(self.db_url)
            cursor = conn.cursor()
            
            # Verifica se a chave existe e está liberada
            query = f"SELECT liberado FROM {self.table_name} WHERE chave = %s"
            cursor.execute(query, (machine_id,))
            result = cursor.fetchone()
            
            cursor.close()
            conn.close()
            
            if result and result[0] == 'sim':
                return True, machine_id
            else:
                return False, machine_id
                
        except Exception as e:
            print(f"Erro ao verificar licença: {e}")
            return False, self.get_machine_id()
    
    def show_license_dialog(self, machine_id):
        """Mostra dialog de licença bloqueada"""
        dialog = QDialog()
        dialog.setWindowTitle("FALHA AO VALIDAR LICENÇA!")
        dialog.setModal(True)
        dialog.resize(450, 300)
        dialog.setStyleSheet("""
            QDialog {
                background-color: #2c3e50;
                color: white;
            }
            QLabel {
                color: white;
                padding: 5px;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        
        layout = QVBoxLayout()
        
        # Título
        title_label = QLabel("FALHA AO VALIDAR LICENÇA!")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(14)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # Subtítulo
        subtitle_label = QLabel("Caso for a primeira execução informe a chave:")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle_label)
        
        # Chave
        key_label = QLabel(machine_id)
        key_font = QFont()
        key_font.setBold(True)
        key_font.setPointSize(16)
        key_label.setFont(key_font)
        key_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        key_label.setStyleSheet("background-color: #34495e; padding: 10px; border-radius: 5px; margin: 10px;")
        layout.addWidget(key_label)
        
        # Contatos
        contact1_label = QLabel("Para: Kaike Wesley\nWhatsapp: (11) 98710-8126\ne-mail: kaike.wesley12@gmail.com")
        contact1_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(contact1_label)
        
        ou_label = QLabel("ou")
        ou_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ou_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(ou_label)
        
        contact2_label = QLabel("Para: Felipe Richter\nWhatsapp: (11) 99691-6400\ne-mail: feliperichter_9@hotmail.com")
        contact2_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(contact2_label)
        
        # Mensagem sobre área de transferência
        clipboard_label = QLabel("chave salva na área de transferência")
        clipboard_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        clipboard_label.setStyleSheet("font-style: italic; color: #bdc3c7;")
        layout.addWidget(clipboard_label)
        
        # Botão OK
        ok_button = QPushButton("Ok")
        ok_button.clicked.connect(dialog.accept)
        layout.addWidget(ok_button)
        
        dialog.setLayout(layout)
        
        # Copia chave para área de transferência
        try:
            pyperclip.copy(machine_id)
        except:
            pass  # Se não conseguir copiar, não faz nada
        
        return dialog.exec()

def check_license_and_run():
    """Função principal para verificar licença antes de executar o app"""
    license_manager = LicenseManager()
    is_valid, machine_id = license_manager.check_license()
    
    if not is_valid:
        # Cria uma aplicação mínima apenas para mostrar o dialog
        from PyQt6.QtWidgets import QApplication
        import sys
        
        app = QApplication(sys.argv)
        license_manager.show_license_dialog(machine_id)
        app.quit()
        return False
    
    return True