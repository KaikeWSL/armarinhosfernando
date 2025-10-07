"""
Interface gráfica principal do Sistema de Cál        self.setGeometry(100, 100, 1400, 900)  # Janela maior para melhor visualizaçãoulo Excel
Design moderno dark theme com logo da Armarinhos Fernando
Desenvolvido com PyQt6 para máxima compatibilidade e performance
"""

import sys
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QFileDialog, QMessageBox,
                             QTableWidget, QTableWidgetItem, QSpinBox, QDoubleSpinBox,
                             QGroupBox, QProgressBar, QTextEdit, QSplitter, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QIcon, QPalette, QColor, QPixmap, QPainter, QLinearGradient
import os
import base64
from assets import COLORS, GRADIENTS
from dark_theme import get_dark_theme_stylesheet

class CalculationThread(QThread):
    """Thread para executar cálculos sem travar a interface"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, processor, percentage):
        super().__init__()
        self.processor = processor
        self.percentage = percentage
        
    def run(self):
        try:
            success = self.processor.calculate_suggestions(self.percentage, self.progress)
            self.finished.emit(success, "Cálculo concluído com sucesso!")
        except Exception as e:
            self.finished.emit(False, f"Erro durante o cálculo: {str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_processor = None
        self.calculation_thread = None
        self.animation_timer = QTimer()
        self.animation_step = 0
        self.init_ui()
        self.apply_dark_theme()
        self.start_logo_animation()
        
    def init_ui(self):
        """Inicializa a interface do usuário"""
        self.setWindowTitle("Sistema de Transposição Excel - Armarinhos Fernando")
        self.setGeometry(100, 100, 1400, 900)
        self.setMinimumSize(1200, 800)
        
        # Cria menu bar
        self.create_menu_bar()
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Banner superior com logo
        header_widget = self.create_header_banner()
        main_layout.addWidget(header_widget)
        
        # Container principal com padding
        content_widget = QWidget()
        content_widget.setObjectName("contentWidget")
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(15)  # Reduzido de 20 para 15
        content_layout.setContentsMargins(15, 15, 15, 15)  # Reduzido de 20 para 15
        
        # Área de seleção de arquivo
        file_group = self.create_modern_file_selection()
        content_layout.addWidget(file_group)
        
        # Parâmetros e controles na mesma linha
        params_controls_layout = QHBoxLayout()
        params_controls_layout.setSpacing(15)
        
        # Área de parâmetros
        params_group = self.create_modern_parameters()
        params_controls_layout.addWidget(params_group, 1)  # 1 parte
        
        # Área de controles
        controls_group = self.create_modern_controls()
        params_controls_layout.addWidget(controls_group, 2)  # 2 partes
        
        # Container para parâmetros e controles
        params_controls_widget = QWidget()
        params_controls_widget.setLayout(params_controls_layout)
        content_layout.addWidget(params_controls_widget)
        
        # Área de visualização dos dados
        data_group = self.create_modern_data_visualization()
        content_layout.addWidget(data_group)
        
        main_layout.addWidget(content_widget)
        
        # Status bar moderno
        self.statusBar().setObjectName("modernStatusBar")
        self.statusBar().showMessage("🚀 Sistema pronto para usar • Selecione um arquivo Excel")
        
    def create_header_banner(self):
        """Cria o banner superior com logo da Armarinhos Fernando"""
        header = QFrame()
        header.setObjectName("headerBanner")
        header.setFixedHeight(80)  # Reduzido de 100 para 80
        
        layout = QHBoxLayout(header)
        layout.setContentsMargins(30, 15, 30, 15)
        
        # Logo da Armarinhos Fernando
        logo_label = QLabel()
        logo_label.setObjectName("logoLabel")
        logo_pixmap = self.create_logo_pixmap()
        logo_label.setPixmap(logo_pixmap)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Título principal
        title_container = QWidget()
        title_layout = QVBoxLayout(title_container)
        title_layout.setSpacing(5)
        
        title_label = QLabel("Sistema Excel Profissional")
        title_label.setObjectName("mainTitle")
        
        subtitle_label = QLabel("Calculadora Avançada de Sugestões • Armarinhos Fernando")
        subtitle_label.setObjectName("subtitle")
        
        title_layout.addWidget(title_label)
        title_layout.addWidget(subtitle_label)
        
        # Versão e status
        version_label = QLabel("v2.0 PREMIUM")
        version_label.setObjectName("versionLabel")
        version_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        
        layout.addWidget(logo_label)
        layout.addWidget(title_container, 1)
        layout.addWidget(version_label)
        
        return header
        
    def create_logo_pixmap(self):
        """Cria o logo da Armarinhos Fernando"""
        # Cria um pixmap de 200x70 com o logo
        pixmap = QPixmap(200, 70)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Fundo com gradiente
        gradient = QLinearGradient(0, 0, 200, 0)
        gradient.setColorAt(0, QColor(COLORS['primary']))
        gradient.setColorAt(1, QColor(COLORS['secondary']))
        
        painter.setBrush(gradient)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(0, 0, 200, 70, 10, 10)
        
        # Texto "AF"
        painter.setPen(QColor(COLORS['text_primary']))
        font = QFont("Arial", 24, QFont.Weight.Bold)
        painter.setFont(font)
        painter.drawText(15, 45, "AF")
        
        # Texto "Armarinhos Fernando"
        font = QFont("Arial", 12, QFont.Weight.Bold)
        painter.setFont(font)
        painter.drawText(60, 30, "Armarinhos")
        painter.drawText(60, 50, "Fernando")
        
        painter.end()
        return pixmap
        
    def start_logo_animation(self):
        """Inicia animação sutil do logo"""
        self.animation_timer.timeout.connect(self.animate_logo)
        self.animation_timer.start(50)  # 50ms
        
    def animate_logo(self):
        """Animação sutil do logo"""
        self.animation_step += 1
        if self.animation_step > 100:
            self.animation_timer.stop()
            
    def create_modern_file_selection(self):
        """Cria área moderna de seleção de arquivo"""
        group = QGroupBox("📁 Seleção de Arquivo Excel")
        group.setObjectName("modernGroup")
        
        layout = QHBoxLayout(group)
        layout.setSpacing(15)
        
        # Container do arquivo
        file_container = QFrame()
        file_container.setObjectName("fileContainer")
        file_layout = QHBoxLayout(file_container)
        
        self.file_label = QLabel("Nenhum arquivo selecionado")
        self.file_label.setObjectName("fileLabel")
        
        file_layout.addWidget(self.file_label, 1)
        
        self.select_file_btn = QPushButton("🔍 Selecionar Arquivo")
        self.select_file_btn.setObjectName("primaryButton")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        
        layout.addWidget(file_container, 1)
        layout.addWidget(self.select_file_btn)
        
        return group
        
    def create_modern_parameters(self):
        """Cria área moderna de parâmetros"""
        group = QGroupBox("⚙️ Parâmetros de Cálculo")
        group.setObjectName("modernGroup")
        
        layout = QHBoxLayout(group)
        layout.setSpacing(20)
        
        # Container da porcentagem
        percent_container = QFrame()
        percent_container.setObjectName("paramContainer")
        percent_layout = QHBoxLayout(percent_container)
        
        percent_label = QLabel("💹 Porcentagem de Ajuste:")
        percent_label.setObjectName("paramLabel")
        
        self.percentage_spin = QDoubleSpinBox()
        self.percentage_spin.setObjectName("modernSpinBox")
        self.percentage_spin.setRange(-100.0, 1000.0)
        self.percentage_spin.setValue(10.0)
        self.percentage_spin.setSuffix("%")
        self.percentage_spin.setDecimals(2)
        self.percentage_spin.setMinimumWidth(120)
        
        percent_layout.addWidget(percent_label)
        percent_layout.addWidget(self.percentage_spin)
        percent_layout.addStretch()
        
        layout.addWidget(percent_container)
        layout.addStretch()
        
        return group
        
    def create_modern_controls(self):
        """Cria área moderna de controles"""
        group = QGroupBox("🎮 Controles")
        group.setObjectName("modernGroup")
        
        layout = QHBoxLayout(group)
        layout.setSpacing(15)
        
        self.calculate_btn = QPushButton("🚀 Calcular Sugestões")
        self.calculate_btn.setObjectName("successButton")
        self.calculate_btn.clicked.connect(self.calculate_suggestions)
        self.calculate_btn.setEnabled(False)
        self.calculate_btn.setMinimumHeight(45)
        
        self.export_btn = QPushButton("💾 Exportar Resultados")
        self.export_btn.setObjectName("warningButton")
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        self.export_btn.setMinimumHeight(45)
        
        # Botão de teste de cores
    

        
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("modernProgressBar")
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(8)
        
        layout.addWidget(self.calculate_btn)
        layout.addWidget(self.export_btn)

        layout.addStretch()
        layout.addWidget(self.progress_bar)
        
        return group
        
    def create_modern_data_visualization(self):
        """Cria área moderna de visualização de dados"""
        group = QGroupBox("📊 Visualização dos Dados")
        group.setObjectName("modernGroup")
        
        layout = QVBoxLayout(group)
        
        self.data_table = QTableWidget()
        self.data_table.setObjectName("modernTable")
        self.data_table.setAlternatingRowColors(False)  # Desabilitado para usar cores customizadas
        self.data_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        
        # Configura estilo básico da tabela (sem background dos itens)
        from assets import COLORS
        self.data_table.setStyleSheet(f"""
            QTableWidget {{
                gridline-color: {COLORS['border']};
                border: 2px solid {COLORS['border']};
                border-radius: 8px;
                font-size: 11px;
            }}
            QHeaderView::section {{
                background-color: {COLORS['table_header']};
                color: {COLORS['text_primary']};
                padding: 12px;
                border: none;
                font-weight: bold;
                font-size: 12px;
            }}
        """)
        
        layout.addWidget(self.data_table)
        
        return group
        
    def create_header(self):
        """Cria o cabeçalho da aplicação"""
        layout = QHBoxLayout()
        
        title_label = QLabel("Sistema de Cálculo Excel Profissional")
        title_font = QFont("Arial", 16, QFont.Weight.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50; margin: 10px;")
        
        layout.addWidget(title_label)
        layout.addStretch()
        
        return layout
        
    def create_file_selection_group(self):
        """Cria o grupo de seleção de arquivo"""
        group = QGroupBox("Seleção de Arquivo Excel")
        layout = QHBoxLayout(group)
        
        self.file_label = QLabel("Nenhum arquivo selecionado")
        self.file_label.setStyleSheet("padding: 5px; background-color: #ecf0f1; border-radius: 3px;")
        
        self.select_file_btn = QPushButton("Selecionar Arquivo Excel")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        
        layout.addWidget(self.file_label, 1)
        layout.addWidget(self.select_file_btn)
        
        return group
        
    def create_parameters_group(self):
        """Cria o grupo de parâmetros"""
        group = QGroupBox("Parâmetros de Cálculo")
        layout = QHBoxLayout(group)
        
        # Porcentagem de ajuste
        layout.addWidget(QLabel("Porcentagem de Ajuste:"))
        self.percentage_spin = QDoubleSpinBox()
        self.percentage_spin.setRange(-100.0, 1000.0)
        self.percentage_spin.setValue(10.0)
        self.percentage_spin.setSuffix("%")
        self.percentage_spin.setDecimals(2)
        layout.addWidget(self.percentage_spin)
        
        layout.addStretch()
        
        return group
        
    def create_controls_group(self):
        """Cria o grupo de controles"""
        group = QGroupBox("Controles")
        layout = QHBoxLayout(group)
        
        self.calculate_btn = QPushButton("🔄 Calcular Sugestões")
        self.calculate_btn.setObjectName("primaryButton")
        self.calculate_btn.clicked.connect(self.calculate_suggestions)
        self.calculate_btn.setEnabled(False)
        
        self.export_btn = QPushButton("💾 Exportar Formato Antigo")
        self.export_btn.setObjectName("successButton")
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        layout.addWidget(self.calculate_btn)
        layout.addWidget(self.export_btn)
        layout.addStretch()
        layout.addWidget(self.progress_bar)
        
        return group
        
    def create_data_visualization_group(self):
        """Cria o grupo de visualização de dados"""
        group = QGroupBox("Visualização dos Dados")
        layout = QVBoxLayout(group)
        
        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        layout.addWidget(self.data_table)
        
        return group
        
    def create_log_group(self):
        """Cria o grupo de log"""
        group = QGroupBox("Log de Operações")
        layout = QVBoxLayout(group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)  # Reduzido para dar mais espaço à tabela
        layout.addWidget(self.log_text)
        
        return group
        
    def apply_dark_theme(self):
        """Aplica o tema dark moderno"""
        self.setStyleSheet(get_dark_theme_stylesheet())
        
    def select_excel_file(self):
        """Seleciona arquivo Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Selecionar Arquivo Excel", 
            "", 
            "Arquivos Excel (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # Import do processador corrigido
                from excel_processor import ExcelProcessor
                
                self.excel_processor = ExcelProcessor(file_path)
                success = self.excel_processor.process_file()
                
                if success:
                    self.file_label.setText(os.path.basename(file_path))
                    self.calculate_btn.setEnabled(True)
                    self.load_data_preview()
                    # Arquivo carregado e processado
                    self.statusBar().showMessage(f"Arquivo processado: {os.path.basename(file_path)}")
                else:
                    raise Exception("Falha no processamento do arquivo")
                
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao carregar arquivo: {str(e)}")
                # Erro ao carregar arquivo
                
    def load_data_preview(self):
        """Carrega preview dos dados na tabela com totais"""
        if not self.excel_processor:
            return
            
        try:
            df = self.excel_processor.get_preview_data()
            
            if df.empty:
                # AVISO: Nenhum dado encontrado no arquivo
                # Cria uma tabela com mensagem de aviso
                self.data_table.setRowCount(1)
                self.data_table.setColumnCount(1)
                self.data_table.setHorizontalHeaderLabels(["Aviso"])
                
                item = QTableWidgetItem("Nenhum dado processado. Verifique se o arquivo tem o formato correto.")
                self.data_table.setItem(0, 0, item)
                return
            
            # Usa a nova função que inclui totais
            self.load_data_to_table(df)
            
        except Exception as e:
            QMessageBox.warning(self, "Aviso", f"Erro ao carregar preview: {str(e)}")
            import traceback
            traceback.print_exc()
            
    def apply_table_colors(self, df):
        """Aplica cores nas linhas da tabela - Sistema completamente novo"""
        if df.empty:
            return


        
        # Primeiro define as configurações de cores
        color_config = {
            'vendas': {'background': "#FFF09C", 'text': '#000000', 'name': 'VENDAS'},      # Amarelo ouro
            'sugestao': {'background': "#9AFF9A", 'text': '#000000', 'name': 'SUGESTÃO'},  # Verde lima
            'entradas': {'background': "#FFFFFF", 'text': "#000000", 'name': 'ENTRADAS'},  # Azul royal
            'default': {'background': '#2C2C2C', 'text': '#FFFFFF', 'name': 'PADRÃO'}     # Cinza escuro
        }
        
        # Primeiro aplica o tema básico da tabela
        self._apply_table_theme()
        
        # Conta quantas linhas de cada tipo foram encontradas
        contadores = {'vendas': 0, 'sugestao': 0, 'entradas': 0, 'default': 0, 
                     'total_vendas': 0, 'total_sugestao': 0, 'total_entradas': 0}
        
        # Processa cada linha DEPOIS do tema básico
        for row in range(len(df)):
            linha_tipo = self._get_row_type(df, row)
            cor_aplicada = self._apply_row_color(row, linha_tipo, color_config)
            if cor_aplicada in contadores:
                contadores[cor_aplicada] += 1
            else:
                contadores['default'] += 1
        
        # Exibe relatório final

    
    def _get_row_type(self, df, row):
        """Identifica o tipo da linha baseado no conteúdo"""
        try:
            # Tenta primeiro a coluna 'Tipo'
            if 'Tipo' in df.columns:
                valor = str(df.iloc[row]['Tipo']).lower().strip()

                return valor
            
            # Se não houver coluna 'Tipo', usa a primeira coluna
            if len(df.columns) > 0:
                valor = str(df.iloc[row, 0]).lower().strip()

                return valor
                
        except Exception as e:
            print(f"❌ Erro ao ler linha {row}: {e}")
        
        return ""
    
    def _apply_row_color(self, row, linha_tipo, color_config):
        """Aplica a cor apropriada na linha baseada no tipo"""
        # Verifica se é uma linha de total especial
        if 'quantidade total' in linha_tipo:
            if 'vendida' in linha_tipo:
                config = {'background': '#FFA500', 'text': '#000000', 'name': 'TOTAL VENDAS'}  # Laranja
                tipo_final = 'total_vendas'
            elif 'solicitada' in linha_tipo:
                config = {'background': '#90EE90', 'text': '#000000', 'name': 'TOTAL SUGESTÃO'}  # Verde claro
                tipo_final = 'total_sugestao'
            elif 'entrada' in linha_tipo:
                config = {'background': '#87CEEB', 'text': '#000000', 'name': 'TOTAL ENTRADAS'}  # Azul claro
                tipo_final = 'total_entradas'
            else:
                config = color_config['default']
                tipo_final = 'default'
        else:
            # Lógica normal para linhas de dados
            if any(palavra in linha_tipo for palavra in ['venda', 'vendas']):
                config = color_config['vendas']
                tipo_final = 'vendas'
            elif any(palavra in linha_tipo for palavra in ['sugest', 'sugestão', 'sugestao']):
                config = color_config['sugestao'] 
                tipo_final = 'sugestao'
            elif any(palavra in linha_tipo for palavra in ['entrada', 'entradas']):
                config = color_config['entradas']
                tipo_final = 'entradas'
            else:
                config = color_config['default']
                tipo_final = 'default'
        
        # Aplica a cor de forma mais agressiva
        cor_fundo = QColor(config['background'])
        cor_texto = QColor(config['text'])
        
        # Força aplicação da cor em cada célula individualmente
        for col in range(self.data_table.columnCount()):
            item = self.data_table.item(row, col)
            if item is not None:
                # Método 1: setBackground/setForeground
                item.setBackground(cor_fundo)
                item.setForeground(cor_texto)
                
                # Método 2: Força cor usando flags do item
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEnabled)
                
                # Método 3: Tenta diferentes roles
                item.setData(Qt.ItemDataRole.BackgroundRole, cor_fundo)
                item.setData(Qt.ItemDataRole.ForegroundRole, cor_texto)
        
        if tipo_final != 'default':
            print(f"✅ Linha {row} = {config['name']} ({config['background']})")
            
            # Força redesenho da linha específica
            for col in range(self.data_table.columnCount()):
                self.data_table.update(self.data_table.model().index(row, col))
        
        return tipo_final
    
    def _apply_table_theme(self):
        """Aplica o tema visual da tabela sem interferir nas cores das células"""
        # CSS básico APENAS para estrutura da tabela - SEM cores de células
        base_style = """
        QTableWidget {
            gridline-color: #4169E1;
            font-size: 11px;
            border: 2px solid #4169E1;
            border-radius: 6px;
            outline: none;
        }
        
        QHeaderView::section {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                        stop:0 #4169E1, stop:1 #2C5AA0);
            color: #FFFFFF;
            padding: 10px;
            border: none;
            font-weight: bold;
            font-size: 12px;
            text-align: center;
        }
        """
        
        # Primeiro aplica só o estilo base
        self.data_table.setStyleSheet(base_style)
        
        # Força atualização ANTES de aplicar cores específicas
        self.data_table.clearSelection()
        self.data_table.viewport().update()
        QApplication.processEvents()
        
    def test_colors_with_sample_data(self):
        """Função de teste para verificar as cores com dados de exemplo"""
        import pandas as pd
        
        # Cria dados de teste com TODOS os tipos de linha para demonstração
        test_data = {
            'Tipo': [
                'VENDAS',           # Linha amarela
                'ENTRADAS',         # Linha azul
                'SUGESTÃO',         # Linha verde
                'VENDAS',           # Linha amarela
                'SUGESTAO',         # Linha verde (sem acento)
                'ENTRADAS',         # Linha azul
                'PRODUTO VENDAS',   # Linha amarela (contém "vendas")
                'ENTRADA LOJA',     # Linha azul (contém "entrada")
                'SUGEST COMPRA',    # Linha verde (contém "sugest")
                'ITEM NORMAL'       # Linha padrão (cinza)
            ],
            'Família': ['DEMO'] * 10,
            'Código': [f'TEST{i:03d}' for i in range(1, 11)],
            'Descrição': [
                'Produto para Vendas - Amarelo',
                'Produto de Entradas - Azul',
                'Produto Sugestão - Verde',  
                'Outro Produto Vendas - Amarelo',
                'Outro Produto Sugestao - Verde',
                'Mais Entradas - Azul',
                'Vendas Especiais - Amarelo',
                'Entrada Especial - Azul', 
                'Sugestão Especial - Verde',
                'Produto Comum - Padrão'
            ],
            'Cx c/': [10, 20, 15, 12, 18, 25, 14, 22, 16, 30],
            'Secundario': [1, 2, 1, 1, 2, 3, 1, 2, 1, 3],
            'Saldo Local': [100, 200, 150, 120, 180, 250, 140, 220, 160, 300],
            'Invent': [50, 100, 75, 60, 90, 125, 70, 110, 80, 150],
            'Devol.': [5, 10, 8, 6, 9, 12, 7, 11, 8, 15],
            'Dep25': [0] * 10,
            'Entradas': [300, 400, 350, 320, 380, 450, 340, 420, 360, 500]
        }
        
        df_test = pd.DataFrame(test_data)
        

        
        # Carrega os dados na tabela
        self.load_data_to_table(df_test)
        
        # Aplica as cores
        self.apply_table_colors(df_test)
        
        # Atualiza o status
        self.statusBar().showMessage("🎨 Teste de cores aplicado! Verifique as linhas coloridas na tabela")
        

        
    def load_data_to_table(self, df):
        """Carrega dados na tabela e adiciona linhas de totais"""
        if df.empty:
            return
            
        # Calcula os totais por tipo
        df_with_totals = self._add_summary_rows(df)
        
        # Configura a tabela
        self.data_table.setRowCount(len(df_with_totals))
        self.data_table.setColumnCount(len(df_with_totals.columns))
        
        # Define headers
        headers = [str(col) for col in df_with_totals.columns]
        self.data_table.setHorizontalHeaderLabels(headers)
        
        # Preenche os dados
        for i in range(len(df_with_totals)):
            for j, value in enumerate(df_with_totals.iloc[i]):
                if pd.isna(value):
                    display_value = ""
                elif isinstance(value, (int, float)):
                    display_value = str(int(value)) if value == int(value) else f"{value:.2f}"
                else:
                    display_value = str(value)
                    
                item = QTableWidgetItem(display_value)
                self.data_table.setItem(i, j, item)
        
        # Aplica as cores (incluindo cores especiais para totais)
        self.apply_table_colors(df_with_totals)
        self.data_table.resizeColumnsToContents()
        
    def _add_summary_rows(self, df):
        """Adiciona linhas de resumo com totais por tipo e coluna de total por linha"""
        # Cria uma cópia do DataFrame original
        df_result = df.copy()
        
        # Identifica colunas das lojas (numéricas, excluindo colunas administrativas)
        lojas_cols = []
        other_numeric_cols = []
        
        for col in df.columns:
            if df[col].dtype in ['int64', 'float64']:
                # Se a coluna não é uma das colunas administrativas, é uma loja
                if col not in ['Cx c/', 'Secundario', 'Saldo Local', 'Invent', 'Devol.', 'Dep25']:
                    lojas_cols.append(col)
                else:
                    other_numeric_cols.append(col)
            elif col in ['Cx c/', 'Secundario', 'Saldo Local', 'Invent', 'Devol.', 'Dep25', 'Entradas']:
                other_numeric_cols.append(col)
        
        # Todas as colunas numéricas (para cálculo do Total da linha)
        all_numeric_cols = lojas_cols + other_numeric_cols
        
        # Adiciona coluna de "Total" que soma TODAS as colunas numéricas por linha (para visualização)
        df_result['Total'] = 0.0
        for i in range(len(df_result)):
            total_linha = 0.0
            for col in all_numeric_cols:
                if col in df_result.columns:
                    valor = df_result.iloc[i][col]
                    if pd.notna(valor) and isinstance(valor, (int, float)):
                        try:
                            total_linha += float(valor)
                        except (ValueError, TypeError):
                            # Se não conseguir converter, ignora o valor
                            continue
            # Arredonda para 2 casas decimais e converte para inteiro se for número inteiro
            total_linha = round(total_linha, 2)
            if total_linha.is_integer():
                df_result.loc[i, 'Total'] = int(total_linha)
            else:
                df_result.loc[i, 'Total'] = total_linha
        
        # Para os totais das linhas de resumo, usa apenas as colunas das lojas
        totals_cols = lojas_cols + ['Total']  # Inclui Total para os cálculos
        
        # Calcula totais por tipo
        totals = {
            'vendas': {'rows': [], 'name': 'Quantidade Total Vendida'},
            'sugestao': {'rows': [], 'name': 'Quantidade Total Solicitada'}, 
            'entradas': {'rows': [], 'name': 'Quantidade Total Entrada'}
        }
        
        # Classifica cada linha por tipo
        for i in range(len(df)):
            linha_tipo = self._get_row_type_for_totals(df, i)
            if linha_tipo in totals:
                totals[linha_tipo]['rows'].append(i)
        
        # Adiciona linha em branco antes dos totais
        empty_row = {}
        for col in df_result.columns:
            empty_row[col] = ""
        df_result = pd.concat([df_result, pd.DataFrame([empty_row])], ignore_index=True)
        
        # Adiciona linhas de totais
        for tipo_key, tipo_data in totals.items():
            if tipo_data['rows']:  # Se há linhas deste tipo
                total_row = {}
                
                # Primeira coluna com o nome do total
                first_col = df.columns[0]
                total_row[first_col] = tipo_data['name']
                
                # Outras colunas de texto ficam vazias
                for col in df_result.columns:
                    if col not in totals_cols and col != first_col:
                        total_row[col] = ""
                
                # Calcula soma APENAS das colunas das lojas (não inclui Cx c/, Secundario, etc.)
                for col in totals_cols:
                    if col in df.columns and col in lojas_cols:
                        # Para colunas das lojas, soma apenas essas colunas
                        try:
                            soma = df.iloc[tipo_data['rows']][col].sum()
                            # Converte para float primeiro, depois para int se for número inteiro
                            soma_float = float(soma)
                            total_row[col] = int(soma_float) if soma_float.is_integer() and soma_float != 0 else (soma_float if soma_float != 0 else 0)
                        except (ValueError, TypeError, OverflowError):
                            total_row[col] = 0
                    elif col == 'Total':
                        # Para a coluna Total, soma APENAS as colunas das lojas
                        total_soma = 0.0
                        for c in lojas_cols:
                            if c in df.columns:
                                try:
                                    valor_col = df.iloc[tipo_data['rows']][c].sum()
                                    total_soma += float(valor_col)
                                except (ValueError, TypeError, OverflowError):
                                    continue
                        # Converte para int se for número inteiro
                        total_row[col] = int(total_soma) if total_soma.is_integer() else round(total_soma, 2)
                
                # Adiciona a linha de total
                df_result = pd.concat([df_result, pd.DataFrame([total_row])], ignore_index=True)
        
        return df_result
    
    def _get_row_type_for_totals(self, df, row):
        """Identifica o tipo da linha para cálculo de totais"""
        try:
            if 'Tipo' in df.columns:
                valor = str(df.iloc[row]['Tipo']).lower().strip()
            elif len(df.columns) > 0:
                valor = str(df.iloc[row, 0]).lower().strip()
            else:
                return 'default'
                
            if any(palavra in valor for palavra in ['venda', 'vendas']):
                return 'vendas'
            elif any(palavra in valor for palavra in ['sugest', 'sugestão', 'sugestao']):
                return 'sugestao'
            elif any(palavra in valor for palavra in ['entrada', 'entradas']):
                return 'entradas'
                
        except:
            pass
            
        return 'default'
            
    def calculate_suggestions(self):
        """Inicia o cálculo das sugestões (ajuste percentual)"""
        if not self.excel_processor:
            return
            
        percentage = self.percentage_spin.value()
        
        try:
            # Desabilita controles durante o cálculo
            self.calculate_btn.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(50)
            
            # Aplica ajuste percentual
            self.excel_processor.apply_percentage_adjustment(percentage)
            
            self.progress_bar.setValue(100)
            
            # Recarrega preview
            self.load_data_preview()
            
            # Habilita exportação
            self.export_btn.setEnabled(True)
            
            QMessageBox.information(self, "Sucesso", f"Sugestões calculadas com {percentage}% de ajuste!")
            # Sugestões calculadas
            self.statusBar().showMessage("Cálculo concluído - Pronto para exportar")
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro no cálculo: {str(e)}")
            # Erro no cálculo
            
        finally:
            self.progress_bar.setVisible(False)
            self.calculate_btn.setEnabled(True)
        
    def on_calculation_finished(self, success, message):
        """Callback quando o cálculo termina"""
        self.progress_bar.setVisible(False)
        self.calculate_btn.setEnabled(True)
        
        if success:
            self.export_btn.setEnabled(True)
            self.load_data_preview()  # Recarrega os dados atualizados
            QMessageBox.information(self, "Sucesso", message)
            self.statusBar().showMessage("Cálculo concluído com sucesso")
        else:
            QMessageBox.critical(self, "Erro", message)
            self.statusBar().showMessage("Erro no cálculo")
            
        # Mensagem de progresso
        
    def export_results(self):
        """Exporta os resultados no formato antigo"""
        if not self.excel_processor:
            return
        
        # Obter famílias disponíveis
        families = self.excel_processor.get_available_families()
        
        if not families:
            QMessageBox.warning(self, "Aviso", "Nenhuma família encontrada nos dados.")
            return
        
        # Dialog para seleção de família
        from PyQt6.QtWidgets import QDialog, QVBoxLayout, QComboBox, QLabel, QPushButton, QHBoxLayout, QRadioButton, QButtonGroup
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Exportar Resultados")
        dialog.setModal(True)
        dialog.resize(400, 200)
        
        layout = QVBoxLayout()
        
        # Opções de exportação
        layout.addWidget(QLabel("Escolha o tipo de exportação:"))
        
        export_group = QButtonGroup()
        
        # Opção 1: Arquivo único
        single_radio = QRadioButton("Arquivo único com família específica")
        single_radio.setChecked(True)
        export_group.addButton(single_radio, 1)
        layout.addWidget(single_radio)
        
        # Combo para família (habilitado apenas se arquivo único for selecionado)
        family_combo = QComboBox()
        family_combo.addItem("Todas as famílias", "Todas")
        for family in families:
            family_combo.addItem(str(family), family)
        layout.addWidget(family_combo)
        
        # Opção 2: Arquivos separados
        multiple_radio = QRadioButton("Arquivos separados por família")
        export_group.addButton(multiple_radio, 2)
        layout.addWidget(multiple_radio)
        
        # Função para habilitar/desabilitar combo
        def toggle_combo():
            family_combo.setEnabled(single_radio.isChecked())
        
        single_radio.toggled.connect(toggle_combo)
        multiple_radio.toggled.connect(toggle_combo)
        
        # Botões
        button_layout = QHBoxLayout()
        ok_button = QPushButton("Exportar")
        cancel_button = QPushButton("Cancelar")
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        dialog.setLayout(layout)
        
        # Conectar botões
        ok_button.clicked.connect(dialog.accept)
        cancel_button.clicked.connect(dialog.reject)
        
        # Mostrar dialog
        if dialog.exec() == QDialog.DialogCode.Accepted:
            export_type = export_group.checkedId()
            
            if export_type == 1:  # Arquivo único
                selected_family = family_combo.currentData()
                
                # Seleção do arquivo
                file_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Salvar Resultados (Formato Antigo)",
                    "",
                    "Arquivos Excel (*.xlsx)"
                )
                
                if file_path:
                    try:
                        self.excel_processor.export_to_excel(file_path, selected_family)
                        family_text = "todas as famílias" if selected_family == "Todas" else f"família {selected_family}"
                        QMessageBox.information(
                            self, 
                            "Sucesso", 
                            f"Arquivo exportado no formato antigo para {family_text}:\n{file_path}\n\n"
                            f"O arquivo contém:\n"
                            f"• Cabeçalho formatado\n"
                            f"• Linhas Entradas/Sugestão/Vendas\n"
                            f"• Colunas para cada loja\n"
                            f"• Formatação profissional"
                        )
                        # Arquivo exportado
                        self.statusBar().showMessage(f"Exportado: {os.path.basename(file_path)}")
                        
                    except Exception as e:
                        QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")
                        # Erro ao exportar
            
            elif export_type == 2:  # Arquivos separados
                # Seleção do diretório
                dir_path = QFileDialog.getExistingDirectory(
                    self,
                    "Selecionar Pasta para Arquivos por Família"
                )
                
                if dir_path:
                    try:
                        exported_files = []
                        
                        for family in families:
                            # Nome do arquivo baseado na família
                            safe_family = str(family).replace("/", "_").replace("\\", "_").replace(":", "_")
                            file_name = f"Relatorio_Familia_{safe_family}.xlsx"
                            file_path = os.path.join(dir_path, file_name)
                            
                            # Exportar família específica
                            self.excel_processor.export_to_excel(file_path, family)
                            exported_files.append(file_name)
                        
                        # Também exportar arquivo com todas as famílias
                        all_file_path = os.path.join(dir_path, "Relatorio_Todas_Familias.xlsx")
                        self.excel_processor.export_to_excel(all_file_path, "Todas")
                        exported_files.append("Relatorio_Todas_Familias.xlsx")
                        
                        files_list = "\n".join([f"• {f}" for f in exported_files])
                        QMessageBox.information(
                            self, 
                            "Sucesso", 
                            f"Arquivos exportados com sucesso!\n\n"
                            f"Pasta: {dir_path}\n\n"
                            f"Arquivos criados:\n{files_list}\n\n"
                            f"Total: {len(exported_files)} arquivos"
                        )
                        # Exportados arquivos por família
                        self.statusBar().showMessage(f"Exportados {len(exported_files)} arquivos")
                        
                    except Exception as e:
                        QMessageBox.critical(self, "Erro", f"Erro ao exportar arquivos: {str(e)}")
                
    def closeEvent(self, event):
        """Tratamento do fechamento da aplicação"""
        if self.calculation_thread and self.calculation_thread.isRunning():
            reply = QMessageBox.question(
                self, 
                "Confirmação", 
                "Há um cálculo em andamento. Deseja mesmo sair?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.calculation_thread.terminate()
                self.calculation_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def create_menu_bar(self):
        """Cria a barra de menu"""
        menubar = self.menuBar()
        
        # Menu Ajuda
        help_menu = menubar.addMenu('&Ajuda')
        
        # Ação Verificar Atualizações
        update_action = help_menu.addAction('🔄 Verificar Atualizações')
        update_action.setStatusTip('Verifica se há atualizações disponíveis')
        update_action.triggered.connect(self.check_for_updates)
        
        help_menu.addSeparator()
        
        # Ação Sobre
        about_action = help_menu.addAction('ℹ️ Sobre')
        about_action.setStatusTip('Informações sobre o aplicativo')
        about_action.triggered.connect(self.show_about)
        
    def check_for_updates(self):
        """Verifica atualizações manualmente"""
        try:
            from auto_updater import AutoUpdater
            updater = AutoUpdater(self)
            updater.check_for_updates(silent=False)
        except ImportError:
            QMessageBox.warning(self, "Erro", "Sistema de atualização não disponível")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao verificar atualizações:\n{str(e)}")
    
    def show_about(self):
        """Mostra informações sobre o aplicativo"""
        about_text = """
        <h2>Sistema de Cálculo Excel Profissional</h2>
        <p><b>Versão:</b> 1.0.0</p>
        <p><b>Desenvolvido para:</b> Armarinhos Fernando Ltda</p>
        <p><b>Descrição:</b> Sistema profissional para processamento e cálculo de planilhas Excel</p>
        <p><b>Recursos:</b></p>
        <ul>
            <li>✅ Transposição automática de dados</li>
            <li>✅ Cálculo de sugestões baseado em vendas</li>
            <li>✅ Exportação com formatação</li>
            <li>✅ Sistema de licenciamento</li>
            <li>✅ Atualizações automáticas</li>
        </ul>
        <p><b>Suporte:</b> megasystems@exemplo.com</p>
        """
        
        QMessageBox.about(self, "Sobre o Sistema", about_text)

def main():
    """Função principal"""
    app = QApplication(sys.argv)
    app.setApplicationName("Sistema de Cálculo Excel")
    app.setOrganizationName("Mega Systems")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()