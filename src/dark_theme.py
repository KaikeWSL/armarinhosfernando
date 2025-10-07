"""
Tema dark moderno para a interface
"""

def get_dark_theme_stylesheet():
    """Retorna o stylesheet do tema dark"""
    from assets import COLORS, GRADIENTS
    
    return f"""
        /* Janela principal */
        QMainWindow {{
            background-color: {COLORS['dark_bg']};
            color: {COLORS['text_primary']};
        }}
        
        /* Content Widget */
        #contentWidget {{
            background-color: {COLORS['dark_bg']};
        }}
        
        /* Header Banner */
        #headerBanner {{
            background: {GRADIENTS['header']};
            border: none;
            border-bottom: 3px solid {COLORS['accent']};
        }}
        
        #mainTitle {{
            color: {COLORS['text_primary']};
            font-size: 16px;
            font-weight: bold;

        }}
        
        #subtitle {{
            color: {COLORS['text_secondary']};
            font-size: 10px;
            margin: 0px;
        }}
        
        #versionLabel {{
            color: {COLORS['text_primary']};
            font-size: 12px;
            font-weight: bold;
            background-color: rgba(255,255,255,0.1);
            padding: 8px 15px;
            border-radius: 15px;
        }}
        
        /* Groups modernos */
        QGroupBox#modernGroup {{
            font-weight: bold;
            font-size: 14px;
            border: 2px solid {COLORS['border']};
            border-radius: 10px;
            margin: 10px 0;
            padding-top: 20px;
            background: {GRADIENTS['card']};
            color: {COLORS['text_primary']};
        }}
        
        QGroupBox#modernGroup::title {{
            subcontrol-origin: margin;
            left: 15px;
            padding: 5px 10px;
            color: {COLORS['accent']};
            background-color: {COLORS['card_bg']};
            border-radius: 5px;
        }}
        
        /* File container */
        #fileContainer {{
            background-color: {COLORS['darker_bg']};
            border: 2px dashed {COLORS['border']};
            border-radius: 8px;
            padding: 15px;
        }}
        
        #fileLabel {{
            color: {COLORS['text_secondary']};
            font-size: 13px;
            padding: 5px;
        }}
        
        /* Parameter container */
        #paramContainer {{
            background-color: {COLORS['darker_bg']};
            border: 1px solid {COLORS['border']};
            border-radius: 8px;
            padding: 10px;
        }}
        
        #paramLabel {{
            color: {COLORS['text_primary']};
            font-weight: bold;
            margin-right: 10px;
        }}
        
        /* Botões */
        QPushButton#primaryButton {{
            background: {GRADIENTS['button']};
            color: {COLORS['text_primary']};
            border: none;
            padding: 12px 25px;
            border-radius: 8px;
            font-weight: bold;
            font-size: 13px;
            min-width: 150px;
        }}
        
        QPushButton#primaryButton:hover {{
            background-color: {COLORS['hover']};
        }}
        
        QPushButton#successButton {{
            background-color: {COLORS['success']};
            color: {COLORS['text_primary']};
            border: none;
            padding: 12px 30px;
            border-radius: 8px;
            font-weight: bold;
            font-size: 14px;
            min-width: 180px;
        }}
        
        QPushButton#successButton:hover {{
            background-color: #45a049;
        }}
        
        QPushButton#successButton:disabled {{
            background-color: {COLORS['text_muted']};
            color: {COLORS['darker_bg']};
        }}
        
        QPushButton#warningButton {{
            background-color: {COLORS['warning']};
            color: {COLORS['text_primary']};
            border: none;
            padding: 12px 30px;
            border-radius: 8px;
            font-weight: bold;
            font-size: 14px;
            min-width: 180px;
        }}
        
        QPushButton#warningButton:hover {{
            background-color: #e68900;
        }}
        
        QPushButton#warningButton:disabled {{
            background-color: {COLORS['text_muted']};
            color: {COLORS['darker_bg']};
        }}
        
        /* SpinBox moderno */
        QDoubleSpinBox#modernSpinBox {{
            background-color: {COLORS['darker_bg']};
            border: 2px solid {COLORS['border']};
            border-radius: 6px;
            padding: 8px 12px;
            color: {COLORS['text_primary']};
            font-size: 13px;
            font-weight: bold;
        }}
        
        QDoubleSpinBox#modernSpinBox:focus {{
            border-color: {COLORS['accent']};
        }}
        

        
        /* Log moderno */
        QTextEdit#modernLog {{
            background-color: {COLORS['darker_bg']};
            border: 1px solid {COLORS['border']};
            border-radius: 8px;
            color: {COLORS['text_secondary']};
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            padding: 10px;
        }}
        
        /* Progress Bar */
        QProgressBar#modernProgressBar {{
            background-color: {COLORS['darker_bg']};
            border: 1px solid {COLORS['border']};
            border-radius: 4px;
            text-align: center;
            color: {COLORS['text_primary']};
        }}
        
        QProgressBar#modernProgressBar::chunk {{
            background: {GRADIENTS['header']};
            border-radius: 3px;
        }}
        
        /* Status Bar */
        #modernStatusBar {{
            background: {GRADIENTS['card']};
            color: {COLORS['text_primary']};
            border-top: 1px solid {COLORS['border']};
            font-weight: bold;
            padding: 5px;
        }}
    """