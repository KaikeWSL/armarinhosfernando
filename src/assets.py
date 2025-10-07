"""
Recursos visuais e assets para a interface
"""

# Logo base64 da Armarinhos Fernando (será substituída pela real)
ARMARINHOS_LOGO_BASE64 = """
iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==
"""

# Cores do tema dark moderno
COLORS = {
    'primary': '#0066CC',          # Azul da logo
    'secondary': '#FF4444',        # Vermelho da logo
    'dark_bg': "#141f33",         # Fundo escuro azulado principal
    'darker_bg': '#0f1419',       # Fundo azul mais escuro
    'card_bg': '#243447',         # Fundo dos cards azulado
    'accent': '#007ACC',          # Azul de destaque
    'text_primary': '#ffffff',     # Texto principal
    'text_secondary': '#cccccc',   # Texto secundário
    'text_muted': '#888888',      # Texto esmaecido
    'success': '#4CAF50',         # Verde sucesso
    'warning': '#FF9800',         # Laranja aviso
    'error': '#f44336',           # Vermelho erro
    'border': '#3d5a80',          # Bordas azuladas
    'hover': '#0052a3',           # Hover azul
    'hover_red': '#cc3333',       # Hover vermelho
    
    # Cores específicas para tabela
    'table_bg': '#1a2f4d',        # Fundo azul da tabela
    'table_header': '#0066CC',    # Header azul
    'vendas_row': "#faed7a",      # Amarelo para linhas de vendas
    'sugestao_row': "#aefdb1",    # Verde para linhas de sugestão
    'entradas_row': "#e5f3ff"     # Azul para linhas de entradas
}

# Gradientes
GRADIENTS = {
    'header': f"qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {COLORS['primary']}, stop:1 {COLORS['secondary']})",
    'button': f"qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {COLORS['primary']}, stop:1 {COLORS['hover']})",
    'card': f"qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {COLORS['card_bg']}, stop:1 {COLORS['darker_bg']})"
}