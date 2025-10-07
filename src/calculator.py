"""
Módulo de cálculos especializados
Implementa lógicas específicas de negócio para processamento dos dados
"""

import pandas as pd
import numpy as np
from typing import List, Dict, Any, Optional

class CalculationEngine:
    """Engine de cálculos para operações especializadas"""
    
    def __init__(self):
        self.calculation_history = []
        
    def calculate_percentage_adjustment(self, 
                                      values: List[float], 
                                      percentage: float) -> List[float]:
        """
        Aplica ajuste percentual em uma lista de valores
        
        Args:
            values: Lista de valores numéricos
            percentage: Porcentagem de ajuste
            
        Returns:
            Lista de valores ajustados
        """
        adjusted_values = []
        
        for value in values:
            if pd.notna(value) and isinstance(value, (int, float)):
                adjusted_value = value * (1 + percentage / 100)
                adjusted_values.append(adjusted_value)
            else:
                adjusted_values.append(value)
                
        return adjusted_values
        
    def sum_range(self, values: List[float]) -> float:
        """
        Soma uma faixa de valores, ignorando valores não numéricos
        
        Args:
            values: Lista de valores para somar
            
        Returns:
            Soma dos valores válidos
        """
        total = 0
        
        for value in values:
            if pd.notna(value) and isinstance(value, (int, float)):
                total += value
                
        return total
        
    def validate_data_row(self, row: pd.Series, required_columns: List[int]) -> bool:
        """
        Valida se uma linha possui dados válidos nas colunas obrigatórias
        
        Args:
            row: Série pandas representando uma linha
            required_columns: Lista de índices de colunas obrigatórias
            
        Returns:
            True se a linha for válida
        """
        for col_idx in required_columns:
            if col_idx < len(row) and pd.notna(row.iloc[col_idx]):
                continue
            else:
                return False
                
        return True
        
    def apply_business_rules(self, 
                           df: pd.DataFrame, 
                           percentage: float,
                           progress_callback=None) -> pd.DataFrame:
        """
        Aplica regras de negócio específicas para o novo formato
        Adiciona coluna de Sugestão baseada nos valores de Vendas
        
        Args:
            df: DataFrame com os dados
            percentage: Porcentagem de ajuste
            progress_callback: Callback para atualizar progresso
            
        Returns:
            DataFrame modificado com coluna de sugestão
        """
        modified_df = df.copy()
        
        # Encontra linhas de produtos (não cabeçalhos de loja)
        product_indices = self.find_product_rows(modified_df)
        
        if not product_indices:
            print("Nenhuma linha de produto encontrada para processar")
            return modified_df
            
        total_operations = len(product_indices)
        print(f"Processando {total_operations} linhas de produtos...")
        
        # Identifica colunas importantes
        cols = list(modified_df.columns)
        vendas_col_idx = 5 if len(cols) > 5 else None  # Coluna F (índice 5): Vendas
        
        # Adiciona coluna de Sugestão se não existir
        if len(cols) < 10:  # Garante que temos colunas suficientes
            # Adiciona colunas vazias até chegar na posição da Sugestão (coluna S = índice 18)
            while len(modified_df.columns) < 19:
                new_col_name = f"Col_{len(modified_df.columns)}"
                modified_df[new_col_name] = ""
                
        # Define nome da coluna de sugestão (coluna S)
        suggestion_col_idx = 18  # Coluna S (índice 18)
        suggestion_col_name = modified_df.columns[suggestion_col_idx] if len(modified_df.columns) > suggestion_col_idx else "Sugestão"
        
        # Se a coluna não existir, adiciona
        if suggestion_col_idx >= len(modified_df.columns):
            modified_df["Sugestão"] = ""
            suggestion_col_name = "Sugestão"
        
        # Processa cada linha de produto
        for i, product_idx in enumerate(product_indices):
            if progress_callback:
                progress = int((i + 1) / total_operations * 100)
                progress_callback.emit(progress)
                
            # Calcula sugestão baseada nas vendas
            if vendas_col_idx and vendas_col_idx < len(modified_df.columns):
                vendas_value = modified_df.iloc[product_idx, vendas_col_idx]
                
                if pd.notna(vendas_value) and isinstance(vendas_value, (int, float)) and vendas_value > 0:
                    suggestion_value = vendas_value * (1 + percentage / 100)
                    modified_df.iloc[product_idx, suggestion_col_idx] = suggestion_value
                    print(f"Linha {product_idx}: Vendas {vendas_value} -> Sugestão {suggestion_value}")
                else:
                    modified_df.iloc[product_idx, suggestion_col_idx] = 0
                    
        # Registra no histórico
        self.calculation_history.append({
            'operation': 'apply_business_rules_new_format',
            'percentage': percentage,
            'rows_processed': len(product_indices),
            'timestamp': pd.Timestamp.now()
        })
        
        return modified_df
        
    def find_product_rows(self, df: pd.DataFrame) -> List[int]:
        """
        Encontra linhas de produtos no novo formato
        """
        product_rows = []
        
        if df.empty or len(df.columns) < 2:
            return product_rows
            
        zona_col = df.columns[0]  # Coluna A: Zona
        codigo_col = df.columns[1]  # Coluna B: Código
        
        for idx, row in df.iterrows():
            zona_val = row[zona_col]
            codigo_val = row[codigo_col]
            
            if pd.notna(codigo_val) and pd.notna(zona_val):
                zona_str = str(zona_val).strip()
                codigo_str = str(codigo_val).strip()
                
                # É linha de produto se não é cabeçalho de loja e tem código
                if not zona_str.startswith("Loja:") and codigo_str.isdigit():
                    product_rows.append(idx)
                    
        return product_rows
        
    def create_suggestion_row(self, vendas_row: pd.Series, percentage: float) -> pd.Series:
        """
        Cria uma linha de sugestão baseada em uma linha de vendas
        
        Args:
            vendas_row: Linha de vendas original
            percentage: Porcentagem de ajuste
            
        Returns:
            Nova linha de sugestão
        """
        suggestion_row = vendas_row.copy()
        
        # Define primeira coluna como "Sugestão"
        suggestion_row.iloc[0] = "Sugestão"
        
        # Copia valores das colunas 2-10 (mantém dados identificadores)
        # Colunas: Tipo, Código, Descrição, Cx c/, Secundário, Saldo Local, etc.
        for col_idx in range(1, min(10, len(suggestion_row))):
            suggestion_row.iloc[col_idx] = vendas_row.iloc[col_idx]
            
        # Aplica porcentagem nas colunas de valores (a partir da coluna de entrada/vendas)
        # Baseado na sua planilha, as colunas numéricas começam depois das informações básicas
        soma = 0
        num_cols = len(suggestion_row)
        
        # Procura pelas colunas numéricas (geralmente a partir da coluna 10)
        for col_idx in range(9, num_cols):  # Ajustado para começar mais tarde
            original_value = vendas_row.iloc[col_idx]
            if pd.notna(original_value) and isinstance(original_value, (int, float)) and original_value != 0:
                adjusted_value = original_value * (1 + percentage / 100)
                suggestion_row.iloc[col_idx] = adjusted_value
                soma += adjusted_value
            else:
                suggestion_row.iloc[col_idx] = original_value
                
        return suggestion_row
        
    def insert_row_after(self, 
                        df: pd.DataFrame, 
                        after_idx: int, 
                        new_row: pd.Series) -> pd.DataFrame:
        """
        Insere uma nova linha após o índice especificado
        
        Args:
            df: DataFrame original
            after_idx: Índice após o qual inserir a linha
            new_row: Nova linha a ser inserida
            
        Returns:
            DataFrame com a nova linha inserida
        """
        # Divide o DataFrame
        before = df.iloc[:after_idx + 1]
        after = df.iloc[after_idx + 1:]
        
        # Cria DataFrame com a nova linha
        new_row_df = pd.DataFrame([new_row])
        
        # Concatena tudo
        result = pd.concat([before, new_row_df, after], ignore_index=True)
        
        return result
        
    def format_numeric_columns(self, 
                             df: pd.DataFrame, 
                             column_range: tuple = (10, 29)) -> pd.DataFrame:
        """
        Formata colunas numéricas para garantir tipos corretos
        
        Args:
            df: DataFrame a ser formatado
            column_range: Tupla com (início, fim) das colunas numéricas
            
        Returns:
            DataFrame com colunas formatadas
        """
        formatted_df = df.copy()
        start_col, end_col = column_range
        
        for col_idx in range(start_col, min(end_col, len(formatted_df.columns))):
            formatted_df.iloc[:, col_idx] = pd.to_numeric(
                formatted_df.iloc[:, col_idx], 
                errors='ignore'
            )
            
        return formatted_df
        
    def validate_calculation_results(self, 
                                   original_df: pd.DataFrame, 
                                   modified_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Valida os resultados dos cálculos
        
        Args:
            original_df: DataFrame original
            modified_df: DataFrame modificado
            
        Returns:
            Dicionário com resultados da validação
        """
        validation_results = {
            'original_rows': len(original_df),
            'modified_rows': len(modified_df),
            'rows_added': len(modified_df) - len(original_df),
            'suggestions_found': 0,
            'vendas_found': 0,
            'validation_passed': True,
            'errors': []
        }
        
        try:
            # Conta sugestões e vendas
            primeira_coluna = modified_df.columns[0] if len(modified_df.columns) > 0 else None
            
            if primeira_coluna:
                for idx, row in modified_df.iterrows():
                    valor = row[primeira_coluna]
                    if pd.notna(valor):
                        cell_value = str(valor).strip()
                        if cell_value == "Sugestão":
                            validation_results['suggestions_found'] += 1
                        elif cell_value == "Vendas":
                            validation_results['vendas_found'] += 1
                        
            # Validações básicas
            if validation_results['suggestions_found'] == 0:
                validation_results['errors'].append("Nenhuma sugestão foi criada")
                validation_results['validation_passed'] = False
                
            if validation_results['vendas_found'] == 0:
                validation_results['errors'].append("Nenhuma linha de vendas encontrada")
                validation_results['validation_passed'] = False
                
        except Exception as e:
            validation_results['errors'].append(f"Erro durante validação: {str(e)}")
            validation_results['validation_passed'] = False
            
        return validation_results
        
    def get_calculation_statistics(self) -> Dict[str, Any]:
        """
        Retorna estatísticas dos cálculos realizados
        
        Returns:
            Dicionário com estatísticas
        """
        if not self.calculation_history:
            return {'message': 'Nenhum cálculo realizado ainda'}
            
        stats = {
            'total_calculations': len(self.calculation_history),
            'last_calculation': self.calculation_history[-1],
            'total_rows_processed': sum(calc.get('rows_processed', 0) 
                                      for calc in self.calculation_history)
        }
        
        return stats