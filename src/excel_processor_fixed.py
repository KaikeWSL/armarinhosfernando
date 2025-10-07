import pandas as pd
import re
from typing import List, Dict, Optional

class ExcelProcessor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df = None
        self.transposed_df = None
        
    def load_excel_file(self):
        """Carrega o arquivo Excel"""
        try:
            self.df = pd.read_excel(self.file_path)
            print(f"📂 Arquivo carregado: {len(self.df)} linhas, {len(self.df.columns)} colunas")
            return True
        except Exception as e:
            print(f"❌ Erro ao carregar arquivo: {e}")
            return False

    def transpose_to_old_format(self):
        """Transpõe dados do novo formato para o formato antigo"""
        if self.df is None:
            print("❌ Nenhum dado carregado")
            return pd.DataFrame()
        
        print("\n🔄 FASE 1: Analisando dados...")
        products_data = []
        stores_found = []
        current_store = None
        
        # Itera pelas linhas do DataFrame
        for idx, row in self.df.iterrows():
            first_col = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            
            # Detecta linha de loja
            if first_col.startswith("Loja:"):
                store_code = self.extract_store_code(first_col)
                if store_code:
                    current_store = store_code
                    if store_code not in stores_found:
                        stores_found.append(store_code)
                    print(f"🏪 Loja: {store_code}")
                continue
            
            # DEBUG específico para produto 1075707
            if "1075707" in str(row.iloc[2]):
                print(f"  🔍 DEBUG linha {idx}: raw_codigo={row.iloc[2]}, req={self.safe_float(row.iloc[12])}, vendas={self.safe_float(row.iloc[14])}")
            
            # Detecta linha de produto
            if current_store and self.is_product_line(row):
                try:
                    product_info = self.extract_product_data(row, current_store)
                    if product_info:
                        products_data.append(product_info)
                        print(f"  📦 {product_info['codigo']} em {current_store}: V={product_info['vendas']}, E={product_info['entradas']}")
                except Exception as e:
                    print(f"⚠️ Erro linha {idx}: {str(e)}")
                    continue

        print(f"\n📊 RESUMO DETECTADO:")
        print(f"🏪 Lojas: {stores_found}")
        print(f"📦 Total de registros produto-loja: {len(products_data)}")
        
        if not stores_found or not products_data:
            print("❌ Nenhum dado válido encontrado")
            return pd.DataFrame()
        
        print("\n🔄 FASE 2: Criando formato antigo...")
        
        # Agrupa produtos por código (SEM DUPLICAÇÃO)
        products_by_code = {}
        for product in products_data:
            codigo = product['codigo']
            if codigo not in products_by_code:
                products_by_code[codigo] = {
                    'descricao': product['descricao'],
                    'secundario': product['secundario'],
                    'lojas_vendas': {},
                    'lojas_entradas': {}
                }
            
            # Armazena valores específicos da loja (SOBRESCREVE se já existe)
            loja = product['loja']
            products_by_code[codigo]['lojas_vendas'][loja] = product['vendas']
            products_by_code[codigo]['lojas_entradas'][loja] = product['entradas']
        
        print(f"🗂️ Produtos únicos: {len(products_by_code)}")
        
        # Cria DataFrame no formato antigo
        result_data = []
        
        for codigo, info in products_by_code.items():
            # Linha de Entradas
            entrada_row = {
                'Codigo': codigo,
                'Descricao': info['descricao'],
                'Secundario': info['secundario'],
                'Tipo': 'Entradas'
            }
            
            # Adiciona valores de entrada para cada loja
            for loja in stores_found:
                entrada_row[loja] = info['lojas_entradas'].get(loja, 0)
            
            # Calcula total
            entrada_row['Qtde Total'] = sum(info['lojas_entradas'].values())
            result_data.append(entrada_row)
            
            # Linha de Vendas
            vendas_row = {
                'Codigo': codigo,
                'Descricao': info['descricao'],
                'Secundario': info['secundario'],
                'Tipo': 'Vendas'
            }
            
            # Adiciona valores de vendas para cada loja
            for loja in stores_found:
                vendas_row[loja] = info['lojas_vendas'].get(loja, 0)
            
            # Calcula total
            vendas_row['Qtde Total'] = sum(info['lojas_vendas'].values())
            result_data.append(vendas_row)
        
        # Cria DataFrame resultado
        if result_data:
            # Ordena colunas: info, tipo, lojas, total
            base_columns = ['Codigo', 'Descricao', 'Secundario', 'Tipo']
            store_columns = sorted(stores_found)
            final_columns = base_columns + store_columns + ['Qtde Total']
            
            result_df = pd.DataFrame(result_data)
            result_df = result_df.reindex(columns=final_columns, fill_value=0)
            
            # Atualiza o DataFrame da instância
            self.transposed_df = result_df
            
            print(f"📊 Resultado: {len(result_df)} linhas, {len(result_df.columns)} colunas")
            print(f"🏪 Colunas de lojas: {[col for col in result_df.columns if col in stores_found]}")
            
            # Debug: mostra primeiros produtos
            if len(result_df) > 0:
                print(f"\n📋 EXEMPLO - Primeiro produto:")
                primeiro_produto = result_df.iloc[0]
                print(f"  Código: {primeiro_produto['Codigo']}")
                for loja in stores_found:
                    if loja in primeiro_produto:
                        print(f"  {loja}: {primeiro_produto[loja]}")
            
            return result_df
        else:
            return pd.DataFrame()

    def extract_store_code(self, text):
        """Extrai código da loja dos parênteses"""
        match = re.search(r'\(([^)]+)\)', text)
        if match:
            return match.group(1).strip()
        return None

    def is_product_line(self, row):
        """Verifica se a linha contém dados de produto"""
        try:
            # Produto na coluna C (index 2)
            if pd.isna(row.iloc[2]):
                return False
            
            codigo_str = str(row.iloc[2]).strip()
            
            # Verifica se é um código numérico válido
            if not codigo_str or codigo_str == "":
                return False
            
            # Remove pontos decimais se houver
            codigo_clean = codigo_str.replace('.0', '')
            
            return codigo_clean.isdigit() and len(codigo_clean) >= 6
        except:
            return False

    def extract_product_data(self, row, loja):
        """Extrai dados do produto de uma linha"""
        try:
            # Código na coluna C (index 2)
            codigo_raw = row.iloc[2]
            if pd.isna(codigo_raw):
                return None
                
            codigo = str(int(float(codigo_raw))) if isinstance(codigo_raw, (int, float)) else str(codigo_raw)
            
            # Descrição na coluna D (index 3)
            descricao = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
            
            # Secundário na coluna E (index 4) 
            secundario = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""
            
            # Requisições na coluna M (index 12)
            entradas = self.safe_float(row.iloc[12])
            
            # Vendas na coluna O (index 14)
            vendas = self.safe_float(row.iloc[14])
            
            return {
                'codigo': codigo,
                'descricao': descricao,
                'secundario': secundario,
                'loja': loja,
                'entradas': entradas,
                'vendas': vendas
            }
        except Exception as e:
            print(f"❌ Erro ao extrair produto: {e}")
            return None

    def safe_float(self, value):
        """Converte valor para float de forma segura"""
        try:
            if pd.isna(value):
                return 0.0
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def get_data_preview(self):
        """Retorna preview dos dados transpostos"""
        if self.transposed_df is not None:
            return self.transposed_df.head(10)
        return pd.DataFrame()

    def calculate_suggestions(self, percentage: float, progress_callback=None):
        """Calcula sugestões com base na porcentagem"""
        if self.transposed_df is None:
            return False
        
        # Encontra linhas de vendas
        vendas_rows = self.transposed_df[self.transposed_df['Tipo'] == 'Vendas'].copy()
        
        if vendas_rows.empty:
            return False
        
        # Identifica colunas de lojas
        info_columns = ['Codigo', 'Descricao', 'Secundario', 'Tipo']
        store_columns = [col for col in self.transposed_df.columns 
                        if col not in info_columns + ['Qtde Total']]
        
        suggestions = []
        total_products = len(vendas_rows)
        
        for idx, (_, row) in enumerate(vendas_rows.iterrows()):
            # Cria linha de sugestão
            suggestion_row = row.copy()
            suggestion_row['Tipo'] = 'Sugestão'
            
            # Aplica porcentagem nas colunas de lojas
            for col in store_columns:
                original_value = row[col]
                if pd.notna(original_value) and original_value > 0:
                    suggestion_row[col] = int(original_value * (1 + percentage / 100))
            
            # Recalcula total
            suggestion_row['Qtde Total'] = sum([suggestion_row[col] for col in store_columns if pd.notna(suggestion_row[col])])
            
            suggestions.append(suggestion_row)
            
            # Callback de progresso
            if progress_callback:
                progress_callback(int((idx + 1) / total_products * 100))
        
        # Adiciona sugestões ao DataFrame
        if suggestions:
            suggestions_df = pd.DataFrame(suggestions)
            self.transposed_df = pd.concat([self.transposed_df, suggestions_df], ignore_index=True)
        
        return True

    def export_results(self, output_path: str):
        """Exporta resultados para Excel"""
        if self.transposed_df is None:
            return False
        
        try:
            self.transposed_df.to_excel(output_path, index=False)
            return True
        except Exception as e:
            print(f"❌ Erro ao exportar: {e}")
            return False

    def get_calculation_summary(self):
        """Retorna resumo"""
        if self.transposed_df is None:
            return {'message': 'Nenhum dado processado'}
        
        vendas_count = len(self.transposed_df[self.transposed_df['Tipo'] == 'Vendas'])
        entradas_count = len(self.transposed_df[self.transposed_df['Tipo'] == 'Entradas'])
        sugestoes_count = len(self.transposed_df[self.transposed_df['Tipo'] == 'Sugestão'])
        
        return {
            'total_rows': len(self.transposed_df),
            'vendas_rows': vendas_count,
            'entradas_rows': entradas_count,
            'suggestion_rows': sugestoes_count,
            'produtos_unicos': vendas_count
        }
    
    def process_file(self):
        """Processa o arquivo completo"""
        try:
            self.load_excel_file()
            self.transpose_to_old_format()
            return True
        except Exception as e:
            print(f"Erro no processamento: {e}")
            return False
    
    def get_preview_data(self) -> pd.DataFrame:
        """Retorna dados para preview"""
        if self.transposed_df is not None:
            return self.transposed_df.head(20)  # Primeiras 20 linhas
        return pd.DataFrame()
    
    def apply_percentage_adjustment(self, percentage: float):
        """Aplica ajuste percentual nas vendas"""
        if self.transposed_df is None:
            return
        
        # Encontra colunas das lojas (entre 'Entrada' e 'Qtde Total')
        try:
            entrada_col_idx = self.transposed_df.columns.get_loc('Entrada')
            qtde_total_col_idx = self.transposed_df.columns.get_loc('Qtde Total')
            store_columns = self.transposed_df.columns[entrada_col_idx + 1:qtde_total_col_idx]
        except KeyError:
            # Se não tiver essas colunas, tenta identificar as lojas
            store_columns = [col for col in self.transposed_df.columns 
                           if col not in ['Produto', 'Tipo', 'Entrada', 'Qtde Total']]
        
        # Processa cada linha de vendas
        suggestions_to_add = []
        for idx, row in self.transposed_df.iterrows():
            if row['Tipo'] == 'Vendas':
                # Cria linha de sugestão
                suggestion_row = row.copy()
                suggestion_row['Tipo'] = 'Sugestão'
                
                # Aplica porcentagem nas colunas de lojas
                for col in store_columns:
                    if pd.notna(row[col]) and row[col] != 0:
                        suggestion_row[col] = int(row[col] * (1 + percentage / 100))
                
                suggestions_to_add.append(suggestion_row)
        
        # Adiciona sugestões ao DataFrame
        if suggestions_to_add:
            suggestions_df = pd.DataFrame(suggestions_to_add)
            self.transposed_df = pd.concat([self.transposed_df, suggestions_df], ignore_index=True)
    
    def export_to_excel(self, output_path: str):
        """Exporta o resultado para Excel no formato antigo"""
        try:
            from openpyxl import Workbook
            from openpyxl.utils.dataframe import dataframe_to_rows
            
            # Cria novo workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Formato Antigo"
            
            # Escreve dados
            for r_idx, row in enumerate(dataframe_to_rows(self.transposed_df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Salva arquivo
            wb.save(output_path)
            print(f"Arquivo exportado: {output_path}")
            
        except Exception as e:
            print(f"Erro ao exportar: {e}")
            raise