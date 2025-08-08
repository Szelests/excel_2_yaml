import yaml
import pandas as pd
import os
from typing import List, Dict, Any

class AtualizadorDeAlarmes:
    """
    Uma classe para ler alarmes de um arquivo Excel e mesclá-los
    em um arquivo YAML, usando os títulos das colunas para mapear os dados.
    """

    def __init__(self, yaml_path: str, excel_path: str):
        self.yaml_path = yaml_path
        self.excel_path = excel_path
        
        # --- NOVO: DICIONÁRIO DE MAPEAMENTO ---
        # Mapeia o nome interno do campo para o título da coluna no Excel.
        # Você pode personalizar os títulos aqui para corresponder ao seu arquivo!
        self.mapa_de_colunas = {
            'alarm_id': 'ID do Alarme',
            'name': 'Nome do Alarme',
            'text': 'Texto Principal',
            'text1': 'Habilitado', # Mudei para um nome que faz mais sentido com sim/não
            'text2': 'Ação Recomendada',
            'config': 'Nível de Configuração',
            'subconfig1': 'Sub-Config 1',
            'subconfig2': 'Sub-Config 2'
        }
        
        print(f"Inicializando atualização do arquivo '{yaml_path}' com dados de '{excel_path}'.")

    @staticmethod
    def _transformar_valor_booleano(valor: Any) -> Any:
        valor_str = str(valor).strip().lower()
        if valor_str == 'sim':
            return True
        elif valor_str == 'não':
            return False
        return valor

    def _carregar_alarmes_yaml(self) -> List[Dict[str, Any]]:
        # Este método não precisa de alterações
        if not os.path.exists(self.yaml_path):
            print(f"Arquivo YAML '{self.yaml_path}' não encontrado. Será criado um novo.")
            return []
        with open(self.yaml_path, 'r', encoding='utf-8') as f:
            try:
                alarms = yaml.safe_load(f) or []
                if isinstance(alarms, list):
                    print(f"Encontrados {len(alarms)} alarmes existentes no arquivo YAML.")
                    return alarms
                else:
                    print("Arquivo YAML não contém uma lista válida. Iniciando com uma lista vazia.")
                    return []
            except yaml.YAMLError as e:
                print(f"Erro ao ler o arquivo YAML: {e}. Iniciando com uma lista vazia.")
                return []

    def _carregar_alarmes_excel(self) -> List[Dict[str, Any]]:
        """
        Lê o Excel usando a primeira linha como cabeçalho e mapeia as colunas pelo título.
        """
        try:
            # --- ALTERAÇÃO CRÍTICA AQUI ---
            # `header=0` diz ao pandas para usar a primeira linha como cabeçalho.
            df = pd.read_excel(self.excel_path, header=0, dtype=str).fillna('')

            # Validação: Verifica se todas as colunas mapeadas existem no Excel
            colunas_necessarias = self.mapa_de_colunas.values()
            colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
            if colunas_faltantes:
                print(f"ERRO: As seguintes colunas obrigatórias não foram encontradas no Excel: {colunas_faltantes}")
                return []

            novos_alarmes = []
            for index, row in df.iterrows():
                # Usa o mapa para buscar os dados pelo NOME da coluna, não pela posição
                mapa = self.mapa_de_colunas
                alarm_data = {
                    'alarm_id': int(row[mapa['alarm_id']]),
                    'name': row[mapa['name']],
                    'text': {
                        'text': row[mapa['text']],
                        'text1': self._transformar_valor_booleano(row[mapa['text1']]),
                        'text2': row[mapa['text2']]
                    },
                    'config': {
                        'config': row[mapa['config']],
                        'subconfig1': row[mapa['subconfig1']],
                        'subconfig2': row[mapa['subconfig2']]
                    }
                }
                novos_alarmes.append(alarm_data)
            print(f"Encontrados {len(novos_alarmes)} alarmes no arquivo Excel para processar.")
            return novos_alarmes
        except FileNotFoundError:
            print(f"Erro: Arquivo Excel '{self.excel_path}' não encontrado.")
            return []
        except ValueError as e:
            print(f"Erro de valor ao processar o Excel. Verifique se o 'ID do Alarme' é sempre um número. Detalhes: {e}")
            return []
        except Exception as e:
            print(f"Erro ao processar o arquivo Excel: {e}")
            return []
    
    def _salvar_yaml(self, data: List[Dict[str, Any]]):
        # Este método não precisa de alterações
        with open(self.yaml_path, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, allow_unicode=True, sort_keys=False, indent=2, default_flow_style=False)
        print(f"Arquivo '{self.yaml_path}' salvo com sucesso com um total de {len(data)} alarmes. A ORDEM ORIGINAL FOI MANTIDA.")


    def executar_atualizacao(self):
        # Este método não precisa de alterações
        alarmes_existentes = self._carregar_alarmes_yaml()
        novos_alarmes = self._carregar_alarmes_excel()

        if not novos_alarmes:
            print("Nenhum alarme novo para adicionar ou erro na leitura. Nenhuma alteração foi feita.")
            return
            
        ids_em_uso = {alarm['alarm_id'] for alarm in alarmes_existentes}
        alarmes_processados = []

        for alarme_novo in novos_alarmes:
            id_alvo = alarme_novo['alarm_id']
            if id_alvo in ids_em_uso:
                id_novo_provisorio = id_alvo
                while id_novo_provisorio in ids_em_uso:
                    id_novo_provisorio += 1
                print(f"  - CONFLITO: ID de prioridade {id_alvo} já está em uso. Atribuindo o próximo ID livre mais próximo: {id_novo_provisorio}.")
                alarme_novo['alarm_id'] = id_novo_provisorio
            ids_em_uso.add(alarme_novo['alarm_id'])
            alarmes_processados.append(alarme_novo)

        lista_final = alarmes_existentes + alarmes_processados
        self._salvar_yaml(lista_final)

# --- Bloco de Execução Principal ---
if __name__ == "__main__":
    ARQUIVO_YAML = 'test.yaml'
    ARQUIVO_EXCEL = 'test_yaml.xlsx'
    
    atualizador = AtualizadorDeAlarmes(yaml_path=ARQUIVO_YAML, excel_path=ARQUIVO_EXCEL)
    atualizador.executar_atualizacao()