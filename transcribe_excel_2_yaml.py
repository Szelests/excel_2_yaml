import yaml
import pandas as pd
import os
from typing import List, Dict, Any

class AtualizadorDeAlarmes:
    """
    Uma classe para ler alarmes de um arquivo Excel e mesclá-los
    em um arquivo YAML, com transformações de dados e respeito à prioridade.
    """

    def __init__(self, yaml_path: str, excel_path: str):
        self.yaml_path = yaml_path
        self.excel_path = excel_path
        print(f"Inicializando atualização do arquivo '{yaml_path}' com dados de '{excel_path}'.")

    # --- NOVO MÉTODO AUXILIAR ---
    @staticmethod
    def _transformar_valor_booleano(valor: Any) -> Any:
        """
        Transforma 'sim'/'não' (case-insensitive) em True/False.
        Retorna o valor original se não for uma correspondência.
        """
        # Converte o valor para string, remove espaços e coloca em minúsculas
        valor_str = str(valor).strip().lower()
        
        if valor_str == 'sim':
            return True
        elif valor_str == 'não':
            return False
        
        # Se não for 'sim' nem 'não', retorna o valor original
        return valor

    def _carregar_alarmes_yaml(self) -> List[Dict[str, Any]]:
        if not os.path.exists(self.yaml_path):
            print(f"Arquivo YAML '{self.yaml_path}' não encontrado. Será criado um novo.")
            return []
        
        with open(self.yaml_path, 'r', encoding='utf-8') as f:
            try:
                alarms = yaml.safe_load(f)
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
        Lê o arquivo Excel e o transforma em uma lista de dicionários de alarme.
        """
        try:
            df = pd.read_excel(self.excel_path, header=None)
            novos_alarmes = []
            for index, row in df.iterrows():
                # Aplica a transformação ao valor da coluna D (índice 3)
                valor_text1_processado = self._transformar_valor_booleano(row[3])
                
                alarm_data = {
                    'alarm_id': int(row[0]), 
                    'name': str(row[1]),
                    'text': {
                        'text': str(row[2]),
                        'text1': valor_text1_processado,
                        'text2': str(row[4])
                    },
                    'config': {
                        'config': str(row[5]), 
                        'subconfig1': str(row[6]), 
                        'subconfig2': str(row[7])
                    }
                }
                novos_alarmes.append(alarm_data)
            print(f"Encontrados {len(novos_alarmes)} alarmes no arquivo Excel para processar.")
            return novos_alarmes
        except FileNotFoundError:
            print(f"Erro: Arquivo Excel '{self.excel_path}' não encontrado.")
            return []
        except Exception as e:
            print(f"Erro ao processar o arquivo Excel: {e}")
            return []

    def _salvar_yaml(self, data: List[Dict[str, Any]]):
        """Salva a lista de dados, mantendo a ordem original."""
        data_sorted = sorted(data, key=lambda k: k['alarm_id'])
        with open(self.yaml_path, 'w', encoding='utf-8') as f:
            yaml.dump(data_sorted, f, allow_unicode=True, sort_keys=False, indent=2, default_flow_style=False)
        print(f"Arquivo '{self.yaml_path}' salvo com sucesso com um total de {len(data_sorted)} alarmes. A ORDEM ORIGINAL FOI MANTIDA.")


    def executar_atualizacao(self):
        alarmes_existentes = self._carregar_alarmes_yaml()
        novos_alarmes = self._carregar_alarmes_excel()

        if not novos_alarmes:
            print("Nenhum alarme novo para adicionar. Nenhuma alteração foi feita.")
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