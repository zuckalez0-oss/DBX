# history_manager.py

import json
from datetime import datetime

class HistoryManager:
    """
    Gerencia o histórico de projetos, salvando e carregando dados de um arquivo JSON.
    """
    def __init__(self, history_path="project_history.json"):
        self.history_path = history_path

    def _load_history(self):
        try:
            with open(self.history_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def _save_history(self, history_data):
        with open(self.history_path, 'w', encoding='utf-8') as f:
            json.dump(history_data, f, indent=4)

    def get_projects(self):
        return sorted(self._load_history().keys())

    def get_project_data(self, project_number):
        return self._load_history().get(project_number, {}).get('pieces', [])

    def save_project(self, project_number, pieces_df):
        history = self._load_history()
        df_copy = pieces_df.copy()
        # Garante que a coluna 'furos' seja uma lista serializável em JSON
        df_copy['furos'] = df_copy['furos'].apply(lambda x: x if isinstance(x, list) else [])
        pieces_list = df_copy.to_dict('records')
        
        history[project_number] = {
            "project_number": project_number,
            "save_date": datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            "pieces": pieces_list
        }
        self._save_history(history)

    def delete_project(self, project_number):
        history = self._load_history()
        if project_number in history:
            del history[project_number]
            self._save_history(history)
            return True
        return False