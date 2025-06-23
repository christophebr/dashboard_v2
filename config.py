import os
from pathlib import Path
import pickle

# Chemins des fichiers de données
AIRCALL_DATA_PATH_V1 = 'data/Affid/Aircall/data_v1'
AIRCALL_DATA_PATH_V2 = 'data/Affid/Aircall/data_v2'
HUBSPOT_TICKET_DATA_PATH = 'data/Affid/Hubspot/ticket'
HUBSPOT_AGENT_DATA_PATH = 'data/Affid/Hubspot/agent'
EVALUATION_DATA_PATH = 'data/Affid/Evaluation/support_notes_filtered.xlsx'

# Chemins des modèles
MODEL_PATH = 'models/random_forest_model.pkl'
TFIDF_PATH = 'models/tfidf_vectorizer.pkl'

# Configuration des caches pour optimisations
CACHE_PATH = 'data/Affid/Cache'
SQLITE_CACHE_PATH = f'{CACHE_PATH}/cache.sqlite'
PICKLE_CACHE_PATH = f'{CACHE_PATH}/processed_cache.pkl'
HUBSPOT_CACHE_PATH = f'{CACHE_PATH}/data_cache.db'

# Configuration des optimisations
ENABLE_CACHE = True
ENABLE_VECTORIZATION = True
ENABLE_PARALLEL_PROCESSING = False  # À activer si vous avez plusieurs cœurs

# Charger les mots de passe hachés à partir du fichier
file_path = Path(__file__).parent / 'hashed_pw.pkl'
with file_path.open('rb') as file:
    hashed_passwords = pickle.load(file)

# Configuration de l'authentification
CREDENTIALS = {
    'usernames': {
        'cbri': {'name': 'Christophe Bri', 'password': hashed_passwords.get('cbri')},
        'mpec': {'name': 'Mourad Pec', 'password': hashed_passwords.get('mpec')},
        'elap': {'name': 'Emilie Lap', 'password': hashed_passwords.get('elap')},
        'pgou': {'name': 'Pierre Gou', 'password': hashed_passwords.get('pgou')},
        'osai': {'name': 'Olivier Sai', 'password': hashed_passwords.get('osai')},
        'fsau': {'name': 'Frédéric Sau', 'password': hashed_passwords.get('fsau')},
        'mhum': {'name': 'Morgane Hum', 'password': hashed_passwords.get('mhum')},
        'akes': {'name': 'Archimède Kes', 'password': hashed_passwords.get('akes')},
        'dlau': {'name': 'David Lau', 'password': hashed_passwords.get('dlau')},
        'jdel': {'name': 'Jean Del', 'password': hashed_passwords.get('jdel')},
    }
}

COOKIE_KEY = 'KwCj_9FTM4gwFSf8BWpeIglekI8iqYm3VUbBuLWvdvs'
