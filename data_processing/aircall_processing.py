import pandas as pd
import os
from datetime import datetime, timedelta
import numpy as np
import sqlite3
import hashlib
import pickle


def get_data_hash(data):
    """Calcule un hash des données pour détecter les changements"""
    return hashlib.md5(pd.util.hash_pandas_object(data).values).hexdigest()


def save_processed_data_to_cache(processed_data, cache_key, cache_path="data/Affid/Cache/processed_cache.pkl"):
    """Sauvegarde les données traitées avec leur hash"""
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    cache_data = {
        'data': processed_data,
        'hash': get_data_hash(processed_data),
        'timestamp': datetime.now()
    }
    with open(cache_path, 'wb') as f:
        pickle.dump(cache_data, f)


def load_processed_data_from_cache(cache_key, cache_path="data/Affid/Cache/processed_cache.pkl"):
    """Charge les données traitées depuis le cache"""
    try:
        if os.path.exists(cache_path):
            with open(cache_path, 'rb') as f:
                cache_data = pickle.load(f)
            return cache_data['data']
    except Exception as e:
        print(f"Erreur lors du chargement du cache: {e}")
    return None


def save_to_sqlite(df, table_name, db_path="data/Affid/Cache/cache.sqlite"):
    with sqlite3.connect(db_path) as conn:
        # Créer les index pour les colonnes fréquemment utilisées
        df.to_sql(table_name, conn, if_exists="replace", index=False)
        cursor = conn.cursor()
        cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_date ON {table_name}(Date)")
        cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_semaine ON {table_name}(Semaine)")
        cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_username ON {table_name}(UserName)")
        conn.commit()

def read_from_sqlite(table_name, db_path="data/Affid/Cache/cache.sqlite"):
    # Utiliser une seule connexion pour toute la session
    if not hasattr(read_from_sqlite, 'conn'):
        read_from_sqlite.conn = sqlite3.connect(db_path)
    
    return pd.read_sql(f"SELECT * FROM {table_name}", read_from_sqlite.conn, 
                      parse_dates=['StartTime', 'Date', 'time (TZ offset incl.)', 'HangupTime'])


def load_aircall_data(path_v1, path_v2, force_reload=False, db_path="data/cache.sqlite"):
    table_name = "aircall_processed"
    
    # Vérifier d'abord le cache des données traitées
    if not force_reload:
        cached_data = load_processed_data_from_cache("aircall_processed")
        if cached_data is not None:
            print("✅ Données traitées chargées depuis le cache pickle")
            return cached_data
    
    # Vérifier le cache SQLite
    if not force_reload:
        try:
            df = read_from_sqlite(table_name, db_path)
            print("✅ Données chargées depuis SQLite")
            # Sauvegarder dans le cache pickle pour les prochaines fois
            save_processed_data_to_cache(df, "aircall_processed")
            return df
        except Exception as e:
            print("❌ Échec chargement cache SQLite :", e)

    # Sinon, traitement complet avec optimisations
    print("🔄 Traitement complet des données Aircall...")
    
    # Chargement parallèle des fichiers
    files_v1 = [file for file in os.listdir(path_v1) if not file.startswith('.')]
    files_v2 = [file for file in os.listdir(path_v2) if not file.startswith('.')]

    # Optimisation : charger seulement les colonnes nécessaires
    needed_columns = ['line', 'date (TZ offset incl.)', 'time (TZ offset incl.)', 'number timezone', 
                     'datetime (UTC)', 'country_code', 'direction', 'from', 'to', 'answered',
                     'missed_call_reason', 'user', 'duration (total)', 'duration (in call)', 
                     'via', 'voicemail', 'tags']
    
    data_v1 = pd.concat([pd.read_excel(os.path.join(path_v1, file), usecols=needed_columns) 
                        for file in files_v1])
    data_v2 = pd.concat([pd.read_excel(os.path.join(path_v2, file), usecols=needed_columns) 
                        for file in files_v2])

    # Ajouter la colonne IVR Branch aux deux datasets
    data_v1['IVR Branch'] = ""
    data_v2['IVR Branch'] = ""
    
    # Définir l'ordre des colonnes final
    final_columns = ['line', 'date (TZ offset incl.)', 'time (TZ offset incl.)', 'number timezone', 'datetime (UTC)', 'country_code', 'direction', 'from',
                     'to', 'answered','missed_call_reason', 'user', 'duration (total)','duration (in call)', 'via', 'voicemail', 'tags', 'IVR Branch']
    
    # S'assurer que les deux datasets ont les mêmes colonnes dans le même ordre
    data_v1 = data_v1[final_columns]
    data_v2 = data_v2[final_columns]
    
    raw_data = pd.concat([data_v1, data_v2])

    processed_data = process_aircall_data(raw_data)
    
    # Sauvegarder dans les deux caches
    save_to_sqlite(processed_data, table_name, db_path)
    save_processed_data_to_cache(processed_data, "aircall_processed")
    
    print("✅ Données traitées et sauvegardées dans SQLite et cache pickle")

    return processed_data


def process_aircall_data(data):
    # Copie pour éviter les modifications sur l'original
    data = data.copy()
    
    # Renommage des colonnes
    data.rename(columns={"answered": "LastState", 
                        'date (TZ offset incl.)': "StartTime", 
                        "duration (total)": "TotalDuration", 
                        "duration (in call)": "InCallDuration", 
                        "from": "FromNumber", "to": "ToNumber", 
                        "user": "UserName", 
                        "comments": "Note", 
                        "tags": "Tags", "missed_call_reason": "ScenarioName"}, inplace=True)
    
    # Optimisation : conversions de dates vectorisées
    data['time (TZ offset incl.)'] = pd.to_datetime(data['time (TZ offset incl.)'], format='%H:%M:%S')
    data['StartTime'] = pd.to_datetime(data['StartTime'])
    data['HangupTime'] = data['time (TZ offset incl.)'] + pd.to_timedelta(data['InCallDuration'], unit='s')
    
    # Calculs vectorisés
    data['Semaine'] = data['StartTime'].dt.strftime("S%Y-%V")
    data['Heure'] = data['time (TZ offset incl.)'].dt.hour
    data['Date'] = data['StartTime'].dt.normalize()
    data['Jour'] = data['StartTime'].dt.day_name()
    
    # Filtrage vectorisé
    weekend_mask = ~data["Jour"].isin(["Saturday", "Sunday"])
    scenario_mask = ~data["ScenarioName"].isin(["Fermé", "out_of_opening_hours", "abandoned_in_ivr", 'short_abandoned'])
    data = data[weekend_mask & scenario_mask]
    
    # Mapping vectorisé pour LastState
    state_mapping = {"ANSWERED": "yes", "VOICEMAIL": "no", "MISSED": "no", 
                    "VOICEMAIL_ANSWERED": "no", "BLIND_TRANSFERED": "no", 
                    "NOANSWER_TRANSFERED": "no", "FAILED": "no", "CANCELLED": "no", 
                    "QUEUE_TIMEOUT": "no", "yes": "yes", "no": "no", 
                    "Yes": "yes", "No": "no"}
    data['LastState'] = data['LastState'].map(state_mapping)
    
    # Optimisation : nettoyage des tags vectorisé
    data['Tags'] = data['Tags'].astype(str).str.replace('[^a-zA-Z-,]', '', regex=True)
    data['NRP'] = 'no'
    
    # Condition vectorisée pour NRP
    nrp_mask = (data['Tags'].isin(['NRP'])) & (data['direction'] == 'outbound')
    data.loc[nrp_mask, 'LastState'] = data.loc[nrp_mask, 'NRP']
    
    # Sélection des colonnes finales
    final_columns = ['line', 'Semaine', 'Date', 'Jour', 'Heure', 'direction', 'LastState', 
                    'ScenarioName', 'StartTime', 'HangupTime', 'time (TZ offset incl.)', 
                    'TotalDuration', 'InCallDuration', 'FromNumber', 'ToNumber', 
                    'UserName', 'Tags', 'IVR Branch']
    data = data[final_columns]
    
    # Remplacements vectorisés pour les noms d'utilisateurs
    data['UserName'] = data['UserName'].str.replace("Archimède KESSI", "Archimede KESSI")
    data['UserName'] = data['UserName'].str.replace("Olivier SAINTE-ROSE", "Olivier Sainte-Rose")
    
    # Filtrage par date vectorisé
    today = pd.Timestamp.today()
    week_prior = today - pd.Timedelta(weeks=50)
    data = data[data['Date'] >= week_prior]
    
    # Tri final
    data = data.sort_values(by='Semaine', ascending=True)
    
    return data


agents = ['Olivier Sainte-Rose', 
        'Mourad HUMBLOT', 
        'Pierre GOUPILLON',
        'Frederic SAUVAN', 
        'Christophe Brichet']

frederic = ['Frederic SAUVAN']

agents_support = ['Olivier Sainte-Rose', 
                'Mourad HUMBLOT', 
                'Pierre GOUPILLON', 
                'Archimede KESSI', 
                'Frederic SAUVAN', 
                'Christophe Brichet']

agents_armatis = ['Melinda Marmin', 
                'Sandrine Sauvage', 
                'Emilie GEST', 
                'Morgane Vandenbussche']

agents_all = [ 'Melinda Marmin',
                'Sandrine Sauvage', 
                'Emilie GEST', 
                'Morgane Vandenbussche',
                'Olivier Sainte-Rose', 
                'Mourad HUMBLOT', 
                'Pierre GOUPILLON', 
                'Archimede KESSI', 
                'Frederic SAUVAN', 
                'Christophe Brichet',
                'Celine Crendal']


line_support = 'technique'
line_armatis = 'armatistechnique'
line_xmed = 'xmed'
line_tous = 'tous'


def def_df_support(df_entrant, df_sortant, line, liste_agents):
    def clean_string(s):
        return ''.join(s.split()).lower()

    # S'assurer que les dates sont au bon format
    df_entrant = df_entrant.copy()
    df_sortant = df_sortant.copy()
    
    df_entrant['Date'] = pd.to_datetime(df_entrant['Date']).dt.normalize()
    df_sortant['Date'] = pd.to_datetime(df_sortant['Date']).dt.normalize()

    df_entrant['line'] = df_entrant['line'].apply(clean_string)
    df_sortant['line'] = df_sortant['line'].apply(clean_string)

    # Filtrage vectorisé
    if line == "tous":
        entrant_mask = (df_entrant['line'].isin(['technique', 'armatistechnique', 'xmed'])) & (df_entrant['direction'] == 'inbound')
        df_entrant = df_entrant[entrant_mask]
    elif line in ['technique', 'armatistechnique', 'xmed']:
        entrant_mask = (df_entrant['line'] == line) & (df_entrant['direction'] == 'inbound')
        df_entrant = df_entrant[entrant_mask]

    sortant_mask = (df_sortant['UserName'].isin(liste_agents)) & (df_sortant['direction'] == 'outbound')
    df_sortant = df_sortant[sortant_mask]

    df_entrant['Number'] = df_entrant['FromNumber']
    df_sortant['Number'] = df_sortant['ToNumber']

    df = pd.concat([df_entrant, df_sortant])

    # Filtrage vectorisé
    weekend_mask = ~df["Jour"].isin(["Saturday", "Sunday"])
    user_mask = ~df["UserName"].isin(["Vincent Gourvat", "Thierry CAROFF", 'Armatis Agent 1'])
    df = df[weekend_mask & user_mask]

    # Calculs vectorisés
    df['Count'] = 1
    df['Entrant_connect'] = ((df['LastState'] == 'yes') & (df['direction'] == 'inbound')).astype(int)
    df['Entrant'] = (df['direction'] == 'inbound').astype(int)
    df['Sortant_connect'] = ((df['direction'] == 'outbound') & (df['InCallDuration'] > 60)).astype(int)
    df['Taux_de_service'] = df['Entrant_connect'] / df['Entrant']
    
    df["Mois"] = df['StartTime'].dt.strftime("%Y-%m")

    # Optimisation : calcul de l'effectif avec groupby vectorisé
    df_grouped = df.groupby(['Date', 'UserName']).size().reset_index(name='TotalAppels')
    df_grouped['Actif'] = (df_grouped['TotalAppels'] >= 2).astype(int)

    # Calculer l'effectif moyen par jour
    df_effectif = df_grouped.groupby('Date')['Actif'].sum().reset_index()
    df_effectif.rename(columns={'Actif': 'Effectif'}, inplace=True)

    # Fusionner l'effectif avec le DataFrame principal
    df = pd.merge(df, df_effectif, on='Date', how='left')

    # Optimisation : fonction vectorisée pour get_ivr_or_tags_transformed
    def get_ivr_or_tags_transformed_vectorized(df):
        # Créer un masque pour chaque condition
        ivr_mask = (df['IVR Branch'].notna()) & (df['IVR Branch'].str.strip() != '')
        armatis_mask = df['line'] == 'armatistechnique'
        
        # Initialiser la colonne avec 'Inconnu'
        logiciel = pd.Series('Inconnu', index=df.index)
        
        # Appliquer les conditions
        logiciel[ivr_mask] = df.loc[ivr_mask, 'IVR Branch']
        logiciel[armatis_mask] = 'Stellair'
        
        # Traitement des tags
        tags_mask = df['Tags'].notna()
        tags_prefix = df.loc[tags_mask, 'Tags'].str[:3].str.upper()
        
        stellair_tags = tags_prefix == 'STE'
        affid_tags = tags_prefix == 'AFD'
        
        logiciel.loc[tags_mask & stellair_tags] = 'Stellair'
        logiciel.loc[tags_mask & affid_tags] = 'Affid'
        
        return logiciel

    df['Logiciel'] = get_ivr_or_tags_transformed_vectorized(df)

    return df
