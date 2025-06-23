import pandas as pd
import os
from datetime import datetime, timedelta

import pandas as pd
import os
import sqlite3

def load_hubspot_data(ticket_path, db_path="data/Cache/data_cache.db"):
    """
    Charge les données de tickets depuis les fichiers Excel si nécessaire,
    puis lit depuis la base SQLite.
    """

    # 1. Définir la table cible
    table_name = "df_tickets"

    # 2. Vérifie si un des fichiers source est plus récent que la table SQL
    excel_files = [os.path.join(ticket_path, f) for f in os.listdir(ticket_path) if f.endswith((".xls", ".xlsx")) and not f.startswith(".")]
    
    if not os.path.exists(db_path):
        needs_refresh = True
    else:
        db_mtime = os.path.getmtime(db_path)
        needs_refresh = any(os.path.getmtime(f) > db_mtime for f in excel_files)

    # 3. Si besoin, recharge les fichiers Excel et les insère dans SQLite
    if needs_refresh:
        print("Mise à jour des données tickets depuis Excel vers SQLite...")
        dfs = [pd.read_excel(f) for f in excel_files]
        df_raw = pd.concat(dfs, ignore_index=True)
        
        from data_processing.hubspot_processing import process_hubspot_data
        df_processed = process_hubspot_data(df_raw)

        for col in df_processed.select_dtypes(include="object").columns:
            df_processed[col] = df_processed[col].fillna("").astype(str)

        # ✅ Crée le dossier 'data/Cache/' s'il n'existe pas
        os.makedirs(os.path.dirname(db_path), exist_ok=True)

        conn = sqlite3.connect(db_path)
        df_processed.to_sql(table_name, conn, if_exists="replace", index=False)
        conn.close()

    # 4. Lecture finale depuis SQLite
    conn = sqlite3.connect(db_path)
    df = pd.read_sql(f"SELECT * FROM {table_name}", conn, parse_dates=["Date de création"])
    conn.close()
    data_ticket = df

    return data_ticket


def process_hubspot_data(data_ticket):
    from datetime import timedelta

    # Fonction pour parser les dates avec gestion des formats différents
    def parse_date_safely(date_series, column_name):
        """
        Parse les dates en gérant différents formats possibles
        """
        # Essayer d'abord le format YYYY-MM-DD (format standard)
        try:
            parsed = pd.to_datetime(date_series, format='%Y-%m-%d %H:%M:%S', errors='coerce')
            if parsed.notna().sum() > 0:
                print(f"Format YYYY-MM-DD reconnu pour {column_name}")
                return parsed
        except:
            pass
        
        # Essayer le format DD/MM/YYYY avec heures
        try:
            parsed = pd.to_datetime(date_series, format='%d/%m/%Y %H:%M', errors='coerce')
            if parsed.notna().sum() > 0:
                print(f"Format DD/MM/YYYY reconnu pour {column_name}")
                return parsed
        except:
            pass
        
        # Essayer le format DD/MM/YYYY sans heures
        try:
            parsed = pd.to_datetime(date_series, format='%d/%m/%Y', errors='coerce')
            if parsed.notna().sum() > 0:
                print(f"Format DD/MM/YYYY (sans heures) reconnu pour {column_name}")
                return parsed
        except:
            pass
        
        # Essayer avec dayfirst=True pour le format DD-MM-YYYY
        try:
            parsed = pd.to_datetime(date_series, dayfirst=True, errors='coerce')
            if parsed.notna().sum() > 0:
                print(f"Format dayfirst=True reconnu pour {column_name}")
                return parsed
        except:
            pass
        
        # En dernier recours, laisser pandas deviner
        print(f"Utilisation du parsing automatique pour {column_name}")
        return pd.to_datetime(date_series, errors='coerce')

    # Conversion des colonnes de dates avec gestion des formats
    data_ticket['Date de création'] = parse_date_safely(data_ticket['Date de création'], 'Date de création')
    data_ticket["Date de la première réponse par e-mail de l'agent"] = parse_date_safely(
        data_ticket["Date de la première réponse par e-mail de l'agent"], 
        "Date de la première réponse par e-mail de l'agent"
    )

    def calculate_working_hours(start, end):
        start_hour, end_hour = 9, 18
        total_working_hours = 0
        current = start
        while current < end:
            next_day = (current + timedelta(days=1)).normalize()
            if current.weekday() < 5:
                work_start = current.replace(hour=start_hour, minute=0, second=0)
                work_end = current.replace(hour=end_hour, minute=0, second=0)
                day_start = max(current, work_start)
                day_end = min(end, work_end, next_day)
                if day_start < day_end:
                    total_working_hours += (day_end - day_start).total_seconds() / 3600
            current = next_day
        return total_working_hours

    data_ticket['working_hours'] = data_ticket.apply(
        lambda row: calculate_working_hours(row['Date de création'], row["Date de la première réponse par e-mail de l'agent"]),
        axis=1
    )

    def convert_hours_to_hms(hours):
        total_seconds = int(hours * 3600)
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    data_ticket['working_hours_hms'] = data_ticket['working_hours'].apply(convert_hours_to_hms)
    data_ticket["Semaine"] = data_ticket['Date de création'].dt.strftime("S%Y-%V")
    data_ticket["Date"] = data_ticket["Date de création"].dt.date
    data_ticket['Nombre_ticket_client'] = data_ticket.apply(
        lambda x: 1 if x['Statut du ticket'] and x['Source'] in ['E-mail', 'Chat', 'Formulaire'] else 0,
        axis=1
    )
    data_ticket['Nombre_ticket_telephone'] = data_ticket.apply(
        lambda x: 1 if x['Statut du ticket'] and x['Source'] in ['Téléphone'] else 0,
        axis=1
    )

    # Renommage des agents
    agent_rename_map = {
        "Archimède KESSI": "Archimede KESSI",
        "Emilie Gest": "Emilie GEST",
        "HUMBLOT NASSUF": "Mourad HUMBLOT",
        "FREDERIC SAUVAN": "Frederic SAUVAN",
        "Morgane VANDENBUSSCHE": "Morgane Vandenbussche",
        "Pierre Goupillon": "Pierre GOUPILLON"
    }
    data_ticket["Propriétaire du ticket"] = data_ticket["Propriétaire du ticket"].replace(agent_rename_map)

    return data_ticket

