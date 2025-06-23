import pandas as pd
import numpy as np
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px




def charge_entrant_sortant(df_support, agents):
    import pandas as pd

    if isinstance(agents, str):
        agents = [agents]

    # Filtrage par les agents et par l'état 'yes'
    df_filtered = df_support[
        (df_support['UserName'].isin(agents)) &
        (df_support['LastState'] == 'yes') &
        (df_support['direction'] == 'inbound')  # On se concentre sur les appels entrants
    ]

    # Groupement par Semaine, Agent, et Direction
    df_grouped = df_filtered.groupby(['Semaine', 'UserName']).agg({'Date': 'count'}).reset_index()

    # Étiquette série = Agent pour la courbe
    df_grouped['Série'] = df_grouped['UserName']

    # Calcul des semaines dans le bon ordre
    ordre_semaines = sorted(df_grouped['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))

    # 📈 Graphique en ligne
    fig_line = px.line(
        df_grouped,
        x='Semaine',
        y='Date',
        color='Série',
        markers=True,
        title='Volume d’appels entrants par agent'
    )

    fig_line.update_layout(
        yaxis_title="Nombre d'appels validés",
        xaxis=dict(categoryorder="array", categoryarray=ordre_semaines),
        yaxis=dict(range=[0, df_grouped['Date'].max() + 10])
    )

    # 🥧 Graphique camembert (somme par agent)
    df_pie = df_grouped.groupby('UserName')['Date'].sum().reset_index()
    df_pie = df_pie.rename(columns={'Date': 'Total appels'})

    fig_pie = px.pie(
        df_pie,
        names='UserName',
        values='Total appels',
        title="Répartition des appels entrants par agent"
    )

    return fig_line, fig_pie



def charge_ticket(df_ticket, agents):
    import pandas as pd

    if isinstance(agents, str):
        agents = [agents]

    # Filtrage par les agents, le pipeline et la source
    df_filtered = df_ticket[
        (df_ticket['Propriétaire du ticket'].isin(agents)) &
        (df_ticket['Pipeline'].isin(['SSIA', 'SSI', 'SPSA'])) &
        (df_ticket['Source'].isin(['Chat', 'E-mail', 'Formulaire']))
    ]

    # Groupement par semaine et agent
    df_grouped = df_filtered.groupby(['Semaine', 'Propriétaire du ticket']).agg({'Date': 'count'}).reset_index()
    df_grouped['Série'] = df_grouped['Propriétaire du ticket']

    # Ordre des semaines
    ordre_semaines = sorted(df_grouped['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))

    # 📈 Graphique ligne
    fig_line = px.line(
        df_grouped,
        x='Semaine',
        y='Date',
        color='Série',
        markers=True,
        title='Volume de tickets par agent'
    )

    fig_line.update_layout(
        yaxis_title="Nombre de tickets",
        xaxis=dict(categoryorder="array", categoryarray=ordre_semaines),
        yaxis=dict(range=[0, df_grouped['Date'].max() + 10])
    )

    # 🥧 Graphique camembert
    df_pie = df_grouped.groupby('Propriétaire du ticket')['Date'].sum().reset_index()
    df_pie = df_pie.rename(columns={'Date': 'Total tickets'})

    fig_pie = px.pie(
        df_pie,
        names='Propriétaire du ticket',
        values='Total tickets',
        title="Répartition des tickets par agent"
    )

    return fig_line, fig_pie




def metrics_nombre_ticket_pipeline_agent(df_hubspot, agents):
    df_hubspot['Date'] = pd.to_datetime(df_hubspot['Date'])
    
    # Définitions des groupes d’agents
    support_agents = ['Archimède KESSI', 'HUMBLOT NASSUF','Olivier Sainte-Rose']
    armatis_agents = ['Emilie Gest', 'Sandrine Sauvage', 'Melinda Marmin', 'Morgane VANDENBUSSCHE']
    
    # Cas spéciaux : groupes prédéfinis
    if agents == 'agents_support':
        selected_agents = support_agents
    elif agents == 'agents_armatis':
        selected_agents = armatis_agents
    elif agents == 'agents_all':
        selected_agents = support_agents + armatis_agents
    else:
        # Si un ou plusieurs agents sont fournis manuellement
        if isinstance(agents, str):
            selected_agents = [agents]
        else:
            selected_agents = agents
    
    # Filtrage
    df_filtered = df_hubspot[
        (df_hubspot['Propriétaire du ticket'].isin(selected_agents)) &
        ((df_hubspot['Source'].isin(['Chat', 'E-mail', 'Formulaire', ''])) | pd.isna(df_hubspot['Source']))
    ]
    
    # Regroupement et agrégation
    df_grouped = df_filtered.groupby("Date").agg(Tickets=('Ticket ID', 'nunique')).reset_index()
    
    return df_grouped['Tickets'].mean() if not df_grouped.empty else 0


def filtrer_par_agent(agent):
    # Dictionnaire associant les noms courts aux noms complets
    correspondance_agents = {
        "Mourad": ["Mourad HUMBLOT", "HUMBLOT NASSUF"],
        "Olivier": ["Olivier Sainte-Rose"],
        "Archimède": ["Archimede KESSI", "Archimède KESSI"],
        "Frederic": ["Frederic SAUVAN", "FREDERIC SAUVAN"],
        "Pierre": ["Pierre GOUPILLON", "Pierre Goupillon"],
        "Christophe": ["Christophe Brichet"],
        "Emilie" : ["Emilie Gest", 'Emilie GEST'], 
        "Morgane" : ["Morgane VANDENBUSSCHE", 'Morgane Vandenbussche'],
        "Melinda" : ["Melinda Marmin"], 
        "Sandrine" : ["Sandrine Sauvage"], 
        "Celine": ['Céline Crendal','Celine Crendal']
    }
    
    # Si l'agent est dans la liste, on filtre, sinon on renvoie tout le DataFrame
    #agent = correspondance_agents[agent]

    if agent in correspondance_agents:
        noms_complets = correspondance_agents[agent]
        
    return noms_complets


def calcul_productivite_appels(df_support, agent):
    # Filtre les données où la durée de l'appel est supérieure à 45
    df_support_all = df_support
    df_support = df_support[df_support['InCallDuration'] > 45]

    # Groupement des appels par date et agent (UserName)
    df_grouped = df_support.groupby(['Date', 'UserName']).size().reset_index(name='TotalAppels')
    df_grouped = df_grouped[df_grouped['UserName'].isin(agent)].groupby(['Date']).agg({'TotalAppels':'mean'})

    # Groupement des appels par agent(UserName)
    df_grouped_total = df_support.groupby(['UserName']).size().reset_index(name='TotalAppels')
    df_grouped_total = df_grouped_total[df_grouped_total['UserName'].isin(agent)].agg({'TotalAppels':'mean'})

    # Filtrage des appels par agent et par date
    df_support = df_support[df_support['UserName'].isin(agent)].groupby(['Date']).agg({'InCallDuration':'sum'})

    # Fusionner les deux dataframes sur la colonne 'Date'
    df_support = pd.merge(df_support, df_grouped, on='Date', how='left')
    
    def convert_minutes_to_hhmmss(minutes):
        # Convertir les minutes en heures, minutes et secondes
        hours = int(minutes // 3600)
        minutes = int(minutes % 60)
        seconds = int((minutes % 1) * 60)
        # Formater en hh:mm:ss
        return f"{hours:02}:{minutes:02}:{seconds:02}"
    
    df_support['InCallDuration_format'] = df_support['InCallDuration'].apply(convert_minutes_to_hhmmss)
    df_support['InCallDuration_format'] = pd.to_datetime(df_support['InCallDuration_format'])

    # Calcul de la moyenne des durées d'appel par jour
    com_jour = df_support['InCallDuration_format'].mean()

    # Vérifier si la valeur de com_jour est valide
    if pd.notnull(com_jour):  # Vérifie si com_jour n'est pas NaT
        com_jour = com_jour.strftime('%H:%M:%S')
    else:
        com_jour = None  # Ou une autre valeur par défaut, comme '' ou '00:00:00'

    # Calcul du temps moyen de communication par appel
    temps_moy_com = (df_support['InCallDuration'] / df_support['TotalAppels']).mean()

    # Nombre moyen d'appels par jour
    nb_appels_jour = df_support['TotalAppels'].mean()
    nb_appels = df_grouped_total['TotalAppels'].mean()

    return com_jour, temps_moy_com, nb_appels_jour, nb_appels

def kpi_agent(agent, df_support):

    if isinstance(agent, str):
        agent = [agent]
    # Filtrer une seule fois pour l'agent donné
    df_agent = df_support[df_support['UserName'].isin(agent)]
    
    # Calcul des KPI de productivité
    com_jour, temps_moy_com, nb_appels_jour, nb_appel = calcul_productivite_appels(df_agent, agent)

    # Calcul des appels entrants et sortants
    nb_appels_jour_entrants = df_agent[df_agent['direction'] == 'inbound']['Entrant_connect'].sum()
    nb_appels_jour_sortants = df_agent[df_agent['direction'] == 'outbound']['Sortant_connect'].sum()

    # Éviter une division par zéro pour le ratio
    total_appels = nb_appels_jour_entrants + nb_appels_jour_sortants
    if total_appels > 0:
        ratio_entrants = nb_appels_jour_entrants / total_appels
        ratio_sortants = nb_appels_jour_sortants / total_appels
    else:
        ratio_entrants = 0
        ratio_sortants = 0

    return com_jour, temps_moy_com, nb_appels_jour, nb_appel, nb_appels_jour_entrants, nb_appels_jour_sortants, ratio_entrants, ratio_sortants

def convert_to_sixtieth(seconds):
    if pd.isnull(seconds):  # Vérifie si seconds est NaN
        return "Invalid"  # Retourne une valeur par défaut ou un message d'erreur
    minutes, seconds = divmod(int(seconds), 60)  # Convertir en heures
    return f"{int(minutes)}m{int(seconds):02d}s"

def filtrer_par_periode(df_support, periode):
    """ Filtre les données en fonction de la période sélectionnée """
    df_support['Date'] = pd.to_datetime(df_support['Date'])  # Assure que 'Date' est bien en datetime
    derniere_date = df_support['Date'].max()

    if periode == "6 derniers mois":
        date_limite = derniere_date - pd.DateOffset(months=6)
    elif periode == "3 derniers mois":
        date_limite = derniere_date - pd.DateOffset(months=3)
    elif periode == "Dernier mois":
        date_limite = derniere_date - pd.DateOffset(months=1)
    else:
        date_limite = df_support['Date'].min()  # Toute la période

    return df_support[df_support['Date'] >= date_limite]

def graph_activite(df_support):
    from plotly.subplots import make_subplots
    import plotly.graph_objects as go

    # Filtrage des données pour l'activité entrante
    data_graph1 = df_support[df_support['direction'] == 'inbound']

    # Agrégation des données par semaine et date
    data_graph2 = (
        data_graph1.groupby(['Semaine'])
        .agg(
            Entrant=('Entrant', 'sum'),
            Entrant_connect=('Entrant_connect', 'sum'),
            Numero_unique=('Number', 'nunique'),
            Effectif=('Effectif', 'mean')
        )
        .reset_index()
    )

    # Agrégation finale par semaine
    data_graph3 = (
        data_graph2.groupby('Semaine')[['Entrant', 'Entrant_connect', 'Numero_unique', 'Effectif']]
        .mean()
        .reset_index()
    )

    # Calcul du taux de service support
    data_graph3['Taux_de_service_support'] = (
        data_graph3['Entrant_connect'] / data_graph3['Entrant']
    )

    # Ajout d'une colonne 100%
    data_graph3['100%'] = 1

    # Création de la figure avec des sous-graphiques
    fig = make_subplots(rows=1, cols=1, specs=[[{"secondary_y": True}]])

    # Ajouter les barres pour le Taux de service support
    fig.add_trace(
        go.Bar(
            x=data_graph3['Semaine'],
            y=data_graph3['Taux_de_service_support'],
            name='Taux',
            opacity=0.7,
            text=data_graph3['Taux_de_service_support'],
            texttemplate='%{text:.0%}'
        ),
        secondary_y=True,
    )

    # Ajouter les lignes empilées pour Numero_unique, Entrant_connect et Entrant
    fig.add_trace(
        go.Scatter(
            x=data_graph3['Semaine'],
            y=data_graph3['Numero_unique'],
            name='Numero_unique',
            fill='tozeroy'
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=data_graph3['Semaine'],
            y=data_graph3['Entrant_connect'],
            name='Entrant_connect',
            fill='tozeroy'
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=data_graph3['Semaine'],
            y=data_graph3['Entrant'],
            name='Entrant',
            fill='tozeroy'
        ),
        secondary_y=False,
    )

    # Mise à jour des axes Y pour les pourcentages
    fig.update_yaxes(range=[0, 1], secondary_y=True)
    fig.update_yaxes(tickformat=".0%", secondary_y=True)

    events = {
        "S2024-03": "Stellair HS",
        "S2024-51": "Mise à jour réglementaire",
        "S2025-01": "Effectif : 2",
        "S2025-05": "Stellair HS",
        "S2025-06": "Stellair HS",
        "S2025-07": "Stellair HS",
        "S2025-09": "Stellair HS",
        "S2025-10": "Stellair HS",
        "S2025-13": "Stellair HS",
    }

    # Préparation des listes pour les bulles
    x_events = []
    y_events = []
    text_events = []

    # Pour positionner la bulle, nous récupérons la valeur maximale parmi les séries 'Entrant', 'Entrant_connect' et 'Numero_unique'
    for week, event_text in events.items():
        if week in data_graph3['Semaine'].values:
            # Récupération de la ligne correspondante
            row = data_graph3[data_graph3['Semaine'] == week].iloc[0]
            # On définit la position verticale comme le maximum des trois séries
            y_val = row['Taux_de_service_support']
            marge = 0.05  
            y_val_event = min(y_val + marge, 1)
            x_events.append(week)
            y_events.append(y_val_event)
            text_events.append(event_text)

    # Ajout du trace pour les bulles d'événements
    fig.add_trace(
        go.Scatter(
            x=x_events,
            y=y_events,
            mode="markers+text",
            marker=dict(
                size=15,
                color="red"
            ),
            text=text_events,
            textposition="top center",
            name="Événements"
        ),
        secondary_y=True,
    )

    # Configuration de la mise en page et des légendes
    fig.update_layout(
        title='Graphique avec Taux en barres et Numero_unique/Entrant en aires empilées',
        template='plotly_dark',
        xaxis_title='Semaine',
        yaxis_title='Valeurs',
        title_text="Activité & Taux de service - 20 semaines",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )

    return fig

def graph_taux_jour (df_support): 

    df_support  = df_support[df_support['direction'] == 'inbound']


    data_graph2 = df_support.groupby(['Semaine','Date', 'Jour']).agg({'Entrant':'sum',
                                                                'Entrant_connect':'sum',
                                                                'Number':'nunique',
                                                                'Effectif':'mean'})
    
    data_graph2 = data_graph2.groupby(['Semaine','Date', 'Jour']).agg({'Entrant':'mean',
                                                    'Entrant_connect':'mean',
                                                    'Number':'mean',
                                                    'Effectif':'mean'}).rename(columns = {"Number":"Numero_unique",})
    
    data_graph2['Taux_de_service_support'] = (data_graph2['Entrant_connect'] 
                                            / data_graph2['Entrant'])

    data_graph2 = data_graph2.reset_index()

    data_moyenne = data_graph2.groupby("Jour")[["Taux_de_service_support", "Entrant"]].mean().reset_index()


    derniere_semaine = data_graph2['Semaine'].iloc[-1]
    data_derniere_semaine = data_graph2.loc[(data_graph2['Semaine'] == derniere_semaine)]


    #def graph_taux_semaine():
        
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    last_week = [data_graph2.tail(1).values[0,0]]


    fig.add_trace(
        go.Bar(x=data_moyenne['Jour'], y=data_moyenne['Taux_de_service_support'], name="Taux de service"),
        secondary_y=False,)

    fig.add_trace(
        go.Scatter(x=data_moyenne['Jour'], y=data_moyenne['Entrant'], mode='lines', name="Entrant"), 
        secondary_y=True,)

    
    fig.update_yaxes(range = [0,1], secondary_y=False)
    fig.update_yaxes(range=[20,120], secondary_y=True)
    fig.update_yaxes(tickformat=".0%", secondary_y=False)


    fig.update_layout(title='Graphique avec Taux en barres et Numero_unique/Entrant en aires empilées',
                template='plotly_dark',
                xaxis_title='Semaine',
                yaxis_title='Valeurs',
                )


    fig.update_layout(
        title_text="Activité & Taux de service / jour",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )

       
    
    return fig

def graph_taux_heure (df_support): 

    data_by_heure = df_support.groupby(['Heure']).agg({'Entrant':'sum', 
                                                   'Entrant_connect':'sum'}).query("Heure > 8 & Heure < 18 != 12").reset_index()


    data_by_heure['Taux_de_service_support'] = data_by_heure['Entrant_connect'] / data_by_heure['Entrant']


    #def graph_taux_heure(): 
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Scatter(x=data_by_heure['Heure'], y=data_by_heure['Entrant'], name="Entrant"),
        secondary_y=True,
    )

    fig.add_trace(
        go.Bar(x=data_by_heure['Heure'], y=data_by_heure['Taux_de_service_support'], name="Taux de service"),
        secondary_y=False,
    )


    # Formater les étiquettes de l'axe y en pourcentages
    fig.update_yaxes(range = [0,1], secondary_y=False)
    fig.update_yaxes(tickformat=".0%", secondary_y=False)
    #fig.update_yaxes(range = [50,150], secondary_y=True)

    fig.update_layout(title='Graphique avec Taux en barres et Numero_unique/Entrant en aires empilées',
                template='plotly_dark',
                xaxis_title='Semaine',
                yaxis_title='Valeurs',
                )


    fig.update_layout(
        title_text="Activité & Taux de service / Heure",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )

    return fig


def compute_ticket_appels_metrics(df_tickets, df_support, agent):
    # Définir les listes d'agents sans doublons
    agents_n1 = list(set([
        'Archimede KESSI', 'Mourad HUMBLOT', 'Olivier Sainte-Rose',
         'Morgane Vandenbussche'
    ]))

    agents_n1_ticket = list(set([
            'Olivier Sainte-Rose','Mourad HUMBLOT', 'Archimede KESSI', 'Morgane Vandenbussche', 
            'Frederic SAUVAN']))
    
    # Traitement spécial si agent == "agents_all"
    if agent == "agents_all":
        selected_agent = agents_n1
    else:
        selected_agent = [agent] if isinstance(agent, str) else agent

    agents_support = [
        'Archimede KESSI', 'Mourad HUMBLOT', 'Olivier Sainte-Rose'
    ]
    agents_armatis = [
        'Morgane Vandenbussche', 'Emilie GEST', 'Sandrine Sauvage', 'Melinda Marmin'
    ]

    # Filtrage des données support
    df_support_agent = df_support[df_support['UserName'].isin(selected_agent)]
    
    # Vérifications si df_support et df_tickets ne sont pas vides
    if df_support_agent.empty or df_tickets.empty:
        return [0] * 14  # Retourne un tableau avec des valeurs nulles (0) si les données sont vides

    # Calcul des métriques de base
    Nombre_appel_jour_agent = df_support_agent['UserName'].nunique()
    Entrant = df_support_agent.groupby('Date').agg({'Entrant': 'sum'}).mean().values[0] if not df_support_agent.empty else 0
    Numero_unique = df_support_agent[df_support_agent['direction'] == 'inbound'].groupby('Date').agg({'Number': 'nunique'}).mean().values[0] if not df_support_agent.empty else 0
    temps_moy_appel = df_support_agent[df_support_agent['InCallDuration'] > 0].InCallDuration.mean() if not df_support_agent.empty else 0

    # Filtrage des tickets
    df_filtered_agent = df_tickets[
        (df_tickets['Propriétaire du ticket'].isin(selected_agent)) &
        (df_tickets['Pipeline'].isin(['SSIA', 'SSI', 'SPSA'])) &
        (df_tickets['Source'].isin(['Chat', 'E-mail', 'Formulaire']) | pd.isna(df_tickets['Source']))
    ]

    df_filtered_pipeline = df_tickets[
        (df_tickets['Pipeline'].isin(['SSI', 'SSIA','SPSA'])) &
        (df_tickets['Source'].isin(['Chat', 'E-mail', 'Formulaire']) | pd.isna(df_tickets['Source']))
    ]

    # Calcul des métriques liées aux tickets
    total_tickets_traite = df_filtered_agent['Ticket ID'].nunique() if not df_filtered_agent.empty else 0

    # Moyennes et autres métriques avec vérification pour éviter la division par zéro
    moyenne_tickets_total = df_tickets.groupby(['Date'])['Ticket ID'].nunique().mean() if not df_tickets.empty else 0
    moyenne_appels_entrant_total = df_support_agent[df_support_agent['direction'] == 'inbound'].groupby('Date')['Count'].sum().mean() if not df_support_agent.empty else 0
    moyenne_tickets_traite_global = df_filtered_agent.groupby(['Date', 'Propriétaire du ticket'])['Ticket ID'].nunique().mean() if not df_filtered_agent.empty else 0

    # Calcul de l'effectif des tickets et des appels
    df_ticket_effectif = df_tickets.groupby('Date')['Propriétaire du ticket'].nunique().mean() if not df_tickets.empty else 0
    df_ticket_ref = df_filtered_pipeline.groupby(['Date', 'Propriétaire du ticket']).size().reset_index(name='TotalAppels')
    df_ticket_std = df_ticket_ref['TotalAppels'].std() if not df_ticket_ref.empty else 0
    df_ticket_mean = df_ticket_ref['TotalAppels'].mean() if not df_ticket_ref.empty else 0
    df_ticket_ref = df_ticket_std + df_ticket_mean

    df_appel_ref = df_support.groupby(['Date', 'UserName']).agg({'Entrant_connect': 'sum', 'Sortant_connect': 'sum'})
    df_appel_std = df_appel_ref.std().sum() if not df_appel_ref.empty else 0
    df_appel_mean = df_appel_ref.mean().sum() if not df_appel_ref.empty else 0
    df_appel_ref = df_appel_std + df_appel_mean

    # Moyenne des appels traités par agent
    moyenne_appels_traite_agent = df_support_agent.groupby('Date')[['Entrant_connect', 'Sortant_connect']].sum().mean().mean() if not df_support_agent.empty else 0

    # Moyenne des appels entrants par agent
    moyenne_appels_entrant_agent = df_support_agent.groupby('Date')[['Entrant_connect']].sum().mean().mean() if not df_support_agent.empty else 0

    # Calcul des tickets traités par agent et les moyennes
    moyenne_tickets_traite_agent = df_filtered_agent.groupby('Date')['Ticket ID'].nunique().mean() if not df_filtered_agent.empty else 0

    total_tickets_moyenne = df_filtered_agent.groupby(['Date'])['Ticket ID'].nunique().mean() if not df_filtered_agent.empty else 0

    # Calcul de la moyenne des appels traités globalement
    moyenne_appels_traite_global = df_support.groupby(['Date'])[['Entrant_connect', 'Sortant_connect']].sum().mean().mean() if not df_support.empty else 0

    return (
        int(total_tickets_moyenne),
        int(moyenne_appels_entrant_total),
        int(moyenne_tickets_total),
        int(moyenne_appels_traite_global),  # Ajout de la variable ici
        round(moyenne_tickets_traite_global, 2),
        round(moyenne_tickets_traite_agent, 2),
        int(moyenne_appels_traite_agent), 
        int(moyenne_appels_entrant_agent),
        round(df_ticket_effectif, 2),
        total_tickets_traite,
        df_ticket_ref,
        df_appel_ref
    )




def score_ticket(row):
    score = row['Nombre de ticket traité'] / row['ref_ticket_agent']

    return score


def score_appel(row):
    score = row["Nombre d'appel traité"] / row['ref_appel_agent']

    if row["% appel entrant agent"] > 0.5:
        score + 0.15    

    return score



import numpy as np  # à ajouter en début de script si ce n’est pas déjà fait

def historique_scores_total(agents_n1, df_tickets, df_support, date_debut=None, nb_semaines=8):
    historique = []

    if date_debut is None:
        max_date = pd.to_datetime(df_tickets['Date']).max().normalize()
    else:
        max_date = pd.to_datetime(date_debut).normalize()

    for i in range(nb_semaines):
        start_of_week = max_date - pd.Timedelta(weeks=i)
        start_of_week = start_of_week - pd.Timedelta(days=start_of_week.weekday())
        end_of_week = start_of_week + pd.Timedelta(days=6)
        week_label = start_of_week.strftime('S%Y-%U')

        df_tickets_semaine = df_tickets[
            (df_tickets['Date'] >= start_of_week) & (df_tickets['Date'] <= end_of_week)
        ]
        df_support_semaine = df_support[
            (df_support['Date'] >= start_of_week) & (df_support['Date'] <= end_of_week)
        ]

        if df_tickets_semaine.empty or df_support_semaine.empty:
            continue

        df_week = df_compute_ticket_appels_metrics(agents_n1, df_tickets_semaine, df_support_semaine)
        df_week['score ticket'] = df_week.apply(score_ticket, axis=1)
        df_week['score appel'] = df_week.apply(score_appel, axis=1)

        # Gestion de la division par zéro
        df_week['score total'] = df_week.apply(
            lambda row: (row['score ticket'] + row['score appel']) / 2 
            if (row['score ticket'] + row['score appel']) != 0 else np.nan,
            axis=1
        )

        df_week['Semaine'] = week_label
        historique.append(df_week)

    if not historique:
        raise ValueError("Aucune donnée disponible pour les semaines demandées.")

    df_historique = pd.concat(historique, ignore_index=True)
    df_historique.sort_values(by="Semaine", inplace=True)

    fig = px.line(
        df_historique, 
        x="Semaine", 
        y="score total", 
        color="Agent", 
        title=f"Évolution du score total par agent (dernières {nb_semaines} semaines depuis {max_date.date()})",
        markers=True
    )

    fig.update_layout(
        yaxis=dict(
            range=[0, 1.5]
        )
    )

    fig.add_shape(
        type="line",
        x0=df_historique["Semaine"].min(),
        x1=df_historique["Semaine"].max(),
        y0=0.60,
        y1=0.60,
        line=dict(
            color="green",
            width=2,
            dash="dash"
        ),
    )

    return fig




def df_compute_ticket_appels_metrics(agents_n1, df_ticket, df_support):
    rows = []

    for agent in agents_n1:
        kpis = generate_kpis(df_support, df_ticket, agent, None, 'b2c')

        nb_appels = kpis.get('nb_appels_jour', 0)
        nb_tickets = kpis.get('moy_ticket_agent', 0)
        total = nb_appels + nb_tickets

        # Calculs protégés contre la division par zéro
        pct_tickets = round((nb_tickets / total) * 100, 2) if total != 0 else 0
        pct_appels = round((nb_appels / total) * 100, 2) if total != 0 else 0

        rows.append({
            'Agent': agent,
            "Nombre d'appel traité": round(nb_appels, 2),
            'Nombre de ticket traité': round(nb_tickets, 2),
            "% tickets": pct_tickets,
            "% appels": pct_appels,
            '% appel entrant agent': round(kpis.get('ratio_entrants', 0), 2) * 100,
            'ref_ticket_agent': round(kpis.get('ref_ticket_agent', 0), 2),
            'ref_appel_agent': round(kpis.get('ref_appel_agent', 0), 2),
        })

    df_kpi_agents = pd.DataFrame(rows)
    return df_kpi_agents




def activite_ticket_source_pipeline(df_tickets): 

    df = df_tickets[~df_tickets['Source'].isin(['Téléphone'])].groupby(['Pipeline'])['Ticket ID'].nunique().reset_index()


    fig = px.pie(
        df, 
        names="Pipeline", 
        values="Ticket ID", 
        title="Activité des tickets par pipeline"
    )


    return fig



def activite_ticket_source_client(df): 

    df = df[df['Source'].isin(['Chat', 'E-mail', 'Formulaire'])].groupby(['Semaine', 'Pipeline'])['Ticket ID'].nunique().reset_index()

    fig = px.bar(
        df,
        x="Semaine",
        y="Ticket ID",
        color="Pipeline",
        title="Activité des tickets par pipeline / semaine",
        labels={"working_hours": "Heures de travail", "Semaine": "Semaine"},
    )
    ordre_semaines = sorted(df['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    
    fig.update_layout(
        xaxis=dict(categoryorder="array", categoryarray=ordre_semaines),  # Définit l'ordre des catégories
        #yaxis_title="Appels",
        #yaxis=dict(range=[0, 200])  # Ajuste la plage de l'axe y entre 0 et 100
    )
    return fig

def metrics_nombre_ticket_categorie(df, partenaire=None):
    if partenaire:
        df = df[(df['Pipeline'] == 'SPSA')
                & (df['Formulaire SPSA'] == 'C2 - Assistance niveau 2')
                & (df['Associated Company'] == partenaire)
                & (~df['Source'].isin(['Migration JIRA']))
                & (~df['Semaine'].isin(['S2024-01']))]
    # Si partenaire est None, il faut spécifier que df n'est pas modifié
    # sinon, la ligne est inutile
    # else : 
    #     df  # Cette ligne est redondante et peut être supprimée

    # Filtrer la catégorie "NC" et grouper par catégorie
    df = df[~df['Catégorie'].isin(['NC'])] \
        .groupby('Catégorie')['Nombre_ticket_client'].sum() \
        .reset_index() \
        .sort_values(by='Nombre_ticket_client', ascending=True) \
        .tail(10)

    # Création du graphique
    fig = px.bar(
        df,
        x="Nombre_ticket_client",
        y="Catégorie",
        labels={"Nombre_ticket_client": "Nombre de tickets", "Catégorie": "Catégorie"},
    )
    return fig


def sla(df, partenaire, canal="partenaire"):
    # Dataframe partenaire
    df_partenaire = df[
        (df['Pipeline'] == 'SPSA') &
        (df['Formulaire SPSA'] == 'C2 - Assistance niveau 2') &
        (df['Associated Company'] == partenaire) &
        (~df['Source'].isin(['Migration JIRA'])) &
        (~df['Semaine'].isin(['S2024-01']))
    ]

    # Dataframe b2c
    df_b2c = df[
        (df['Pipeline'].isin(['SSI', 'SSIA','SPSA'])) &
        (df['Source'].isin(['Chat', 'E-mail', 'Formulaire'])) &
        (~df['Source'].isin(['Migration JIRA'])) &
        (~df['Semaine'].isin(['S2024-01']))
    ]

    # Sélection du bon dataframe pour les calculs
    if canal == "partenaire":
        df_selected = df_partenaire
    elif canal == "b2c":
        df_selected = df_b2c
    else:
        raise ValueError("Le paramètre 'canal' doit être 'partenaire' ou 'b2c'.")

    # Filtrage des outliers (mean ± std)
    #mean = df_selected['working_hours'].mean()
    #std = df_selected['working_hours'].std()

    #df_filtered = df_selected[
    #    (df_selected['working_hours'] >= mean - std) &
    #    (df_selected['working_hours'] <= mean + std)
    #]

    df_filtered = df_selected

    # Calculs SLA
    sla_inferieur_2 = (df_filtered['working_hours'] < 2).mean() * 100
    delai_moyen = df_filtered['working_hours'].mean()

    # Moyenne par semaine pour le graphique
    df_semaine = df_filtered.groupby('Semaine', as_index=False)['working_hours'].mean()

    fig = px.bar(
        df_semaine, x='Semaine', y='working_hours',
        title=f'Moyenne délais de réponse (1er mail) - {canal.upper()}',
        labels={'working_hours': 'Délai', 'Semaine': 'Semaine'}
    )

    fig.add_shape(
        type="line", x0=df_semaine['Semaine'].min(), x1=df_semaine['Semaine'].max(),
        y0=2, y1=2, line=dict(color="red", width=2, dash="dash")
    )

    fig.add_annotation(
        xref="paper", x=1.01, y=2, text="Seuil: 2h",
        showarrow=False, font=dict(color="red")
    )

    fig.update_layout(xaxis_tickangle=-45)

    return fig, sla_inferieur_2, delai_moyen, df_partenaire, df_b2c



def evo_appels_ticket(df_ticket, df_support):
    # Agrégation des appels entrants par semaine
    entrant = df_support.groupby(['Semaine','Date'], as_index=False).agg({'Number': 'nunique'})
    entrant = entrant.rename(columns={'Number': 'Entrant'})
    entrant = entrant.groupby('Semaine')['Entrant'].sum()


    # Filtrage des tickets selon la source, puis agrégation
    ticket = df_ticket[
    (df_ticket['Source'].isin(['Chat', 'E-mail', 'Formulaire'])) &
    (df_ticket['Pipeline'].isin(['SSI', 'SSIA', 'SPSA']))
    ].groupby('Semaine', as_index=False)['Ticket ID'].nunique()

    ticket.rename(columns={'Ticket ID': 'Nb_Tickets'}, inplace=True)

    ticket_s17 = ticket[ticket['Semaine'] == 'S2025-15']


    # Fusionner les deux pour garantir la cohérence des semaines
    df_merged = pd.merge(entrant, ticket, on='Semaine', how='outer').fillna(0)

    activite_appels = df_merged['Entrant'].sum() / (df_merged['Nb_Tickets'].sum() + df_merged['Entrant'].sum() )
    activite_tickets = df_merged['Nb_Tickets'].sum() / (df_merged['Nb_Tickets'].sum() + df_merged['Entrant'].sum() )

    # Trier les semaines dans l’ordre chronologique
    ordre_semaines = sorted(df_merged['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df_merged['Semaine'] = pd.Categorical(df_merged['Semaine'], categories=ordre_semaines, ordered=True)
    df_merged.sort_values('Semaine', inplace=True)

    # Création du graphique en aires empilées
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df_merged['Semaine'],
        y=df_merged['Entrant'],
        mode='lines',
        name='Appels entrants',
        stackgroup='one',
        line=dict(color='lightblue')
    ))

    fig.add_trace(go.Scatter(
        x=df_merged['Semaine'],
        y=df_merged['Nb_Tickets'],
        mode='lines',
        name='Tickets',
        stackgroup='one',
        line=dict(color='indianred')
    ))

    fig.update_layout(
        title="Évolution hebdomadaire : Appels entrants + Tickets (aire empilée)",
        xaxis_title="Semaine",
        yaxis_title="Volume total",
        xaxis_tickangle=-45,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified"
    )

    return fig, activite_appels, activite_tickets, ticket_s17


def calcul_taux_reponse(df_support):
    df = df_support
    df['LastState'] = df['LastState'] == 'yes'
    entrants = df[df['direction'] == 'inbound']
    sortants = df[df['direction'] == 'outbound']
    grouped_sortants = sortants.groupby(['Date', 'Number']).agg({'LastState': 'any', 'StartTime': 'mean'}).reset_index()
    grouped_entrants = entrants.groupby(['Date', 'Number']).agg({'LastState': 'any', 'StartTime': 'mean'}).reset_index()
    grouped_sortants = grouped_sortants.rename(columns={'LastState': 'Repondu_sortant', 'StartTime': 'Heure_sortant'})
    merged = pd.merge(grouped_entrants, grouped_sortants, on=['Date', 'Number'], how='left')
    merged['Repondu_sortant'].fillna(False, inplace=True)
    merged['Repondu_total'] = merged['LastState'] | merged['Repondu_sortant']
    taux_reponse = merged['Repondu_total'].mean()
    merged['StartTime'] = pd.to_datetime(merged['StartTime'])
    merged['Heure_sortant'] = pd.to_datetime(merged['Heure_sortant'])
    merged['minute_difference'] = np.where(merged['Repondu_sortant'] == True, (merged['Heure_sortant'] - merged['StartTime']) / pd.Timedelta(minutes=1), np.nan)
    mean_difference = merged['minute_difference'][merged['minute_difference'] >= 0].mean()
    return taux_reponse, mean_difference, merged

def metrics_support(df_support, agents):
    # Gestion des NaN pour 'Taux_de_service' avant conversion
    mean_taux_de_service = df_support['Taux_de_service'].mean()
    Taux_de_service = int(mean_taux_de_service * 100) if not pd.isna(mean_taux_de_service) else 0

    # Gestion des NaN et calcul des appels entrants
    mean_entrant = df_support.groupby('Date').agg({'Entrant': 'sum'}).mean().values[0]
    Entrant = int(mean_entrant) if not pd.isna(mean_entrant) else 0

    # Gestion des NaN et calcul des numéros uniques
    mean_numero_unique = df_support[df_support['direction'] == 'inbound'].groupby('Date').agg({'Number': 'nunique'}).mean().values[0]
    Numero_unique = int(mean_numero_unique) if not pd.isna(mean_numero_unique) else 0

    # Calcul de la moyenne de la durée des appels, en excluant les valeurs <= 0
    temps_moy_appel = df_support[df_support['InCallDuration'] > 0].InCallDuration.mean()
    temps_moy_appel = temps_moy_appel if not pd.isna(temps_moy_appel) else 0

    # Filtrage des agents
    if agents == "agents_support":
        df_support = df_support[df_support['UserName'].isin(["Mourad HUMBLOT", 'Olivier Sainte-Rose', 'Archimede KESSI'])]
    elif agents == "agents_armatis":
        df_support = df_support[df_support['UserName'].isin(["Emilie GEST", 'Morgane Vandenbussche', 'Melinda Marmin', 'Sandrine Sauvage'])]
    elif agents == "agents_all":
        pass  # Pas besoin de filtrer
    else:
        df_support = df_support[df_support['UserName'].isin(agents)]

    # Calcul du nombre d'appels par agent
    Nombre_appel_jour_agent = df_support['UserName'].nunique()

    return [Taux_de_service, Entrant, Numero_unique, temps_moy_appel, Nombre_appel_jour_agent]



def sla_2h_spsa(df_tickets):
    df_sla_2h = df_tickets[(df_tickets['working_hours'] < 2) & (df_tickets['working_hours'] > 0.1)].groupby(['Semaine'])['working_hours'].count().reset_index().rename(columns={'working_hours': 'SLA < 2h'})
    df_tickets_std = df_tickets.groupby(['Semaine'])['working_hours'].count().reset_index().rename(columns={'working_hours': 'Nb tickets'})
    df_tickets_sla = pd.merge(df_tickets_std, df_sla_2h, how='left', on='Semaine')
    pourcentage_sla = (df_tickets_sla['SLA < 2h'].mean() / df_tickets_sla['Nb tickets'].mean()) * 100
    pourcentage_sla = f"{pourcentage_sla:.2f}%"
    return pourcentage_sla, df_tickets_std

def activite_ticket_source_client(df_tickets):
    df = df_tickets[df_tickets['Source'].isin(['Chat', 'E-mail', 'Formulaire'])].groupby(['Semaine', 'Pipeline'])['Ticket ID'].nunique().reset_index()
    fig = px.bar(df, x="Semaine", y="Ticket ID", color="Pipeline", title="Activité des tickets par pipeline / semaine", labels={"working_hours": "Heures de travail", "Semaine": "Semaine"})
    return fig


def graph_charge_affid_stellair(df_support):


    df_support = df_support[df_support['Semaine'] != 'S2024-01']

    essai = df_support[df_support['Semaine'] == 'S2025-6']['IVR Branch'].value_counts()

    affid = df_support[(df_support['LastState'] == True) & (df_support['Logiciel'] == 'Affid')] \
        .groupby(['Semaine']).agg({'Count': 'sum'}).rename(columns={'Count': 'Affid'}).reset_index()
    stellair = df_support[(df_support['LastState'] == True) & (df_support['Logiciel'] == 'Stellair')] \
        .groupby(['Semaine']).agg({'Count': 'sum'}).rename(columns={'Count': 'Stellair'}).reset_index()

    df_resultats = pd.merge(stellair, affid, on='Semaine').reset_index()

    df_resultats['Charge_support_call_stellair'] = df_resultats['Stellair'] / (df_resultats['Stellair'] + df_resultats['Affid'])
    df_resultats['Charge_support_call_affid'] = df_resultats['Affid'] / (df_resultats['Stellair'] + df_resultats['Affid'])

    fig = px.bar(df_resultats, 
                 x='Semaine',
                 y=['Charge_support_call_stellair', 'Charge_support_call_affid'],
                 template='plotly_dark')
    
    fig.update_yaxes(tickformat=".0%")
    fig.update_layout(
        title_text="Activité en % - NXT & Stellair",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    fig2 = px.bar(df_resultats, 
                  x='Semaine',
                  y=['Stellair', 'Affid'],
                  template='plotly_dark')
    
    fig2.update_layout(
        title_text="Activité en Nb - NXT & Stellair",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    return fig, fig2, essai



def generate_kpis(df_support, df_tickets, agents, partenaire=None, partenaires=None):
    """
    Génère les KPIs pour un ou plusieurs agents.
    Retourne un dictionnaire avec les métriques, y compris les graphes, SLA, etc.
    """
    # Sauvegarde de l'entrée pour condition ultérieure
    agent_original = agents

    # --- DÉFINITIONS DES GROUPES D’AGENTS ---
    support_agents = ['Archimede KESSI', 'Mourad HUMBLOT', 'Olivier Sainte-Rose']
    armatis_agents = ['Emilie GEST', 'Sandrine Sauvage', 'Melinda Marmin', 'Morgane Vandenbussche']
    special_agents = ['agents_support', 'agents_armatis', 'agents_all']


    # --- SÉLECTION DES AGENTS CIBLÉS ---
    if agents == 'agents_support':
        selected_agents = support_agents
    elif agents == 'agents_armatis':
        selected_agents = armatis_agents
    elif agents == 'agents_all':
        selected_agents = support_agents + armatis_agents
    else:
        selected_agents = [agents] if isinstance(agents, str) else agents

    # --- KPIs SUPPORT ---
    taux_reponse, mean_difference, df_taux_reponse = calcul_taux_reponse(df_support)
    com_jour, temps_moy_com, nb_appels_jour, nb_appel, nb_appels_jour_entrants, nb_appels_jour_sortants, ratio_entrants, ratio_sortants = kpi_agent(selected_agents, df_support)
    Taux_de_service, Entrant, Numero_unique, temps_moy_appel, Nombre_appel_jour_agent = metrics_support(df_support, selected_agents)
    charge_affid_stellair_1, charge_affid_stellair_2, essai = graph_charge_affid_stellair(df_support)


    # --- KPIs TICKETS ---
    (
        total_tickets_moyenne,
        moyenne_appels_entrant_total,
        moyenne_tickets_total,
        moyenne_appels_traite_global,
        moyenne_tickets_traite_global,
        moyenne_tickets_traite_agent,
        moyenne_appels_traite_agent, 
        moyenne_appels_entrant_agent,
        df_ticket_effectif,
        total_tickets_traite,
        df_ticket_ref,
        df_appel_ref
    ) = compute_ticket_appels_metrics(df_tickets, df_support, selected_agents)

    activite_ticket = activite_ticket_source_pipeline(df_tickets)
    activite_ticket_semaine = activite_ticket_source_client(df_tickets)
    activite_categorie = metrics_nombre_ticket_categorie(df_tickets, partenaire)
    evo_appels_tickets, activite_appels_pourcentage,  activite_tickets_pourcentage, ticket_s17= evo_appels_ticket(df_tickets, df_support)

    # --- SLA / RAPPORT PARTENAIRE ---
    sla_fig, sla_inferieur_2, delai_moyen, df_partenaire, df_b2c = sla(df_tickets, None, 'b2c')
    sla_partenaire, describe_partenaire = sla_2h_spsa(df_tickets)

    # --- GRAPHIQUES ---
    fig_activite_ticket = activite_ticket_source_client(df_tickets)
    fig_activite = graph_activite(df_support)
    fig_taux_jour = graph_taux_jour(df_support)
    fig_taux_heure = graph_taux_heure(df_support)

    # --- RÉPARTITION TICKETS / APPELS ---
    total_interactions = moyenne_tickets_total + nb_appels_jour_entrants
    pourcentage_tickets = moyenne_tickets_total / total_interactions if total_interactions > 0 else 0
    pourcentage_appels = nb_appels_jour_entrants / total_interactions if total_interactions > 0 else 0

    pourcentage_tickets_agent = moyenne_tickets_traite_agent / (moyenne_tickets_traite_agent + moyenne_appels_traite_agent)
    pourcentage_appels_agent = moyenne_appels_traite_agent / (moyenne_tickets_traite_agent + moyenne_appels_traite_agent)

    # --- STRUCTURE DU DICTIONNAIRE DE SORTIE ---
    kpis = {

        # SLA / partenaires
        'sla_partenaire': sla_partenaire,
        'describe_partenaire': describe_partenaire,
        'sla_fig': sla_fig,
        '%_sla_2': sla_inferieur_2,
        'delai_moyen_reponse': delai_moyen,

        # Activité générale
        'taux_reponse': taux_reponse,
        'mean_difference': mean_difference,
        'Taux_de_service': Taux_de_service,
        'Entrant': Entrant,
        'Numero_unique': Numero_unique,
        'temps_moy_appel': temps_moy_appel,
        'Nombre_appel_jour_agent': nb_appels_jour,
        'Nomnbre_appel_total': nb_appel,

        # Tickets / Appels pour objectif
        'moy_appel_jour': moyenne_appels_entrant_total,
        'moy_ticket_jour': moyenne_tickets_total,
        'moy_appel_traite': moyenne_appels_traite_global,
        'moy_ticket_traite': moyenne_tickets_traite_global,
        'moy_ticket_agent': moyenne_tickets_traite_agent,
        'ref_ticket_agent':df_ticket_ref,
        'ref_appel_agent': df_appel_ref,
        'moy_appel_agent': moyenne_appels_traite_agent,
        'moy_entrant_agent': moyenne_appels_entrant_agent,
        'Total ticket traite': total_tickets_traite,
        '% tickets': pourcentage_tickets,
        '% Appels': pourcentage_appels,
        '% tickets agent': pourcentage_tickets_agent,
        '% Appels agent': pourcentage_appels_agent,
        'Effectif moyen': df_ticket_effectif,
        'activite_appels_pourcentage': activite_appels_pourcentage,
        'activite_tickets_pourcentage': activite_tickets_pourcentage,
        

        # Évolution
        'evo_appels_tickets': evo_appels_tickets,
        'essai': ticket_s17,

        # Charge Stellair
        'charge_affid_stellair_%': charge_affid_stellair_1,
        'charge_affid_stellair_v': charge_affid_stellair_2,
        'essai': essai,

        # Graphiques
        'fig_activite_ticket': fig_activite_ticket,
        'activite_categorie':activite_categorie,
        'activite_ticket_semaine': activite_ticket_semaine,
        'fig_activite': fig_activite,
        'fig_taux_jour': fig_taux_jour,
        'fig_taux_heure': fig_taux_heure,
    }

    # --- AJOUT DES KPIs AGENT UNIQUEMENT SI PAS UN GROUPE ---
    if agent_original not in special_agents:
        kpis.update({
            'com_jour': com_jour,
            'temps_moy_com': temps_moy_com,
            'nb_appels_jour': nb_appels_jour,
            'nb_appels_jour_entrants': nb_appels_jour_entrants,
            'nb_appels_jour_sortants': nb_appels_jour_sortants,
            'ratio_entrants': ratio_entrants,
            'ratio_sortants': ratio_sortants,
        })

    return kpis
