import pandas as pd
import numpy as np
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px



def filter_evaluation(df_evaluation, agents, periodes_eval):
    return df_evaluation[
        (df_evaluation["agent"].isin(agents)) &
        (df_evaluation["quarter"].isin(periodes_eval))]


def charge_entrant_sortant(df_support, agents):
    import pandas as pd

    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_support = df_support[df_support['Semaine'] != 'S2024-01'].copy()

    if isinstance(agents, str):
        agents = [agents]

    # Filtrage par les agents et par l'état 'yes'
    df_filtered = df_support[
        (df_support['UserName'].isin(agents)) &
        (df_support['direction'] == 'inbound')  # On se concentre sur les appels entrants
    ]

    # Vérifier si le dataframe filtré est vide
    if df_filtered.empty:
        print("Aucune donnée trouvée pour les agents sélectionnés")
        return None, None

    # Groupement par Semaine, Agent, et Direction
    df_grouped = df_filtered.groupby(['Semaine', 'UserName']).agg({'Date': 'count'}).reset_index()

    # Vérifier si le groupement est vide
    if df_grouped.empty:
        print("Aucune donnée après groupement")
        return None, None

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
        title="Volume d'appels entrants par agent"
    )

    fig_line.update_layout(
        yaxis_title="Nombre d'appels",
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
        #title="Répartition des appels entrants par agent"
    )

    return fig_line, fig_pie



def charge_ticket(df_ticket, agents):
    import pandas as pd

    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_ticket = df_ticket[df_ticket['Semaine'] != 'S2024-01'].copy()

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
        title="Volume de tickets par agent"
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
        #title="Répartition des tickets par agent"
    )

    return fig_line, fig_pie




def metrics_nombre_ticket_pipeline_agent(df_hubspot, agents):
    df_hubspot['Date'] = pd.to_datetime(df_hubspot['Date'])
    
    # Définitions des groupes d'agents
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
        "Céline": ['Céline Crendal', 'Celine Crendal', 'Céline CRENDAL', 'Celine CRENDAL']
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
    df_support['InCallDuration_format'] = pd.to_timedelta(df_support['InCallDuration_format'], errors='coerce')

    # Calcul de la moyenne des durées d'appel par jour
    com_jour = df_support['InCallDuration_format'].mean()

    def timedelta_to_hms(td):
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"


    # Vérifier si la valeur de com_jour est valide
    if pd.notnull(com_jour):  # Vérifie si com_jour n'est pas NaT
        com_jour = timedelta_to_hms(com_jour)

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
    
    # Calcul des KPI de productivité (4 valeurs seulement)
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


def calculer_scores_equipe(df_support, df_tickets, agents_n1):
    """
    Calcule les scores de l'équipe en utilisant la même méthode que le dashboard.
    """
    rows = []
    
    for agent in agents_n1:
        # Calcul des KPIs pour chaque agent
        kpis = generate_kpis(df_support, df_tickets, agent, None, 'b2c')
        
        nb_appels = kpis.get('nb_appels_jour', 0)
        nb_tickets = kpis.get('moy_ticket_agent', 0)
        volume_total = nb_appels + nb_tickets
        
        # Créer une ligne avec le format attendu par calculate_performance_score
        row_data = {
            'Agent': agent,
            "Nombre d'appel traité": nb_appels,
            'Nombre de ticket traité': nb_tickets,
            '% appel entrant agent': kpis.get('ratio_entrants', 0) * 100
        }
        
        # Utiliser calculate_performance_score pour le calcul du score
        score_total = calculate_performance_score(
            row_data,
            objectif_total=25,
            ratio_appels=0.7,
            ratio_tickets=0.3,
            objectif_taux_service=0.70
        )
        
        # Calculer les pourcentages
        if volume_total > 0:
            pct_appels = (nb_appels / volume_total) * 100
            pct_tickets = (nb_tickets / volume_total) * 100
        else:
            pct_appels = 0
            pct_tickets = 0
        
        # Ajouter les résultats
        rows.append({
            'Agent': agent,
            'Score Total': score_total,
            'Volume Total': volume_total,
            '% Appels': pct_appels,
            '% Tickets': pct_tickets
        })
    
    return pd.DataFrame(rows)

def historique_scores_total(agents_n1, df_tickets, df_support, date_debut=None, nb_semaines=12):
    historique = []

    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()
    df_support = df_support[df_support['Semaine'] != 'S2024-01'].copy()

    if isinstance(agents_n1, str):
        agents_n1 = [agents_n1]

    # Identifier toutes les semaines uniques et les trier correctement
    all_weeks = df_tickets['Semaine'].unique()
    
    # Gérer les cas où il n'y a pas de semaines
    if len(all_weeks) == 0:
        fig = px.line()
        fig.update_layout(title="Aucune donnée de semaine disponible", yaxis=dict(range=[0, 100]))
        return fig
        
    sorted_weeks = sorted(all_weeks, key=lambda x: (int(x[1:5]), int(x[6:])))
    
    # Prendre les `nb_semaines` dernières semaines
    weeks_to_process = sorted_weeks[-nb_semaines:]
    
    for week_label in weeks_to_process:
        # Filtrer les dataframes pour la semaine en cours
        df_tickets_semaine = df_tickets[df_tickets['Semaine'] == week_label]
        df_support_semaine = df_support[df_support['Semaine'] == week_label]

        if df_tickets_semaine.empty and df_support_semaine.empty:
            continue

        df_week = df_compute_ticket_appels_metrics(agents_n1, df_tickets_semaine, df_support_semaine)
        
        if df_week.empty:
            continue

        df_week['Semaine'] = week_label
        historique.append(df_week)

    if not historique:
        print("Aucune donnée disponible pour les semaines demandées.")
        fig = px.line()
        fig.update_layout(
            title="Aucune donnée disponible pour l'historique des scores",
            yaxis=dict(range=[0, 100])
        )
        return fig

    df_historique = pd.concat(historique, ignore_index=True)
    
    # Assurer le bon ordre pour le graphique
    df_historique['Semaine'] = pd.Categorical(df_historique['Semaine'], categories=weeks_to_process, ordered=True)
    df_historique = df_historique.sort_values(by="Semaine")

    # Calculer le score de performance pour chaque ligne
    df_historique['score_performance'] = df_historique.apply(
        lambda row: calculate_performance_score(
            row,
            objectif_total=25,
            ratio_appels=0.7,
            ratio_tickets=0.3,
            objectif_taux_service=0.70
        ),
        axis=1
    )

    fig = px.line(
        df_historique, 
        x="Semaine", 
        y="score_performance", 
        color="Agent", 
        #title=f"Évolution du score total par agent (dernières {nb_semaines} semaines depuis {max_date.date()})",
        markers=True
    )

    # Mise à jour des paramètres du graphique
    fig.update_layout(
        yaxis=dict(
            range=[0, 100],  # Échelle de 0 à 100
            tickformat=".0f"  # Format sans décimales
        )
    )

    # Ajout des lignes de seuil
    fig.add_shape(  # Seuil vert (70%)
        type="line",
        x0=df_historique["Semaine"].min(),
        x1=df_historique["Semaine"].max(),
        y0=70,
        y1=70,
        line=dict(
            color="green",
            width=2,
            dash="dash"
        )
    )

    fig.add_shape(  # Seuil orange (70%)
        type="line",
        x0=df_historique["Semaine"].min(),
        x1=df_historique["Semaine"].max(),
        y0=70,
        y1=70,
        line=dict(
            color="orange",
            width=2,
            dash="dash"
        )
    )

    return fig

def graph_activite_xmed(df_support):
    from plotly.subplots import make_subplots
    import plotly.graph_objects as go

    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_support = df_support[df_support['Semaine'] != 'S2024-01'].copy()

    # Filtrage des données pour l'activité entrante et la ligne XMED
    data_graph1 = df_support[(df_support['direction'] == 'inbound') & (df_support['line'] == 'xmed')]

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

    # Tri chronologique des semaines
    ordre_semaines = sorted(data_graph3['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    data_graph3['Semaine'] = pd.Categorical(data_graph3['Semaine'], categories=ordre_semaines, ordered=True)
    data_graph3 = data_graph3.sort_values('Semaine')

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
        title="Graphique avec Taux en barres et Numero_unique/Entrant en aires empilées",
        template="plotly_dark",
        xaxis_title="Semaine",
        yaxis_title="Valeurs",
        title_text="Activité & Taux de service XMED - 20 semaines",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )

    return fig

def calculate_ticket_response_time(df_tickets, agents=None):
    """
    Calcule les temps de réponse aux tickets avec filtrage des valeurs aberrantes.
    
    Parameters:
    -----------
    df_tickets : pandas DataFrame
        DataFrame contenant les données de tickets
    agents : list or None
        Liste des agents à filtrer. Si None, tous les agents sont inclus.
    
    Returns:
    --------
    tuple : (moyenne_temps_reponse, graphique_evolution)
    """
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    
    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()
    
    # Utiliser EXACTEMENT les mêmes filtres que evo_appels_ticket
    df_filtered = df_tickets[
        (df_tickets['Source'].isin(['Chat', 'E-mail', 'Formulaire'])) &
        (df_tickets['Pipeline'].isin(['SSI', 'SSIA', 'SPSA']))
    ]

    
    # Filtrer par agents si spécifié
    if agents is not None:
        df_filtered = df_filtered[df_filtered['Propriétaire du ticket'].isin(agents)]
    
    # Filtrer les valeurs aberrantes (±5 écarts-types) pour le calcul de la moyenne uniquement
    working_hours = df_filtered['working_hours'].dropna()
    
    if len(working_hours) > 0:
        mean_hours = working_hours.mean()
        std_hours = working_hours.std()
        
        # Filtrage des valeurs aberrantes pour le calcul de la moyenne uniquement
        lower_bound = mean_hours - 5 * std_hours
        upper_bound = mean_hours + 5 * std_hours
        
        df_clean_for_mean = df_filtered[
            (df_filtered['working_hours'] >= lower_bound) &
            (df_filtered['working_hours'] <= upper_bound) &
            (df_filtered['working_hours'] > 0)  # Exclure les valeurs négatives ou nulles
        ]
        
        # Calcul de la moyenne des temps de réponse nettoyés
        moyenne_temps_reponse = df_clean_for_mean['working_hours'].mean()
        
        # Pour le graphique d'évolution, utiliser TOUS les tickets filtrés (même logique que evo_appels_ticket)
        # sans aucun filtrage supplémentaire pour avoir exactement le même comptage
        df_evolution = df_filtered.groupby('Semaine')['working_hours'].agg(['mean', 'count']).reset_index()
        
        # Trier les semaines chronologiquement (même logique que evo_appels_ticket)
        ordre_semaines = sorted(df_evolution['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
        df_evolution['Semaine'] = pd.Categorical(df_evolution['Semaine'], categories=ordre_semaines, ordered=True)
        df_evolution = df_evolution.sort_values('Semaine')
        
        # Créer le graphique
        fig = make_subplots(
            rows=2, cols=1,
            subplot_titles=('Évolution du temps de réponse moyen par semaine', 'Nombre de tickets par semaine'),
            vertical_spacing=0.15,
            row_heights=[0.7, 0.3]
        )
        
        # Graphique du temps de réponse moyen
        fig.add_trace(
            go.Scatter(
                x=df_evolution['Semaine'],
                y=df_evolution['mean'],
                mode='lines+markers',
                name='Temps de réponse moyen (h)',
                line=dict(color='blue', width=2),
                marker=dict(size=8)
            ),
            row=1, col=1
        )
        
        # Ligne de référence à 2h
        fig.add_hline(
            y=2, 
            line_dash="dash", 
            line_color="red",
            annotation_text="Seuil 2h",
            row=1, col=1
        )
        
        # Graphique du nombre de tickets (même comptage que evo_appels_ticket)
        fig.add_trace(
            go.Bar(
                x=df_evolution['Semaine'],
                y=df_evolution['count'],
                name='Nombre de tickets',
                marker_color='lightblue'
            ),
            row=2, col=1
        )

        fig.update_layout(
            title="Évolution des temps de réponse aux tickets",
        template="plotly_dark",
            height=700,
            showlegend=True,
            xaxis2_title="Semaine",
            yaxis_title="Temps de réponse (heures)",
            yaxis2_title="Nombre de tickets"
        )
        
        # Rotation des étiquettes de l'axe x et ajustement des marges
        fig.update_xaxes(tickangle=-45)
        
        # Ajuster les marges pour éviter les chevauchements
        fig.update_layout(
            margin=dict(l=50, r=50, t=100, b=100)
        )
        
    else:
        moyenne_temps_reponse = 0
        fig = go.Figure()
        fig.add_annotation(
            text="Aucune donnée disponible pour les temps de réponse",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        fig.update_layout(
            title="Évolution des temps de réponse aux tickets",
            template="plotly_dark"
        )
    
    return moyenne_temps_reponse, fig

def graph_tickets_n2_par_semaine(df_tickets):
    """
    Crée un graphique à barres cumulées du nombre de tickets passés par le N2 par semaine, 
    avec une courbe du nombre de tickets encore ouverts.
    """
    import plotly.express as px
    import plotly.graph_objects as go
    import pandas as pd

    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()

    df_n2 = df_tickets[df_tickets['Passé par le support N2'] == 'Oui'].copy()
    
    if df_n2.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucun ticket passé par le support N2", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(title="Tickets passés par le N2 par semaine et par pipeline", template="plotly_dark")
        return fig

    # --- Données pour les barres (total des tickets passés au N2 par semaine/pipeline) ---
    df_grouped = df_n2.groupby(['Semaine', 'Pipeline'])['Ticket ID'].nunique().reset_index()

    ordre_semaines = sorted(df_grouped['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df_grouped['Semaine'] = pd.Categorical(df_grouped['Semaine'], categories=ordre_semaines, ordered=True)
    df_grouped = df_grouped.sort_values('Semaine')

    # --- Création du graphique à barres ---
    fig = px.bar(
        df_grouped,
        x='Semaine',
        y='Ticket ID',
        color='Pipeline',
        title="Tickets N2 passés par semaine (barres) et ouverts (ligne)",
        labels={'Ticket ID': 'Nombre de tickets'},
        barmode='stack'
    )

    # --- Données pour la courbe (stock de tickets ouverts via Temps de fermeture) ---
    weekly_open_counts = []
    
    # Pour cette logique, nous avons besoin d'une seule ligne par ticket avec sa date de création et son temps de fermeture.
    df_n2['Date'] = pd.to_datetime(df_n2['Date'])
    
    # Obtenir la date de création (la plus ancienne date pour chaque ticket)
    creation_dates = df_n2.groupby('Ticket ID')['Date'].min().reset_index().rename(columns={'Date': 'Date de création'})
    
    # Obtenir le temps de fermeture depuis la dernière entrée de chaque ticket
    latest_entries = df_n2.loc[df_n2.groupby('Ticket ID')['Date'].idxmax()][['Ticket ID', 'Temps de fermeture (HH:mm:ss)']]
    
    # Combiner pour avoir une vue unique et propre par ticket
    df_n2_summary = pd.merge(creation_dates, latest_entries, on='Ticket ID', how='left')

    # Calculer la date de fermeture réelle
    df_n2_summary['Temps de fermeture'] = pd.to_timedelta(df_n2_summary['Temps de fermeture (HH:mm:ss)'], errors='coerce')
    df_n2_summary['Date de fermeture'] = df_n2_summary['Date de création'] + df_n2_summary['Temps de fermeture']

    # --- Données pour la courbe des tickets N2 ouverts de Pierre Goupillon ---
    weekly_pierre_open_counts = []
    
    # Gestion des variantes de nom pour Pierre Goupillon
    correspondance_pierre = ['Pierre Goupillon', 'Pierre GOUPILLON']
    
    # Obtenir les tickets N2 de Pierre Goupillon
    df_n2_pierre = df_n2[df_n2['Propriétaire du ticket'].isin(correspondance_pierre)].copy()
    
    if not df_n2_pierre.empty:
        # Obtenir la date de création pour les tickets de Pierre
        creation_dates_pierre = df_n2_pierre.groupby('Ticket ID')['Date'].min().reset_index().rename(columns={'Date': 'Date de création'})
        
        # Obtenir le temps de fermeture depuis la dernière entrée de chaque ticket de Pierre
        latest_entries_pierre = df_n2_pierre.loc[df_n2_pierre.groupby('Ticket ID')['Date'].idxmax()][['Ticket ID', 'Temps de fermeture (HH:mm:ss)']]
        
        # Combiner pour avoir une vue unique et propre par ticket de Pierre
        df_n2_pierre_summary = pd.merge(creation_dates_pierre, latest_entries_pierre, on='Ticket ID', how='left')
        
        # Calculer la date de fermeture réelle pour les tickets de Pierre
        df_n2_pierre_summary['Temps de fermeture'] = pd.to_timedelta(df_n2_pierre_summary['Temps de fermeture (HH:mm:ss)'], errors='coerce')
        df_n2_pierre_summary['Date de fermeture'] = df_n2_pierre_summary['Date de création'] + df_n2_pierre_summary['Temps de fermeture']

    for week in ordre_semaines:
        # Déterminer la date de fin de la semaine
        year, week_num = int(week[1:5]), int(week[6:])
        week_end_date = pd.to_datetime(f'{year}-W{week_num}-0', format='%Y-W%W-%w') + pd.Timedelta(days=6)
        
        # 1. Sélectionner les tickets créés avant la fin de cette semaine
        tickets_created_so_far = df_n2_summary[df_n2_summary['Date de création'] <= week_end_date]
        
        # 2. Compter ceux qui sont encore ouverts à cette date
        # Un ticket est ouvert si sa date de fermeture est future OU si elle n'existe pas (NaT)
        open_count = tickets_created_so_far[
            (tickets_created_so_far['Date de fermeture'] > week_end_date) | 
            (pd.isna(tickets_created_so_far['Date de fermeture']))
        ].shape[0]
        
        weekly_open_counts.append({'Semaine': week, 'Tickets ouverts': open_count})
        
        # 3. Calculer les tickets ouverts de Pierre Goupillon pour cette semaine
        if not df_n2_pierre.empty:
            tickets_pierre_created_so_far = df_n2_pierre_summary[df_n2_pierre_summary['Date de création'] <= week_end_date]
            
            pierre_open_count = tickets_pierre_created_so_far[
                (tickets_pierre_created_so_far['Date de fermeture'] > week_end_date) | 
                (pd.isna(tickets_pierre_created_so_far['Date de fermeture']))
            ].shape[0]
        else:
            pierre_open_count = 0
            
        weekly_pierre_open_counts.append({'Semaine': week, 'Tickets ouverts Pierre': pierre_open_count})

    df_open_counts = pd.DataFrame(weekly_open_counts)
    df_pierre_open_counts = pd.DataFrame(weekly_pierre_open_counts)
    
    # --- Ajout de la courbe au graphique ---
    fig.add_trace(go.Scatter(
        x=df_open_counts['Semaine'],
        y=df_open_counts['Tickets ouverts'],
        mode='lines+markers',
        name='Stock de Tickets N2 Ouverts',
        line=dict(color='yellow', width=3)
    ))
    
    # --- Ajout de la courbe des tickets de Pierre Goupillon ---
    fig.add_trace(go.Scatter(
        x=df_pierre_open_counts['Semaine'],
        y=df_pierre_open_counts['Tickets ouverts Pierre'],
        mode='lines+markers',
        name='Stock de Tickets N2 Ouverts - Pierre Goupillon',
        line=dict(color='red', width=3, dash='dash')
    ))

    # --- Mise en page finale ---
    fig.update_xaxes(
        categoryorder='array',
        categoryarray=ordre_semaines,
        tickangle=-45
    )
    fig.update_layout(
        template="plotly_dark",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis_title="Nombre de tickets"
    )

    return fig

def graph_tickets_n2_resolus_par_agent(df_tickets):
    """
    Crée un graphique du nombre de tickets résolus par le N2 par semaine et par agent.
    """
    import plotly.express as px
    import plotly.graph_objects as go
    import pandas as pd
    
    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()
    
    df_n2_resolus = df_tickets[
        (df_tickets['Passé par le support N2'] == 'Oui') &
        (df_tickets['Statut du ticket'] == 'Résolu')
    ].copy()
    
    if df_n2_resolus.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucun ticket résolu par le support N2", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(title="Tickets N2 résolus par semaine et par agent", template="plotly_dark")
        return fig

    df_grouped = df_n2_resolus.groupby(['Semaine', 'Propriétaire du ticket'])['Ticket ID'].nunique().reset_index()
    
    # Tri chronologique des semaines
    ordre_semaines = sorted(df_grouped['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df_grouped['Semaine'] = pd.Categorical(df_grouped['Semaine'], categories=ordre_semaines, ordered=True)
    df_grouped = df_grouped.sort_values('Semaine')
    
    fig = px.bar(
        df_grouped,
        x='Semaine',
        y='Ticket ID',
        color='Propriétaire du ticket',
        title="Tickets N2 résolus par semaine et par agent",
        labels={'Ticket ID': 'Nombre de tickets'}
    )
    
    # Forcer l'ordre chronologique sur l'axe X
    fig.update_xaxes(
        categoryorder='array',
        categoryarray=ordre_semaines,
        tickangle=-45
    )
    
    fig.update_layout(
        template="plotly_dark",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    return fig


def graph_tickets_pierre_goupillon(df_tickets):
    """
    Crée un graphique du nombre de tickets de Pierre Goupillon par semaine.
    """
    import plotly.express as px
    import plotly.graph_objects as go
    import pandas as pd
    
    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()
    
    # Gestion des variantes de nom pour Pierre Goupillon
    correspondance_pierre = ['Pierre Goupillon', 'Pierre GOUPILLON']
    
    # Filtrer avec toutes les variantes possibles
    df_pierre = df_tickets[df_tickets['Propriétaire du ticket'].isin(correspondance_pierre)].copy()
    
    if df_pierre.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucun ticket pour Pierre Goupillon", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(title="Tickets de Pierre Goupillon par semaine", template="plotly_dark")
    return fig

    df_grouped = df_pierre.groupby('Semaine')['Ticket ID'].nunique().reset_index()

    # Tri chronologique des semaines
    ordre_semaines = sorted(df_grouped['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df_grouped['Semaine'] = pd.Categorical(df_grouped['Semaine'], categories=ordre_semaines, ordered=True)
    df_grouped = df_grouped.sort_values('Semaine')
    
    fig = px.bar(
        df_grouped,
        x='Semaine',
        y='Ticket ID',
        title="Tickets de Pierre Goupillon par semaine",
        labels={'Ticket ID': 'Nombre de tickets'}
    )

    fig.update_xaxes(
        categoryorder='array',
        categoryarray=ordre_semaines,
        tickangle=-45
    )

    fig.update_layout(
        template="plotly_dark"
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

    # Tri chronologique des semaines
    ordre_semaines = sorted(df_semaine['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df_semaine['Semaine'] = pd.Categorical(df_semaine['Semaine'], categories=ordre_semaines, ordered=True)
    df_semaine = df_semaine.sort_values('Semaine')

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

def metrics_nombre_ticket_categorie(df, partenaire=None):
    if partenaire:
        df = df[(df['Pipeline'] == 'SPSA')
                & (df['Formulaire SPSA'] == 'C2 - Assistance niveau 2')
                & (df['Associated Company'] == partenaire)
                & (~df['Source'].isin(['Migration JIRA']))
                & (~df['Semaine'].isin(['S2024-01']))]

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

def generate_kpis(df_support, df_tickets, agents, partenaire=None, partenaires=None):
    """
    Génère les KPIs pour un ou plusieurs agents.
    Retourne un dictionnaire avec les métriques, y compris les graphes, SLA, etc.
    """
    # Sauvegarde de l'entrée pour condition ultérieure
    agent_original = agents

    # --- DÉFINITIONS DES GROUPES D'AGENTS ---
    support_agents = ['Archimede KESSI', 'Mourad HUMBLOT', 'Olivier Sainte-Rose']
    armatis_agents = ['Emilie GEST', 'Sandrine Sauvage', 'Melinda Marmin', 'Morgane Vandenbussche', 'Céline Crendal']
    special_agents = ['agents_support', 'agents_armatis', 'agents_all']

    # Gestion des variantes de nom pour Céline Crendal
    correspondance_noms = {
        'Céline Crendal': ['Céline Crendal', 'Celine Crendal', 'Céline CRENDAL', 'Celine CRENDAL'],
        'Celine Crendal': ['Céline Crendal', 'Celine Crendal', 'Céline CRENDAL', 'Celine CRENDAL']
    }

    # --- SÉLECTION DES AGENTS CIBLÉS ---
    if agents == 'agents_support':
        selected_agents = support_agents
    elif agents == 'agents_armatis':
        selected_agents = armatis_agents
    elif agents == 'agents_all':
        selected_agents = support_agents + armatis_agents
    else:
        # Si agents est une liste, on l'utilise directement
        if isinstance(agents, list):
            # Si Céline est dans la liste, on ajoute toutes ses variantes
            selected_agents = []
            for a in agents:
                if a in correspondance_noms:
                    selected_agents.extend(correspondance_noms[a])
                else:
                    selected_agents.append(a)
        else:
            if agents in correspondance_noms:
                selected_agents = correspondance_noms[agents]
            else:
                selected_agents = [agents]

    # --- KPIs SUPPORT ---
    taux_reponse, mean_difference, df_taux_reponse = calcul_taux_reponse(df_support)
    com_jour, temps_moy_com, nb_appels_jour, nb_appel, nb_appels_jour_entrants, nb_appels_jour_sortants, ratio_entrants, ratio_sortants = kpi_agent(selected_agents, df_support)
    Taux_de_service, Entrant, Numero_unique, temps_moy_appel, Nombre_appel_jour_agent = metrics_support(df_support, selected_agents)
    charge_affid_stellair_1, charge_affid_stellair_2, essai = graph_charge_affid_stellair(df_support)

    # --- KPIs TICKETS ---
    # Filtrer les tickets pour l'agent
    agent_tickets = df_tickets[
        (df_tickets['Propriétaire du ticket'].isin(selected_agents)) & 
        (df_tickets['Pipeline'].isin(['SSIA', 'SSI', 'SPSA'])) &
        (df_tickets['Source'].isin(['Chat', 'E-mail', 'Formulaire']) | pd.isna(df_tickets['Source']))
    ]
    
    # Calculer la moyenne de tickets par jour
    nb_jours = len(df_support['Date'].unique())
    if nb_jours == 0:
        nb_jours = 1  # Pour éviter la division par zéro
    moy_ticket_agent = len(agent_tickets) / nb_jours

    activite_ticket = activite_ticket_source_pipeline(df_tickets)
    activite_ticket_semaine = activite_ticket_source_client(df_tickets)
    activite_categorie = metrics_nombre_ticket_categorie(df_tickets, partenaire)
    evo_appels_tickets, activite_appels_pourcentage, activite_tickets_pourcentage, ticket_s17 = evo_appels_ticket(df_tickets, df_support)

    # --- SLA / RAPPORT PARTENAIRE ---
    canal = "partenaire" if partenaire else "b2c"
    sla_fig, sla_inferieur_2, delai_moyen, df_partenaire, df_b2c = sla(df_tickets, partenaire, canal)
    sla_partenaire, describe_partenaire = sla_2h_spsa(df_tickets)

    # --- GRAPHIQUES ---
    fig_activite_ticket = activite_ticket_source_client(df_tickets)
    fig_activite = graph_activite(df_support)
    fig_taux_jour = graph_taux_jour(df_support)
    fig_taux_heure = graph_taux_heure(df_support)

    # --- RÉPARTITION TICKETS / APPELS ---
    total_interactions = moy_ticket_agent + nb_appels_jour_entrants
    pourcentage_tickets = moy_ticket_agent / total_interactions if total_interactions > 0 else 0
    pourcentage_appels = nb_appels_jour_entrants / total_interactions if total_interactions > 0 else 0

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
        'moy_ticket_agent': moy_ticket_agent,
        'nb_appels_jour': nb_appels_jour,
        'nb_appels_jour_entrants': nb_appels_jour_entrants,
        'nb_appels_jour_sortants': nb_appels_jour_sortants,
        'ratio_entrants': ratio_entrants,
        'ratio_sortants': ratio_sortants,
        '% tickets': pourcentage_tickets * 100,
        '% Appels': pourcentage_appels * 100,
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
        'activite_categorie': activite_categorie,
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
        })

    return kpis

def convert_to_sixtieth(seconds):
    if pd.isnull(seconds):  # Vérifie si seconds est NaN
        return "Invalid"  # Retourne une valeur par défaut ou un message d'erreur
    minutes, seconds = divmod(int(seconds), 60)  # Convertir en heures
    return f"{int(minutes)}m{int(seconds):02d}s"

def filtrer_par_periode(df_support, periode):
    """ Filtre les données en fonction de la période sélectionnée """
    df_support = df_support.copy()
    df_support['Date'] = pd.to_datetime(df_support['Date'])  # Assure que 'Date' est bien en datetime
    derniere_date = df_support['Date'].max()

    if periode == "1 an":
        date_limite = derniere_date - pd.DateOffset(years=1)
    elif periode == "6 derniers mois":
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

    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_support = df_support[df_support['Semaine'] != 'S2024-01'].copy()

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

    # Tri chronologique des semaines
    ordre_semaines = sorted(data_graph3['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    data_graph3['Semaine'] = pd.Categorical(data_graph3['Semaine'], categories=ordre_semaines, ordered=True)
    data_graph3 = data_graph3.sort_values('Semaine')

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
        "S2025-19": "Stellair MAJ",
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
        title="Graphique avec Taux en barres et Numero_unique/Entrant en aires empilées",
        template="plotly_dark",
        xaxis_title="Semaine",
        yaxis_title="Valeurs",
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

def evo_appels_ticket(df_ticket, df_support):
    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_ticket = df_ticket[df_ticket['Semaine'] != 'S2024-01'].copy()
    df_support = df_support[df_support['Semaine'] != 'S2024-01'].copy()

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

    # Trier les semaines dans l'ordre chronologique
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
        df_support = df_support[df_support['UserName'].isin(["Emilie GEST", 'Morgane Vandenbussche', 'Melinda Marmin', 'Sandrine Sauvage', 'Céline Crendal'])]
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
    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()

    df = df_tickets[df_tickets['Source'].isin(['Chat', 'E-mail', 'Formulaire'])].groupby(['Semaine', 'Pipeline'])['Ticket ID'].nunique().reset_index()

    # Tri chronologique
    ordre_semaines = sorted(df['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df['Semaine'] = pd.Categorical(df['Semaine'], categories=ordre_semaines, ordered=True)
    df = df.sort_values('Semaine')

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

    # Tri chronologique
    ordre_semaines = sorted(df_resultats['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df_resultats['Semaine'] = pd.Categorical(df_resultats['Semaine'], categories=ordre_semaines, ordered=True)
    df_resultats = df_resultats.sort_values('Semaine')

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

def graph_taux_jour(df_support):
    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_support = df_support[df_support['Semaine'] != 'S2024-01'].copy()

    # Vérifier si le DataFrame est vide
    if df_support.empty:
        # Retourner un graphique vide avec un message
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_annotation(
            text="Pas de données disponibles pour cette période",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        fig.update_layout(
            title="Activité & Taux de service / jour",
            template="plotly_dark"
        )
        return fig

    df_support = df_support[df_support['direction'] == 'inbound']

    data_graph2 = df_support.groupby(['Semaine','Date', 'Jour']).agg({'Entrant':'sum',
                                                                'Entrant_connect':'sum',
                                                                'Number':'nunique',
                                                                'Effectif':'mean'})
    
    # Vérifier si data_graph2 est vide après l'agrégation
    if data_graph2.empty:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_annotation(
            text="Pas de données après agrégation",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        fig.update_layout(
            title="Activité & Taux de service / jour",
            template="plotly_dark"
        )
        return fig
    
    data_graph2 = data_graph2.groupby(['Semaine','Date', 'Jour']).agg({'Entrant':'mean',
                                                    'Entrant_connect':'mean',
                                                    'Number':'mean',
                                                    'Effectif':'mean'}).rename(columns = {"Number":"Numero_unique",})
    
    data_graph2['Taux_de_service_support'] = (data_graph2['Entrant_connect'] 
                                            / data_graph2['Entrant'])

    data_graph2 = data_graph2.reset_index()

    data_moyenne = data_graph2.groupby("Jour")[["Taux_de_service_support", "Entrant"]].mean().reset_index()

    # Vérifier si data_graph2 a des données avant d'accéder au dernier élément
    if len(data_graph2) > 0:
        derniere_semaine = data_graph2['Semaine'].iloc[-1]
        data_derniere_semaine = data_graph2.loc[(data_graph2['Semaine'] == derniere_semaine)]
    else:
        derniere_semaine = None
        data_derniere_semaine = pd.DataFrame()

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(x=data_moyenne['Jour'], y=data_moyenne['Taux_de_service_support'], name="Taux de service"),
        secondary_y=False,)

    fig.add_trace(
        go.Scatter(x=data_moyenne['Jour'], y=data_moyenne['Entrant'], mode='lines', name="Entrant"), 
        secondary_y=True,)

    fig.update_yaxes(range = [0,1], secondary_y=False)
    fig.update_yaxes(range=[20,120], secondary_y=True)
    fig.update_yaxes(tickformat=".0%", secondary_y=False)

    fig.update_layout(
        title="Graphique avec Taux en barres et Numero_unique/Entrant en aires empilées",
        template="plotly_dark",
        xaxis_title="Semaine",
        yaxis_title="Valeurs",
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

    fig.update_layout(
        title="Graphique avec Taux en barres et Numero_unique/Entrant en aires empilées",
        template="plotly_dark",
        xaxis_title="Semaine",
        yaxis_title="Valeurs",
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

def df_compute_ticket_appels_metrics(agents_n1, df_ticket, df_support):
    import pandas as pd
    rows = []
    
    for agent in agents_n1:
        # Calcul des KPIs pour chaque agent
        kpis = generate_kpis(df_support, df_ticket, agent, None, 'b2c')

        rows.append({
            'Agent': agent,
            "Nombre d'appel traité": round(kpis.get('nb_appels_jour', 0),2),
            'Nombre de ticket traité': round(kpis.get('moy_ticket_agent', 0),2),
            "% tickets": round(kpis.get('moy_ticket_agent', 0) / (kpis.get('nb_appels_jour', 0) + kpis.get('moy_ticket_agent', 0)),2) *100 if (kpis.get('nb_appels_jour', 0) + kpis.get('moy_ticket_agent', 0)) > 0 else 0,
            "% appels": round(kpis.get('nb_appels_jour', 0) / (kpis.get('nb_appels_jour', 0) + kpis.get('moy_ticket_agent', 0)),2) * 100 if (kpis.get('nb_appels_jour', 0) + kpis.get('moy_ticket_agent', 0)) > 0 else 0,
            '% appel entrant agent': round(kpis.get('ratio_entrants', 0),2) *100,
            # Les champs suivants sont optionnels selon vos besoins :
            #'ref_ticket_agent':round(kpis.get('ref_ticket_agent',0),2),
            #'ref_appel_agent':round(kpis.get('ref_appel_agent',0),2),
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

def calculate_performance_score(row, objectif_total=25, ratio_appels=0.7, ratio_tickets=0.3, objectif_taux_service=0.70):
    """
    Calcule le score de performance d'un agent basé sur plusieurs critères
    
    Parameters:
    -----------
    row : pandas Series
        Ligne de données contenant les métriques de l'agent
    objectif_total : int
        Nombre total de demandes attendues par jour (défaut: 25)
    ratio_appels : float
        Pourcentage attendu d'appels (défaut: 0.7)
    ratio_tickets : float
        Pourcentage attendu de tickets (défaut: 0.3)
    objectif_taux_service : float
        Taux de service cible (défaut: 0.70)
    
    Returns:
    --------
    float
        Score de performance entre 0 et 100
    """
    # 1. Calcul du volume total traité
    nb_appels = row["Nombre d'appel traité"]
    nb_tickets = row["Nombre de ticket traité"]
    volume_total = nb_appels + nb_tickets
    
    # a. Score volume (45%)
    score_volume = min(100, (volume_total / objectif_total) * 100)
    
    # b. Score répartition (25%)
    if volume_total > 0:
        pct_appels = (nb_appels / volume_total) * 100
        pct_tickets = (nb_tickets / volume_total) * 100
    
        # Calcul des écarts absolus
        ecart_appels = abs(pct_appels - (ratio_appels * 100))
        ecart_tickets = abs(pct_tickets - (ratio_tickets * 100))
        ecart_total = ecart_appels + ecart_tickets
        
        # Score répartition = 100 - (écart total × 2)
        score_repartition = max(0, 100 - (ecart_total * 2))
    else:
        score_repartition = 0
    
    # c. Score comparaison à la moyenne (15%)
    # On utilise la moyenne du service passée via l'objectif pour l'instant
    score_comparaison = min(100, (volume_total / objectif_total) * 100)
    
    # d. Score taux d'appels entrants (15%)
    # Calcul du taux d'appels entrants
    if nb_appels > 0:
        taux_appels_entrants = row.get('% appels', 0)  # Pourcentage d'appels entrants
        # Application du seuil de 60%
        if taux_appels_entrants >= 60:
            score_taux_appels_entrants = 100
        else:
            score_taux_appels_entrants = (taux_appels_entrants / 60) * 100
    else:
        score_taux_appels_entrants = 0
    
    # Calcul du score final avec les nouvelles pondérations
    score_final = (
        score_volume * 0.45 +                    # 45% pour le volume
        score_repartition * 0.25 +               # 25% pour la répartition
        score_comparaison * 0.15 +               # 15% pour la comparaison
        score_taux_appels_entrants * 0.15        # 15% pour le taux d'appels entrants
    )
    
    return score_final  # Score entre 0 et 100

def graph_tickets_n2_cumulatif(df_tickets, pipeline_filter=None):
    """
    Crée un graphique cumulatif des tickets N2 (passés au support N2) avec le stock de tickets ouverts par semaine.
    """
    import plotly.express as px
    import plotly.graph_objects as go
    import pandas as pd

    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()

    # Filtrer les tickets passés par le N2
    df_n2 = df_tickets[df_tickets['Passé par le support N2'] == 'Oui'].copy()

    # Obtenir la liste des pipelines disponibles avant de filtrer
    pipelines_disponibles = sorted(df_n2['Pipeline'].unique().tolist())
    
    # Appliquer le filtre par pipeline si spécifié
    if pipeline_filter:
        if isinstance(pipeline_filter, str): pipeline_filter = [pipeline_filter]
        df_n2 = df_n2[df_n2['Pipeline'].isin(pipeline_filter)]

    if df_n2.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucun ticket N2 trouvé pour cette sélection", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig, pipelines_disponibles, pd.DataFrame()

    # --- LOGIQUE DE CALCUL DU STOCK (identique à N1) ---
    df_n2['Date de création'] = pd.to_datetime(df_n2['Date'])
    df_n2['Temps de fermeture'] = pd.to_timedelta(df_n2['Temps de fermeture (HH:mm:ss)'], errors='coerce').fillna(pd.Timedelta(seconds=0))
    df_n2['Date de fermeture'] = df_n2['Date de création'] + df_n2['Temps de fermeture']

    all_weeks = sorted(df_n2['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    weekly_data = []

    for week in all_weeks:
        year, week_num = int(week[1:5]), int(week[6:])
        week_end_date = pd.to_datetime(f'{year}-W{week_num}-0', format='%Y-W%W-%w') + pd.Timedelta(days=6)

        created_so_far = df_n2[df_n2['Date de création'] <= week_end_date]
        closed_so_far = created_so_far[
            (created_so_far['Temps de fermeture'] > pd.Timedelta(seconds=0)) &
            (created_so_far['Date de fermeture'] <= week_end_date)
        ]
        
        open_stock = len(created_so_far) - len(closed_so_far)
        
        weekly_data.append({
            'Semaine': week,
            'Tickets créés (cumul)': len(created_so_far),
            'Tickets fermés (cumul)': len(closed_so_far),
            'Stock de tickets ouverts': open_stock
        })

    df_stock = pd.DataFrame(weekly_data)
    
    if df_stock.empty:
        fig = go.Figure()
        fig.add_annotation(text="Impossible de calculer le stock de tickets N2", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig, pipelines_disponibles, pd.DataFrame()
        
    # --- Création du graphique avec axe secondaire ---
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # Ajouter les courbes cumulatives sur l'axe Y principal
    fig.add_trace(go.Scatter(x=df_stock['Semaine'], y=df_stock['Tickets créés (cumul)'], mode='lines', name='Tickets N2 créés (cumulatif)', line=dict(color='blue')), secondary_y=False)
    fig.add_trace(go.Scatter(x=df_stock['Semaine'], y=df_stock['Tickets fermés (cumul)'], mode='lines', name='Tickets N2 fermés (cumulatif)', line=dict(color='green')), secondary_y=False)

    # Ajouter la courbe de stock sur l'axe Y secondaire
    fig.add_trace(go.Scatter(x=df_stock['Semaine'], y=df_stock['Stock de tickets ouverts'], mode='lines', name='Stock de tickets N2 ouverts', line=dict(color='red', dash='dash')), secondary_y=True)

    title = "Stock de tickets N2 ouverts par semaine"
    if pipeline_filter:
        title += f" - Pipeline(s): {', '.join(pipeline_filter)}"

        fig.update_layout(
        title=title,
        xaxis_title="Semaine",
        template="plotly_dark",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    # Définir les titres des axes Y
    fig.update_yaxes(title_text="Total tickets (cumulatif)", secondary_y=False)
    fig.update_yaxes(title_text="Stock de tickets ouverts", secondary_y=True)
    
    last_week_end = pd.to_datetime(f'{all_weeks[-1][1:5]}-W{all_weeks[-1][6:]}-0', format='%Y-W%W-%w') + pd.Timedelta(days=6)
    tickets_ouverts_fin_periode = df_n2[
        (df_n2['Date de création'] <= last_week_end) & 
        ((df_n2['Temps de fermeture'] == pd.Timedelta(seconds=0)) | (df_n2['Date de fermeture'] > last_week_end))
    ]

    return fig, pipelines_disponibles, tickets_ouverts_fin_periode

def graph_tickets_ouverts_pierre_goupillon(df_tickets):
    """
    Crée un graphique du nombre de tickets OUVERTS de Pierre Goupillon par semaine (basé sur Date de création + Temps de fermeture).
    """
    import plotly.express as px
    import plotly.graph_objects as go
    import pandas as pd
    
    # Exclure la semaine S2024-01 qui fausse les graphiques
    df_tickets = df_tickets[df_tickets['Semaine'] != 'S2024-01'].copy()
    
    # Gestion des variantes de nom pour Pierre Goupillon
    correspondance_pierre = ['Pierre Goupillon', 'Pierre GOUPILLON']
    
    # Filtrer avec toutes les variantes possibles
    df_pierre = df_tickets[df_tickets['Propriétaire du ticket'].isin(correspondance_pierre)].copy()
    
    if df_pierre.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucun ticket pour Pierre Goupillon", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(title="Tickets ouverts de Pierre Goupillon par semaine", template="plotly_dark")
        return fig

    # --- NOUVELLE LOGIQUE DE CALCUL DU STOCK ---
    # 1. Préparation des dates
    df_pierre['Date'] = pd.to_datetime(df_pierre['Date'])
    
    # Obtenir la date de création (la plus ancienne date pour chaque ticket)
    creation_dates = df_pierre.groupby('Ticket ID')['Date'].min().reset_index().rename(columns={'Date': 'Date de création'})
    
    # Obtenir le temps de fermeture depuis la dernière entrée de chaque ticket
    latest_entries = df_pierre.loc[df_pierre.groupby('Ticket ID')['Date'].idxmax()][['Ticket ID', 'Temps de fermeture (HH:mm:ss)']]
    
    # Combiner pour avoir une vue unique et propre par ticket
    df_pierre_summary = pd.merge(creation_dates, latest_entries, on='Ticket ID', how='left')

    # Calculer la date de fermeture réelle
    df_pierre_summary['Temps de fermeture'] = pd.to_timedelta(df_pierre_summary['Temps de fermeture (HH:mm:ss)'], errors='coerce')
    df_pierre_summary['Date de fermeture'] = df_pierre_summary['Date de création'] + df_pierre_summary['Temps de fermeture']

    # Obtenir toutes les semaines uniques et les trier
    all_weeks = sorted(df_pierre['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    weekly_open_counts = []

    for week in all_weeks:
        # Déterminer la date de fin de la semaine
        year, week_num = int(week[1:5]), int(week[6:])
        week_end_date = pd.to_datetime(f'{year}-W{week_num}-0', format='%Y-W%W-%w') + pd.Timedelta(days=6)
        
        # 1. Sélectionner les tickets créés avant la fin de cette semaine
        tickets_created_so_far = df_pierre_summary[df_pierre_summary['Date de création'] <= week_end_date]
        
        # 2. Compter ceux qui sont encore ouverts à cette date
        # Un ticket est ouvert si sa date de fermeture est future OU si elle n'existe pas (NaT)
        open_count = tickets_created_so_far[
            (tickets_created_so_far['Date de fermeture'] > week_end_date) | 
            (pd.isna(tickets_created_so_far['Date de fermeture']))
        ].shape[0]
        
        weekly_open_counts.append({'Semaine': week, 'Tickets ouverts': open_count})

    df_open_counts = pd.DataFrame(weekly_open_counts)
    
    if df_open_counts.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucun ticket ouvert pour Pierre Goupillon", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(title="Tickets ouverts de Pierre Goupillon par semaine", template="plotly_dark")
        return fig

    # Tri chronologique des semaines
    ordre_semaines = sorted(df_open_counts['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    df_open_counts['Semaine'] = pd.Categorical(df_open_counts['Semaine'], categories=ordre_semaines, ordered=True)
    df_open_counts = df_open_counts.sort_values('Semaine')
    
    fig = px.bar(
        df_open_counts,
        x='Semaine',
        y='Tickets ouverts',
        title="Stock de tickets ouverts de Pierre Goupillon par semaine",
        labels={'Tickets ouverts': 'Nombre de tickets ouverts'}
    )

    fig.update_xaxes(
        categoryorder='array',
        categoryarray=ordre_semaines,
        tickangle=-45
    )

    fig.update_layout(
        template="plotly_dark"
    )

    return fig

def graph_tickets_n1_cumulatif(df_tickets, agent_filter=None):
    """
    Crée un graphique cumulatif des tickets N1 (pipelines SSI, SSIA, SPSA) avec le stock de tickets ouverts par semaine.
    """
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    import pandas as pd

    # --- FILTRAGE INITIAL ---
    # 1. Filtrer par source
    sources_a_inclure = ['Chat', 'E-mail', 'Formulaire']
    df_filtered = df_tickets[df_tickets['Source'].isin(sources_a_inclure)].copy()
    
    # 2. Filtrer les tickets des pipelines N1 (SSI, SSIA, SPSA)
    df_n1 = df_filtered[df_filtered['Pipeline'].isin(['SSI', 'SSIA', 'SPSA'])].copy()

    # Obtenir la liste des agents disponibles avant de filtrer
    agents_disponibles = sorted(df_n1['Propriétaire du ticket'].unique().tolist())
    
    # Appliquer le filtre par agent si spécifié
    if agent_filter:
        if isinstance(agent_filter, str): agent_filter = [agent_filter]
        df_n1 = df_n1[df_n1['Propriétaire du ticket'].isin(agent_filter)]

    if df_n1.empty:
        fig = go.Figure()
        fig.add_annotation(text="Aucun ticket N1 trouvé pour cette sélection", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig, agents_disponibles, pd.DataFrame()

    # --- NOUVELLE LOGIQUE DE CALCUL DU STOCK ---
    # 1. Préparation des dates
    df_n1['Date de création'] = pd.to_datetime(df_n1['Date'])
    
    # Convertir 'Temps de fermeture (HH:mm:ss)' en timedelta
    # Remplacer les valeurs non valides (NaN, etc.) par un timedelta nul
    df_n1['Temps de fermeture'] = pd.to_timedelta(df_n1['Temps de fermeture (HH:mm:ss)'], errors='coerce').fillna(pd.Timedelta(seconds=0))
    
    # 2. Calculer la date de fermeture
    df_n1['Date de fermeture'] = df_n1['Date de création'] + df_n1['Temps de fermeture']
    # Les tickets non fermés (temps de fermeture = 0) auront une date de fermeture égale à la date de création.
    # Nous les identifierons car leur temps de fermeture est de 0.

    # 3. Calculer les métriques par semaine
    all_weeks = sorted(df_n1['Semaine'].unique(), key=lambda x: (int(x[1:5]), int(x[6:])))
    weekly_data = []

    for week in all_weeks:
        # Extraire l'année et le numéro de semaine pour calculer la date de fin de semaine
        year, week_num = int(week[1:5]), int(week[6:])
        # Le + '6' pour obtenir le dimanche (fin de semaine)
        week_end_date = pd.to_datetime(f'{year}-W{week_num}-0', format='%Y-W%W-%w') + pd.Timedelta(days=6)

        # Tickets créés jusqu'à la fin de cette semaine
        created_so_far = df_n1[df_n1['Date de création'] <= week_end_date]
        
        # Tickets fermés jusqu'à la fin de cette semaine
        # Un ticket est considéré comme fermé si son temps de fermeture n'est pas nul ET que sa date de fermeture est passée
        closed_so_far = created_so_far[
            (created_so_far['Temps de fermeture'] > pd.Timedelta(seconds=0)) &
            (created_so_far['Date de fermeture'] <= week_end_date)
        ]
        
        # Le stock est la différence
        open_stock = len(created_so_far) - len(closed_so_far)
        
        weekly_data.append({
            'Semaine': week,
            'Tickets créés (cumul)': len(created_so_far),
            'Tickets fermés (cumul)': len(closed_so_far),
            'Stock de tickets ouverts': open_stock
        })

    df_stock = pd.DataFrame(weekly_data)
    
    if df_stock.empty:
        fig = go.Figure()
        fig.add_annotation(text="Impossible de calculer le stock de tickets", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig, agents_disponibles, pd.DataFrame()
        
    # --- Création du graphique avec axe secondaire ---
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # Ajouter les courbes cumulatives sur l'axe Y principal
    fig.add_trace(go.Scatter(x=df_stock['Semaine'], y=df_stock['Tickets créés (cumul)'], mode='lines', name='Tickets créés (cumulatif)', line=dict(color='blue')), secondary_y=False)
    fig.add_trace(go.Scatter(x=df_stock['Semaine'], y=df_stock['Tickets fermés (cumul)'], mode='lines', name='Tickets fermés (cumulatif)', line=dict(color='green')), secondary_y=False)
    
    # Ajouter la courbe de stock sur l'axe Y secondaire
    fig.add_trace(go.Scatter(x=df_stock['Semaine'], y=df_stock['Stock de tickets ouverts'], mode='lines', name='Stock de tickets ouverts', line=dict(color='red', dash='dash')), secondary_y=True)

    title = "Stock de tickets N1 ouverts par semaine"
    if agent_filter:
        title += f" - Agent(s): {', '.join(agent_filter)}"

    fig.update_layout(
        title=title,
        xaxis_title="Semaine",
        template="plotly_dark",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    # Définir les titres des axes Y
    fig.update_yaxes(title_text="Total tickets (cumulatif)", secondary_y=False)
    fig.update_yaxes(title_text="Stock de tickets ouverts", secondary_y=True)
    
    # La notion de "tickets en cours" à télécharger peut maintenant être basée sur le dernier snapshot
    last_week_end = pd.to_datetime(f'{all_weeks[-1][1:5]}-W{all_weeks[-1][6:]}-0', format='%Y-W%W-%w') + pd.Timedelta(days=6)
    tickets_ouverts_fin_periode = df_n1[
        (df_n1['Date de création'] <= last_week_end) & 
        ((df_n1['Temps de fermeture'] == pd.Timedelta(seconds=0)) | (df_n1['Date de fermeture'] > last_week_end))
    ]

    return fig, agents_disponibles, tickets_ouverts_fin_periode
