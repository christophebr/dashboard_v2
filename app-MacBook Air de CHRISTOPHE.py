import streamlit as st
from utils.authentification import authenticate_user
from data_processing.aircall_processing import process_aircall_data, def_df_support, agents_all, line_tous
from data_processing.hubspot_processing import process_hubspot_data
from data_processing.kpi_generation import generate_kpis
from utils.streamlit_helpers import setup_page, load_data
import config
from config import CREDENTIALS
import streamlit_authenticator as stauth

st.set_page_config(
    page_title=":bar_chart: Dashboard support",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Authentification de l'utilisateur
authenticator = stauth.Authenticate(
    config.CREDENTIALS,
    'dashboard_support',
    config.COOKIE_KEY,  # Assurez-vous que cette clé est unique et sécurisée
    cookie_expiry_days=2
)

# Utilisation du nouvel argument 'fields'
name, authentification_status, username = authenticator.login('main', fields={'Form name': 'custom_form_name'})

if authentification_status == False:
    st.error('Nom d\'utilisateur/mot de passe incorrect')
elif authentification_status == None:
    st.warning('Veuillez entrer votre nom d\'utilisateur et mot de passe')
elif authentification_status:
    st.sidebar.title(f"Bienvenue, {name}")
    authenticator.logout('Logout', 'sidebar')

    # Chargement des données
    df_aircall, df_hubspot = load_data()

    # Traitement des données
    df_support = def_df_support(process_aircall_data(df_aircall), process_aircall_data(df_aircall), line_tous, agents_all)
    df_tickets = process_hubspot_data(df_hubspot)

    # Génération des KPIs
    kpis = generate_kpis(df_support, df_tickets, 'all')

    PAGES = {
        "Support": "support",
        "Agents": "agents",
        "Tickets": "tickets"
    }

    selection_page = st.sidebar.selectbox("Choix de la page", list(PAGES.keys()), key="unique_page_selection")

    if selection_page == "Support":
        from data_processing.kpi_generation import filtrer_par_periode
        from data_processing.aircall_processing import process_aircall_data, def_df_support, agents_all, line_tous, agents_support, line_support, line_armatis, agents_armatis
        from data_processing.kpi_generation import generate_kpis, convert_to_sixtieth, graph_activite , graph_taux_jour, graph_taux_heure


        def support():

            dataframe_option = st.sidebar.selectbox("Choisir le dataframe", ["support_suresnes", "support_armatis", "support_stellair", "support_affid"], key="unique_dataframe_selection")

            df_stellair = def_df_support(process_aircall_data(df_aircall), process_aircall_data(df_aircall), line_tous, agents_all)
            df_stellair = df_stellair[(df_stellair['line'] == 'armatistechnique') 
                                          | (df_stellair['IVR Branch'] == 'Stellair')]

            df_affid = def_df_support(process_aircall_data(df_aircall), process_aircall_data(df_aircall), line_tous, agents_support)
            df_affid = df_affid[df_affid['IVR Branch'] == 'Affid']

            dataframe_config = {
                    "support_suresnes": {
                        "df": def_df_support(process_aircall_data(df_aircall), process_aircall_data(df_aircall), line_support, agents_support),
                        "agents": agents_support
                    },
                    "support_armatis": {
                        "df": def_df_support(process_aircall_data(df_aircall), process_aircall_data(df_aircall), line_armatis, agents_armatis),
                        "agents": agents_armatis
                    },
                    "support_stellair": {
                        "df": df_stellair,
                        "agents": agents_all
                    },
                    "support_affid": {
                        "df": df_affid,
                        "agents": agents_support
                    }
                }

                # Chargement du dataframe sélectionné
            df_support = dataframe_config[dataframe_option]["df"]
            agents = dataframe_config[dataframe_option]["agents"]

            periode = st.selectbox("Sélectionnez une période :", 
                    ["Toute la période", "3 derniers mois", "Dernier mois", "Dernière semaine"])
            
            df_support = filtrer_par_periode(df_support, periode)

            kpis = generate_kpis(filtrer_par_periode(df_support, periode), df_tickets, 'all')

            col1, col2, col3 = st.columns(3)
            col1.metric("Taux de service en %", kpis['Taux_de_service'])
            col2.metric("Appels entrant / Jour", kpis['Entrant'])
            col3.metric("Numéros uniques / Jour", kpis['Numero_unique'])

            col4, col5, col6 = st.columns(3)
            col4.metric("Temps Moy / Appel", convert_to_sixtieth(kpis['temps_moy_appel']))
            col5.metric("Nombre d'appels / Agent / Jour", kpis['Nombre_appel_jour_agent'])
            col6.metric("Taux de réponse (%)", round(kpis['taux_reponse']* 100))


            st.plotly_chart(graph_activite(df_support), use_container_width=True)

            # Affichage des graphiques de taux
            col_graph1, col_graph2 = st.columns(2)
            col_graph1.plotly_chart(graph_taux_jour(df_support), use_container_width=True)
            col_graph2.plotly_chart(graph_taux_heure(df_support), use_container_width=True)

        support()



    elif selection_page == "Agents":
        import pandas as pd
        from data_processing.kpi_generation import filtrer_par_periode, generate_kpis, convert_to_sixtieth, filtrer_par_agent, charge_entrant_sortant
        from data_processing.hubspot_processing import process_hubspot_data
        from data_processing.aircall_processing import process_aircall_data, def_df_support, agents_all, line_tous

        agent = st.selectbox("Sélectionnez un agent :", ["Olivier", "Mourad", "Frederic", "Archimède", 
                                                         "Christophe", "Pierre", "Emilie", "Morgane",
                                                           "Melinda", "Sandrine"], key="unique_agent_selection")
        

        periode = st.selectbox("Sélectionnez une période :", 
                ["Toute la période", "3 derniers mois", "Dernier mois", "Dernière semaine"])
        
        df_support = def_df_support(process_aircall_data(df_aircall), process_aircall_data(df_aircall), line_tous, agents_all)
        df_tickets = process_hubspot_data(filtrer_par_periode(df_tickets, periode))
        kpis = generate_kpis(filtrer_par_periode(df_support, periode), df_tickets, filtrer_par_agent(agent))

    


        col1, col2, col3, col4, col5 = st.columns(5) 
        col1.metric(f'{agent} - Com Moy / Jour',kpis['com_jour'])
        col2.metric(f'{agent} - Temps Moy / Appel',convert_to_sixtieth(kpis['temps_moy_com']))
        col3.metric(f'{agent} - Nb Appels / Jour', round(kpis['nb_appels_jour'], 2))
        col4.metric(f'{agent} - Nb Tickets / Jour', round(kpis['df_tickets_agent'], 2))

        col5.metric(
            f"{agent} - Répartition Entrants / Sortants",
            f"{int(kpis['ratio_entrants'] * 100)}% / {int(kpis['ratio_sortants'] * 100)}%")

        st.plotly_chart(charge_entrant_sortant(filtrer_par_periode(df_support, periode), filtrer_par_agent(agent)), use_container_width=True)


    elif selection_page == "Tickets":
        partenaire = st.selectbox("Sélectionnez un partenaire :", ["Follow", "Odaiji", "Help Info", "Oppysoft"], key="unique_partenaire_selection")
