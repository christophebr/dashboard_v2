import os
import markdown
import base64
from io import BytesIO
import pandas as pd
import plotly.io as pio

from data_processing.kpi_generation import (
    generate_kpis, convert_to_sixtieth, graph_activite, graph_taux_jour, graph_taux_heure,
    filtrer_par_periode, charge_entrant_sortant, charge_ticket, historique_scores_total,
    df_compute_ticket_appels_metrics, score_ticket, score_appel
)
from data_processing.aircall_processing import process_aircall_data, def_df_support, agents_all, line_tous
from data_processing.hubspot_processing import process_hubspot_data
from utils.streamlit_helpers import load_data

# Configuration
pio.kaleido.scope.default_format = "png"
os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib'

# Chargement des données
df_aircall, df_hubspot, df_evaluation = load_data()
df_support = def_df_support(process_aircall_data(df_aircall), process_aircall_data(df_aircall), line_tous, agents_all)
df_tickets = process_hubspot_data(df_hubspot)


def fig_to_base64_img(fig):
    buffer = BytesIO()
    fig.write_image(buffer, format="png", scale=2)
    encoded_image = base64.b64encode(buffer.getvalue()).decode("utf-8")
    return f"data:image/png;base64,{encoded_image}"

def generate_html_report(kpis, df_support, periode):
    # Prépare les images encodées
    image_path_fig1 = fig_to_base64_img(graph_activite(df_support))
    image_path_fig2 = fig_to_base64_img(graph_taux_jour(df_support))
    image_path_fig3 = fig_to_base64_img(graph_taux_heure(df_support))
    image_path_fig4 = fig_to_base64_img(kpis['evo_appels_tickets'])
    image_path_fig5 = fig_to_base64_img(kpis['fig_activite_ticket'])
    image_path_fig6 = fig_to_base64_img(kpis['charge_affid_stellair_%'])

    # Construction HTML
    html_content = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{ size: A4 landscape; margin: 1cm; }}
            body {{ font-family: 'Segoe UI', sans-serif; padding: 20px; }}
            h1, h2 {{ text-align: center; }}
            img {{ max-width: 90%; display: block; margin: 20px auto; }}
            .kpi-container {{ display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; }}
            .kpi-card {{ border: 1px solid #ccc; border-radius: 8px; padding: 10px; width: 200px; text-align: center; }}
            footer {{ margin-top: 50px; text-align: center; color: #aaa; }}
        </style>
    </head>
    <body>
        <h1>📊 Rapport d'Activité - Support</h1>
        <h2>Période : {periode}</h2>
        <div class="kpi-container">
            <div class="kpi-card"><strong>Taux de service</strong><br>{kpis.get('Taux_de_service', 'N/A')}%</div>
            <div class="kpi-card"><strong>Appels entrants / jour</strong><br>{kpis.get('Entrant', 'N/A')}</div>
            <div class="kpi-card"><strong>N° appelants uniques</strong><br>{kpis.get('Numero_unique', 'N/A')}</div>
            <div class="kpi-card"><strong>Temps moyen / appel</strong><br>{convert_to_sixtieth(kpis.get('temps_moy_appel', 0))}</div>
            <div class="kpi-card"><strong>Appels / agent / jour</strong><br>{round(kpis.get('Nombre_appel_jour_agent', 0), 2)}</div>
            <div class="kpi-card"><strong>Entrants vs Tickets (%)</strong><br>{round(kpis.get('activite_appels_pourcentage', 0)*100, 2)}% / {round(kpis.get('activite_tickets_pourcentage', 0)*100, 2)}%</div>
        </div>
        <h2>Graphique d'Activité</h2>
        <img src="{image_path_fig1}" alt="Graphique d'Activité">
        <h2>Graphique par jour</h2>
        <img src="{image_path_fig2}" alt="Jour">
        <h2>Graphique par heure</h2>
        <img src="{image_path_fig3}" alt="Heure">
        <h2>Appels vs Tickets</h2>
        <img src="{image_path_fig4}" alt="Évolution appels/tickets">
        <h2>Activité Tickets</h2>
        <img src="{image_path_fig5}" alt="Activité tickets">
    </body>
    </html>
    """

    # Détermine le chemin absolu sûr

    root_dir = os.getcwd()
    docker_dir = os.path.join(root_dir, "rapport", "docker")
    html_file_path = os.path.abspath(os.path.join(docker_dir, "rapport_agents.html"))

    # Écrit le fichier HTML
    with open(html_file_path, "w", encoding="utf-8") as file:
        file.write(html_content)

    print(f"✅ HTML généré : {html_file_path}")

    import subprocess
    # Génération PDF via Docker
    try:
        subprocess.run([
            "docker", "run", "--rm",
            "-v", f"{docker_dir}:/app",
            "weasyprint-pdf",
            "python", "-c",
            "from weasyprint import HTML; HTML('/app/rapport_agents.html').write_pdf('/app/rapport_agents.pdf')"
        ], check=True)
        print("✅ PDF généré avec succès !")
    except subprocess.CalledProcessError as e:
        print("❌ Erreur lors de la génération PDF :", e)

    return html_file_path  # retourne le chemin pour que l'appelant puisse l'afficher




def generate_html_report_agent(df_support, df_tickets, periode, df_evaluation_filre):
    agents_n1 = ['Olivier Sainte-Rose', 'Mourad HUMBLOT', 'Archimede KESSI', 'Morgane Vandenbussche']
    agents_n1_tickets = agents_n1 + ['Frederic SAUVAN']

    df_support_filtered = filtrer_par_periode(df_support, periode)
    df_tickets_filtered = filtrer_par_periode(df_tickets, periode)

    pio.templates.default = "plotly_white"

    fig_line_entrant, fig_pie_entrant = charge_entrant_sortant(df_support_filtered, agents_n1)
    fig_line_ticket, fig_pie_ticket = charge_ticket(df_tickets_filtered, agents_n1_tickets)
    fig_historique = historique_scores_total(agents_n1, df_tickets_filtered, df_support_filtered)

    image_path_fig1 = fig_to_base64_img(fig_line_entrant)
    image_path_fig2 = fig_to_base64_img(fig_pie_entrant)
    image_path_fig3 = fig_to_base64_img(fig_line_ticket)
    image_path_fig4 = fig_to_base64_img(fig_pie_ticket)
    image_path_fig5 = fig_to_base64_img(fig_historique)

    df_conforme = df_compute_ticket_appels_metrics(agents_n1, df_tickets_filtered, df_support_filtered)
    df_conforme['score ticket'] = df_conforme.apply(score_ticket, axis=1)
    df_conforme['score appel'] = df_conforme.apply(score_appel, axis=1)
    df_conforme['score total'] = (df_conforme['score ticket'] + df_conforme['score appel']) / 2

    df_conforme = df_conforme[[ 'Agent', 'score total', 'score appel', 'score ticket',
                                "Nombre d'appel traité", 'Nombre de ticket traité', 'ref_ticket_agent',
                                'ref_appel_agent', '% appel entrant agent']]

    def style_scores(df):
        def style_score_ticket(val):
            if val < 0.50:
                return 'background-color: #f28e8e'
            elif val < 0.60:
                return 'background-color: #f7c97f'
            else:
                return 'background-color: #a7dba7'

        def style_score_appel(val, percent_entrant):
            if abs(percent_entrant) < 0.50 or val < 0.50:
                return 'background-color: #f28e8e'
            elif val < 0.60:
                return 'background-color: #f7c97f'
            else:
                return 'background-color: #a7dba7'

        def style_score_total(ticket, appel, total):
            if abs(ticket - appel) > 0.30 or total < 0.50:
                return 'background-color: #f28e8e'
            elif total < 0.60:
                return 'background-color: #f7c97f'
            else:
                return 'background-color: #a7dba7'

        styled_ticket = df['score ticket'].apply(style_score_ticket)
        styled_appel = [
            style_score_appel(appel, percent)
            for appel, percent in zip(df['score appel'], df['% appel entrant agent'])
        ]
        styled_total = [
            style_score_total(ticket, appel, total)
            for ticket, appel, total in zip(df['score ticket'], df['score appel'], df['score total'])
        ]

        style_df = pd.DataFrame('', index=df.index, columns=df.columns)
        style_df['score ticket'] = styled_ticket
        style_df['score appel'] = styled_appel
        style_df['score total'] = styled_total

        return df.style.apply(lambda _: style_df, axis=None)

    # Génère le HTML stylisé du tableau
    styled_df_html = (
        style_scores(df_conforme)
        .set_table_attributes('class="styled-table"')
        .to_html()
    )


    # Texte Markdown pour les règles
    markdown_rules = """
    ### 🎯 Règles de scoring des performances - Agent Niveau 1

    **Les scores sont évalués selon les seuils suivants :**

    - 🟢 **Vert** : score ≥ 0.60
    - 🟠 **Orange** : 0.50 ≤ score < 0.60
    - 🔴 **Rouge** : score < 0.50

    #### 📞 Spécificité appels :
    - Si le **% d'appels entrants** est inférieur à **50%**, le score est considéré comme 🔴 **Rouge**.

    #### ⚖️ Score global :
    - Si l’**écart entre le score ticket et le score appel dépasse 0.30**, le **score total est marqué en rouge**, quelle que soit sa valeur.
    """
    markdown_html = markdown.markdown(markdown_rules)

    # HTML complet
    html_content = f"""
    <html>
    <head>
        <meta charset='UTF-8'>
        <style>
            @page {{
                size: A4 landscape;
                margin: 2cm;
            }}
            body {{
                font-family: 'Segoe UI', sans-serif;
                padding: 20px;
                font-size: 12px;
            }}
            img {{
                max-width: 90%;
                margin: 20px auto;
                display: block;
                page-break-inside: avoid;
            }}
            h1, h2 {{
                text-align: center;
                page-break-after: avoid;
            }}
            table.styled-table {{
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
                word-wrap: break-word;
                page-break-inside: avoid;
                font-size: 10px;
            }}
            .styled-table th, .styled-table td {{
                border: 1px solid #ddd;
                padding: 6px;
                text-align: center;
                overflow: hidden;
                white-space: normal;
            }}
            .styled-table th {{
                background-color: #f2f2f2;
            }}
            .styled-table tbody tr {{
                page-break-inside: avoid;
                page-break-after: auto;
            }}
        </style>
    </head>
    <body>
        <h1>📋 Rapport par Agent - Support Niveau 1</h1>
        <h2>Période : {periode}</h2>
        <img src="{image_path_fig1}" alt="Charge appels">
        <img src="{image_path_fig2}" alt="Répartition appels">
        <img src="{image_path_fig3}" alt="Charge tickets">
        <img src="{image_path_fig4}" alt="Répartition tickets">
        <img src="{image_path_fig5}" alt="Historique scores">
        <h2>Scoring des Agents</h2>
        <div style="transform: scale(0.95); transform-origin: top center;">
            {styled_df_html}
        </div>
        <h2>Évaluations Managériales</h2>
        <div style="transform: scale(0.95); transform-origin: top center;">
            {df_evaluation_filre_html}
        </div>
        <div>{markdown_html}</div>
    </body>
    </html>
    """

    # Écrit le fichier HTML dans le dossier Docker
    root_dir = os.getcwd()
    docker_dir = os.path.join(root_dir, "rapport", "docker")
    html_file_path = os.path.abspath(os.path.join(docker_dir, "rapport_scrores_agents.html"))

    with open(html_file_path, "w", encoding="utf-8") as file:
        file.write(html_content)

    print(f"✅ HTML généré : {html_file_path}")

    # Appelle Docker pour générer le PDF avec WeasyPrint
    import subprocess
    try:
        subprocess.run([
            "docker", "run", "--rm",
            "-v", f"{docker_dir}:/app",
            "weasyprint-pdf",
            "python", "-c",
            "from weasyprint import HTML; HTML('/app/rapport_scrores_agents.html').write_pdf('/app/rapport_scrores_agents.pdf')"
        ], check=True)
        print("✅ PDF généré avec succès !")
    except subprocess.CalledProcessError as e:
        print("❌ Erreur lors de la génération PDF :", e)

    return html_file_path
