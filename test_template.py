import os
from utils.olaqin_template import OlaqinTemplate

def test_presentation():
    # Créer le template
    template = OlaqinTemplate()
    
    # Slide de titre
    template.add_title_slide(
        "Test du Template Olaqin",
        "Vérification des images et du style"
    )
    
    # Slide de contenu avec tableau
    slide = template.add_content_slide("Test du Tableau")
    table = template.create_table(slide, rows=3, cols=3)
    
    # Remplir le tableau
    headers = ["Colonne 1", "Colonne 2", "Colonne 3"]
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
    
    # Formater l'en-tête
    template.format_table_header(table)
    
    # Ajouter quelques données avec OK/NOK
    data = [
        ["Test 1", "OK", "NOK"],
        ["Test 2", "NOK", "OK"]
    ]
    
    for i, row in enumerate(data):
        for j, value in enumerate(row):
            cell = table.cell(i + 1, j)
            if value in ["OK", "NOK"]:
                template.add_ok_nok(cell, value == "OK")
            else:
                cell.text = value
    
    # Sauvegarder la présentation
    template.prs.save("test_template.pptx")
    print("Présentation de test générée : test_template.pptx")

if __name__ == "__main__":
    test_presentation() 