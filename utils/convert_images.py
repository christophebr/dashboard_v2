from PIL import Image
import os

def convert_to_png():
    assets_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'assets')
    
    # Liste des fichiers à convertir
    files_to_convert = [
        'olaqin_hex',
        'olaqin_hex_small',
        'olaqin_logo_small'
    ]
    
    for base_name in files_to_convert:
        # Chercher le fichier JPEG
        for ext in ['.jpeg', '.jpg']:
            jpeg_path = os.path.join(assets_dir, f'{base_name}{ext}')
            if os.path.exists(jpeg_path):
                # Créer le chemin pour le fichier PNG
                png_path = os.path.join(assets_dir, f'{base_name}.png')
                
                # Ouvrir et convertir l'image
                with Image.open(jpeg_path) as img:
                    # Convertir en mode RGB si nécessaire
                    if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                        img = img.convert('RGBA')
                    else:
                        img = img.convert('RGB')
                    
                    # Sauvegarder en PNG
                    img.save(png_path, 'PNG', quality=95)
                    print(f'Converti : {base_name}{ext} -> {base_name}.png')
                
                # Supprimer le fichier JPEG original
                os.remove(jpeg_path)
                print(f'Supprimé : {base_name}{ext}')
                break

if __name__ == '__main__':
    convert_to_png() 