import pandas as pd     # pour la manipulation de donnees et la conversion vers Excel
import os               # pour la manipulation des fichiers et des repertoires
import argparse         # pour analyser les arguments de la ligne de commande
import chardet          # pour detecter l'encodage des fichiers texte
#import shutil           # pour la manipulation des fichiers et des repertoires
from datetime import datetime


def normalize_date(date_str):
    """
    Normalise une date au format jj/mm/aaaa.
    
    :param date_str: La chaîne de caractères représentant la date.
    :return: La date normalisée au format jj/mm/aaaa ou une chaîne vide si la date n'est pas valide.
    """
    if not date_str.strip():
        return ""
    
    # Essayer de détecter et convertir les formats connus
    for fmt in ("%Y%m%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    
    # Si aucun format n'a fonctionné, retourner la chaîne originale
    return ""


# Fonction pour appliquer la normalisation sur une colonne de données
def normalize_dates_in_column(column):
    """
    Applique la normalisation des dates sur une colonne de données.
    
    :param column: La liste des dates à normaliser.
    :return: La liste des dates normalisées.
    """
    return [normalize_date(date) for date in column if date != 'InstallDate']






def clear_temp_directory(directory_path):
    ''' 
    Cette fonction supprime tous les fichiers du repertoire

    Argument:
        directory_path -- chemin du repertoire
    '''
    try:
        # si le repertoire existe
        if os.path.exists(directory_path):

            # on boucle sur tous les fichiers du repertoire
            for filename in os.listdir(directory_path):

                # on construit le nom de fichier a partir du nom du repertoire
                file_path = os.path.join(directory_path, filename)

                try:
                    # si c'est un vrai fichier ou un lien symbolique sur un fichier
                    if os.path.isfile(file_path) or os.path.islink(file_path):

                        # on l'efface
                        os.unlink(file_path)    

                # si un probleme on leve l'exception
                except Exception as e:
                    print(f'Erreur lors de la suppression de {file_path}. Raison: {e}')

        else:
            print(f'Le répertoire {directory_path} n\'existe pas.')

    except Exception as e:
        print(f'Erreur lors de la vérification du répertoire {directory_path}. Raison: {e}')


def detect_encoding(file_path):
    ''' 
    Cette fonction detecte l'encodage d'un fichier en utilisant 'chardet'
    
    Argument:
    file_path -- le chemin du fichier a analyser

    Retourne:
    l'encodage du fichier
    '''

    with open(file_path, 'rb') as f:            # Ouvre le fichier binaire en lecture
        result = chardet.detect(f.read())       # Detecte l'encodage du fichier

    return result['encoding']


def read_file_lines(file_path, encoding):

    ''' 
    Ouvre le fichier en lecture avec le bon encodage et
    creer une liste avec toutes les lignes

    Argument:
    --------
        file_path -- le chemin du fichier a lire
        encoding  -- l'encodage du fichier
 
    Retourne:
    --------
        une liste de lignes lu dans le fichier
    '''

    with open(file_path, 'r', encoding=encoding) as f:
        lines = f.readlines()       

    return lines


def extract_columns(input_file, col_specs):
    ''' 
    Cette fonction extrait des colonnes specifiques du fichier texte

    Arguments:
        input_file --  le chemin vers le fichier texte
        col_specs  --  numero de la colonne a extraire

    Retourne:
        la liste des colonnes extraites
    '''
    
    # on recupere l'encodage du fichier texte
    encoding = detect_encoding(input_file)          

    # on recupere la liste des lignes a partir du fichier texte
    lines = read_file_lines(input_file, encoding)   
    
    # creation d'une liste de sous liste vide qu'il y a d'elements dans col_specs
    cols = [[] for _ in col_specs]

    # pour chaque ligne dans la liste lignes sauf la second ligne
    for i, line in enumerate(lines):

        if i == 1:      # on n evalue pas la seconde ligne du fichier
            continue

        # Pour chacune des 4 colonnes on extrait debut et fin
        for j, (start, end) in enumerate(col_specs):
            cols[j].append(line[start:end].strip())

    return cols





def save_columns_to_files(columns, output_files):
    ''' 
    Cette fonction enregistre dans des fichiers temporaires les colonnes,
    un fichier par colonne en lisant une ligne dans chaque fichier grace a zip

    Arguments:
        columns     -- la liste des colonnes
        output_file -- la liste des noms de fichiers colonne
    ''' 
    for col, output_file in zip(columns, output_files):
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(col) + '\n')


def merge_columns_to_csv(output_file, col_files, delimiter):
    ''' 
    Cette fonction reunit dans un fichier les colonnes,
    en lisant une ligne dans chaque fichier grace a zip

    Arguments:
        output_file -- la liste des noms de fichiers colonne
        output_file -- la liste des fichiers colonnes
        delimiter   -- le separateur de champ
    ''' 

    col_data = [read_file_lines(col_file, 'utf-8') for col_file in col_files]
    with open(output_file, 'w', encoding='utf-8') as f_out:
        for row in zip(*col_data):
            f_out.write(delimiter.join(cell.strip() for cell in row) + '\n')


def convert_csv_to_excel(input_file, output_file, delimiter="\t"):
    try:
        df = pd.read_csv(input_file, delimiter=delimiter, engine='python')
        df.to_excel(output_file, index=False, engine='xlsxwriter')
        print(f"Données insérées dans {output_file}")
    except Exception as e:
        print(f"Erreur lors de la conversion: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convertir un fichier texte en fichier Excel")
    parser.add_argument("-i", "--input", required=True, help="Chemin vers le fichier texte d'entrée")
    parser.add_argument("-o", "--output", required=True, help="Chemin vers le fichier Excel de sortie")
    parser.add_argument("-d", "--delimiter", default="\t", help="Délimiteur utilisé dans le fichier texte (par défaut: tabulation)")
    
    args = parser.parse_args()
    
    input_text_file_path = args.input
    output_excel_file_path = args.output
    delimiter = args.delimiter
    
    # on cree le repertoire temporaire
    temp_dir = os.path.join('data', 'tmp')
    os.makedirs(temp_dir, exist_ok=True)

    # liste de tuples, ou chaque tuple specifie le debut et la fin des positions des colonnes a extraire
    col_specs = [(0, 63), (63, 81), (81, 111), (111, None)]

    # liste de tuples, ou chaque tuple specifie le nom du fichier colonne temporaire
    col_files = [os.path.join(temp_dir, f"col{i+1}.txt") for i in range(4)]
    
    # on extrait les colonnes et on recupere une liste des colonnes
    columns = extract_columns(input_text_file_path, col_specs)


    # Normaliser la dernière colonne
    # columns[-1] = normalize_dates_in_column(columns[-1])



    # on sauve les colonnes dans des fichiers temporaires
    save_columns_to_files(columns, col_files)
    
    output_csv_file = os.path.join(temp_dir, 'fichier.csv')
    merge_columns_to_csv(output_csv_file, col_files, delimiter)
    
    convert_csv_to_excel(output_csv_file, output_excel_file_path, delimiter)
    
    #clear_temp_directory(temp_dir)
