import pandas as pd
import os
import argparse
import chardet
import shutil



def clear_temp_directory(directory_path):

    print(directory_path)
    
    try:
        # Vérifie si le répertoire existe
        if os.path.exists(directory_path):

            # Parcourt tous les fichiers et dossiers dans le répertoire
            for filename in os.listdir(directory_path):

                file_path = os.path.join(directory_path, filename)
                try:
                    # Supprime les fichiers
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)

                    # Supprime les répertoires et leur contenu
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)

                except Exception as e:
                    print(f'Erreur lors de la suppression de {file_path}. Raison: {e}')

        else:
            print(f'Le répertoire {directory_path} n\'existe pas.')

    except Exception as e:
        print(f'Erreur lors de la vérification du répertoire {directory_path}. Raison: {e}')




def fin_ligne(input_file):
    try:
        # Detecter l'encodage du fichier
        with open(input_file, 'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']

        # Ouvrir le fichier d'entrée en lecture avec l'encodage detecte
        with open(input_file, 'r', encoding=encoding) as f_in:

            # Lire la première ligne pour déterminer la longueur
            premiere_ligne = f_in.readline()
            longueur_ligne = len(premiere_ligne.rstrip('\n'))

            # Revenir au début du fichier
            f_in.seek(0)

    except FileNotFoundError:   
        print(f"Erreur : Le fichier '{input_file}' n'a pas été trouvé.")

    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

    return longueur_ligne

    

def extraire_colonne(input_file, output_file, debut, fin):
    """
    Lit un fichier texte ligne par ligne et conserve un ensemble de caracteres contigues.
    
    :param input_file: Nom du fichier texte en entrée
    :param output_file: Nom du fichier texte en sortie
    :param debut: premiere position du caractere a conserver
    :param fin: derniere position du caractere a conserver
    """
    try:
        # Detecter l'encodage du fichier
        with open(input_file, 'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']

        # Ouvrir le fichier d'entrée en lecture avec l'encodage detecte
        with open(input_file, 'r', encoding=encoding) as f_in:
        
            # Ouvrir le fichier de sortie en ecriture
            with open(output_file, 'w', encoding='utf-8') as f_out:

                # Lire chaque ligne du fichier d'entrée
                for ligne in f_in:

                    # Conserver les caracteres entre debut et fin
                    caracteres = ligne[debut:fin]
                   
                    # Écrire les caracteres dans le fichier de sortie
                    f_out.write(caracteres + '\n')

        print(f"Les caractères ont été sauvegardés dans '{output_file}'.")

    except FileNotFoundError:
        print(f"Erreur : Le fichier '{input_file}' n'a pas été trouvé.")

    except Exception as e:
        print(f"Une erreur s'est produite : {e}")



def netoyer_colonne(input_fichier, output_fichier):
    """
    Lit un fichier texte ligne par ligne, 
    supprime tous les espace avant le retour chariot, 
    et ecrit les lignes modifiées dans un nouveau fichier.
    
    :param input_fichier: Nom du fichier texte en entree
    :param output_fichier: Nom du fichier texte en sortie
    """
    try:
        # Ouvrir le fichier d'entree en lecture
        with open(input_fichier, 'r') as f_in:

            # Ouvrir le fichier de sortie en écriture
            with open(output_fichier, 'w') as f_out:

                # Lire chaque ligne du fichier d'entrée
                for ligne in f_in:

                    # Supprimer les espaces de fin 
                    ligne_strip = ligne.rstrip()

                    # Ajouter un retour chariot pour marquer la fin de ligne
                    nouvelle_ligne = ligne_strip + '\n'

                    # Écrire la ligne modifiee dans le fichier de sortie
                    f_out.write(nouvelle_ligne)        
                    
        print(f"Les lignes modifiées ont été sauvegardées dans '{output_fichier}'.")

    except FileNotFoundError:
        print(f"Erreur : Le fichier '{input_fichier}' n'a pas été trouvé.")

    except Exception as e:
        print(f"Une erreur s'est produite dans netoyer_colonne(): {e}")




def reconstruire_csv(fichier_col1, fichier_col2, fichier_col3, fichier_col4, fichier_output, delimiteur=','):
    """
    Reconstruit un fichier CSV à partir de quatre fichiers texte, chacun représentant une colonne.

    :param fichier_col1: Nom du fichier texte pour la colonne 1
    :param fichier_col2: Nom du fichier texte pour la colonne 2
    :param fichier_col3: Nom du fichier texte pour la colonne 3
    :param fichier_col4: Nom du fichier texte pour la colonne 4
    :param fichier_output: Nom du fichier CSV en sortie
    :param delimiteur: Délimiteur à utiliser dans le fichier CSV (par défaut ',')
    """
    try:

        # Ouvrir les fichiers d'entrée en lecture
        with open(fichier_col1, 'r') as f1, \
             open(fichier_col2, 'r') as f2, \
             open(fichier_col3, 'r') as f3, \
             open(fichier_col4, 'r') as f4:

            # Ouvrir le fichier de sortie en écriture
            with open(fichier_output, 'w') as fout:

                # Lire les fichiers ligne par ligne simultanément
                for ligne1, ligne2, ligne3, ligne4 in zip(f1, f2, f3, f4):

                    # Supprimer les espaces et les retours à la ligne des lignes lues
                    col1 = ligne1.strip()
                    col2 = ligne2.strip()
                    col3 = ligne3.strip()
                    col4 = ligne4.strip()

                    # Combiner les colonnes avec le délimiteur choisi
                    ligne_csv = delimiteur.join([col1, col2, col3, col4])

                    # Écrire la ligne combinée dans le fichier de sortie
                    fout.write(ligne_csv + '\n')

        print(f"Le fichier CSV '{fichier_output}' a été reconstruit avec succès.")

    except FileNotFoundError as e:
        print(f"Erreur : {e}")

    except Exception as e:
        print(f"Une erreur s'est produite : {e}")



def convert_cvs_to_excel(input_file, output_file, delimiter="\t"):
    try:
        ###########
        # Lecture #
        ###########

        # # Lire le fichier texte en DataFrame avec le delimiteur specifie
        # df = pd.read_csv(input_file, delimiter=delimiter, engine='python')

        # Lire le fichier texte ligne par ligne avec readlines()
        with open(input_file, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        
        # Diviser chaque ligne en colonnes en utilisant le délimiteur
        data = [line.strip().split(delimiter) for line in lines]



        #print(data)




        # Créer un DataFrame à partir des données
        df = pd.DataFrame(data[1:], columns=data[0])

        print(df)






        ############
        # Ecriture #
        ############

        # Sauvegarder le DataFrame dans un fichier Excel
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


    # on extrait les 4 colonnes du fichier initial
    input_fichier_text = input_text_file_path

    col1 = os.path.join("data", "tmp", "col1.txt")
    col2 = os.path.join("data", "tmp", 'col2.txt')
    col3 = os.path.join("data", "tmp", 'col3.txt')
    col4 = os.path.join("data", "tmp", 'col4.txt')

    
    eol = fin_ligne(input_fichier_text)
    

    extraire_colonne(input_fichier_text, col1, 0, 63)
    extraire_colonne(input_fichier_text, col2, 63, 81)
    extraire_colonne(input_fichier_text, col3, 81, 111)
    extraire_colonne(input_fichier_text, col4, 111, eol)



    # on netoye les 4 colonnes des espaces a la fin de chaque ligne
    col1clean = os.path.join("data", "tmp",'col1clean.txt')
    col2clean = os.path.join("data", "tmp",'col2clean.txt')
    col3clean = os.path.join("data", "tmp",'col3clean.txt')
    col4clean = os.path.join("data", "tmp",'col4clean.txt')

    netoyer_colonne(col1, col1clean)
    netoyer_colonne(col2, col2clean)
    netoyer_colonne(col3, col3clean)
    netoyer_colonne(col4, col4clean)



    # on reunit les colonnes pour en faire un fichier cvs
    output_fichier_cvs = os.path.join("data", "tmp",'fichier.cvs')
    delimiteur = delimiter
    reconstruire_csv(col1clean, col2clean, col3clean, col4clean, output_fichier_cvs, delimiteur)


    # on transforme le fichier cvs en fichier excel
    input_fichier_cvs = output_fichier_cvs
    output_fichier = output_excel_file_path
    convert_cvs_to_excel(input_fichier_cvs, output_fichier, delimiteur)


    # on clean le reptoire temporaire
    clear_temp_directory('data/tmp')

