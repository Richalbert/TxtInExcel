import pandas as pd
import os
import argparse
import shutil



def convert_txt_to_excel(input_file, output_file, delimiter="\t"):
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

        # print(data)

        # Créer un DataFrame à partir des données
        df = pd.DataFrame(data[1:], columns=data[0])

        # print(df)

        ############
        # Ecriture #
        ############

        # Sauvegarder le DataFrame dans un fichier Excel
        df.to_excel(output_file, index=False, engine='xlxs')
        
        print(f"Données insérées dans {output_file}")
    except Exception as e:
        print(f"Erreur lors de la conversion: {e}")
        

if __name__ == "__main__":
    # # Chemins par défaut
    # input_text_file_path = os.path.join("data", "input", "example.txt")
    # output_excel_file_path = os.path.join("data", "output", "output.xlsx")

    parser = argparse.ArgumentParser(description="Convertir un fichier texte en fichier Excel")
    parser.add_argument("-i", "--input", required=True, help="Chemin vers le fichier texte d'entrée")
    parser.add_argument("-o", "--output", required=True, help="Chemin vers le fichier Excel de sortie")
    parser.add_argument("-d", "--delimiter", default="\t", help="Délimiteur utilisé dans le fichier texte (par défaut: tabulation)")

    args = parser.parse_args()

    input_text_file_path = args.input
    output_excel_file_path = args.output
    delimiter = args.delimiter


    print(input_text_file_path)
    print(output_excel_file_path)
    print(delimiter)

    # Convertir le fichier texte en fichier Excel
    convert_txt_to_excel(input_text_file_path, output_excel_file_path, delimiter)




