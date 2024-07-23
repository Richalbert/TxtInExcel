import pandas as pd
import os

def convert_txt_to_excel(input_file, output_file, delimiter="\t"):
    try:
        # Lire le fichier texte en DataFrame
        df = pd.read_csv(input_file, delimiter=delimiter)
        
        # Sauvegarder le DataFrame dans un fichier Excel
        df.to_excel(output_file, index=False)
        
        print(f"Données insérées dans {output_file}")
    except Exception as e:
        print(f"Erreur lors de la conversion: {e}")

if __name__ == "__main__":
    # Chemins par défaut
    input_text_file_path = os.path.join("data", "input", "example.txt")
    output_excel_file_path = os.path.join("data", "output", "output.xlsx")

    # Convertir le fichier texte en fichier Excel
    convert_txt_to_excel(input_text_file_path, output_excel_file_path)

