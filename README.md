# TxtInExcel

TxtInExcel est un outil simple pour convertir des fichiers texte en fichiers Excel.

## Prérequis

Installez les dépendances nécessaires avec :

```sh
pip install -r requirements.txt
```

## Utilisation


```sh
python TxtInExcel.py -i data/input/exemple.txt -o data/output/exemple.xlsx -d ";"
```


## Le contexte / la problematique

J'ai un fichier texte qui a ete genere avec la commande suivante :

```sh
Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Format-Table –AutoSize > scanpc.txt
```

Cette commande permet de connaitre tous les programmes installes.

L'idee est de recuperer ce fichier et de le rendre compatible avec une feuille excel


## l'Algo

En regardant le fichier il est constitue de 4 colonnes. 
- recuperer chaque colonne dans des fichiers
- reunir chaque fichier colonne pour former un fichier cvs
- rendre ce fichier cvs compatible avec une feuile excel

## Remarque

- le fichier source est dans un codage different de utf8, donc il y a une detection du format a l'aide de la bibliotheque *chardet*
- le fichier source semble avoir des lignes de taille fixe, il faut donc a un moment connaitre cette taille pour avoir la fin de la derniere colonne
- la derniere colonne aui correspond a la date doit etre normalise