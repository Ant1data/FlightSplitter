# FlightSplitter - Rotation vols Air France

Ce projet permet de générer automatiquement des fichiers Excel de rotations de vols Air France à partir d’un fichier source.  
Interface graphique développée en Python avec Tkinter, barre de progression et animation.

## Comment utiliser

1. Double-cliquez sur `lancer_rotation.bat` pour lancer l’interface (aucune console noire).  
2. Choisissez le fichier Excel source.  
3. Choisissez le dossier de sortie.  
4. Cliquez sur Lancer le traitement.  
5. Les fichiers Excel seront générés et le dossier de sortie s’ouvrira automatiquement.

## Dépendances

- Python 3.x (Tkinter inclus)  
- pandas  
- openpyxl  

 Pour créer l’environnement Python requis   
```bash
conda create -n mini-flightsplitter python=3.x pandas openpyxl

