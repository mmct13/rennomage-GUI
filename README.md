# Renommer des Fichiers en Utilisant un Fichier Excel - Script Python

Ce script Python permet de renommer plusieurs fichiers dans un dossier en se basant sur les données d'un fichier Excel. Il est utile lorsque tu as une grande quantité de fichiers à renommer en utilisant des correspondances dans une feuille de calcul.

## Fonctionnalités

- Renommage de fichiers en masse : Renomme plusieurs fichiers dans un dossier en une seule opération.
- Utilisation d'un fichier Excel pour les correspondances : Les nouveaux noms des fichiers sont définis dans une colonne d'un fichier Excel.
- Gestion des conflits de nom de fichier : Le script gère automatiquement les cas où un fichier avec le même nom existe déjà dans le dossier.

## Prérequis

Avant d'utiliser ce script, assurez-vous d'avoir installé les dépendances suivantes :

- Python 3.x
- Les bibliothèques Python suivantes :
  - pandas
  - tkinter
  - pillow (PIL)

## Utilisation

1. Assurez-vous que vos fichiers à renommer et votre fichier Excel de correspondance sont prêts.
2. Exécutez le script en utilisant Python en ligne de commande.
3. Suivez les instructions de la fenêtre GUI pour sélectionner le dossier contenant les fichiers à renommer et le fichier Excel de correspondance.
4. Une fois l'opération terminée, vous recevrez une notification indiquant que le processus de renommage est terminé.

## Avertissement

Ce script est conçu pour être utilisé avec précaution. Assurez-vous de sauvegarder vos fichiers avant de les renommer en masse pour éviter toute perte de données accidentelle.

## Remarque

Ce script a été développé à des fins éducatives et peut nécessiter des ajustements en fonction de vos besoins spécifiques. Utilisez-le avec prudence et testez-le d'abord sur un ensemble de fichiers de test pour vous assurer qu'il fonctionne comme prévu.
