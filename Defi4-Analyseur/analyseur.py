#appel du fichier test/ demande à l'utilisateur
nom_fichier = input("Quel fichier voulez_vous lire ?")
chemin_dossier = r'C:\Users\Valentin\Scopi\Support - Documents\apprentissage site\15 défis claude\Defi4-Analyseur\\'
chemin_complet = chemin_dossier + nom_fichier

#ajout de gestion d'erreur
try:
    fichier = open(chemin_complet, 'r', encoding='utf-8')

#lecture du contenu
    contenu = fichier.read()
    print(contenu)

#découpage du texte en mots et mise en liste
    mots = contenu.split()

#compte le nombre total de mots 
    nombre_mots = len(mots)
    print("Nombre total de mots:", nombre_mots)

#crée le dictionnaire
    frequence = {}

#Boucle pour parcourir chaque mots et compter
    for mot in mots:
        mot_propre = mot.strip("!.,?").lower()
        if mot_propre in frequence:
            frequence[mot_propre] = frequence[mot_propre] + 1
        else:
            frequence[mot_propre] = 1


    print(frequence)

#Triage du dictionnaire .items donne les paires, lambda trie selon la valeur et reverse=true trie dans l'ordre décroissant
    mots_tries = sorted(frequence.items(), key=lambda x: x[1], reverse=True)
    print("\nMots les plus fréquents :")
    print(mots_tries)

#Création et ouverture du fichier rapport.txt
    fichier_rapport = open(r'C:\Users\Valentin\Scopi\Support - Documents\apprentissage site\15 défis claude\Defi4-Analyseur\rapport.txt', 'w' , encoding='utf-8')

#écriture dans le fichier rapport.txt
    fichier_rapport.write("=== RAPPORT D'ANALYSE ===\n")
    fichier_rapport.write(f"Fichier analysé : {nom_fichier}\n")
    fichier_rapport.write(f"Nombre total de mots: {nombre_mots}\n")
    fichier_rapport.write("\n")
    fichier_rapport.write("=== MOTS LES PLUS FRÉQUENTS ===\n")
    fichier_rapport.write(f"Mots les plus fréquents: ")

#Boucle pour afficher les mots les plus fréquents. Index sépare le tuple et enumerate numérote la liste, on commence à 1 pas à 0
    for index, (mot, frequence) in enumerate(mots_tries, 1):
        fichier_rapport.write(f"{index}. {mot} : {frequence} occurrences\n")


#ferme les fichiers
    fichier_rapport.close()

    fichier.close()

except FileNotFoundError :
    print("Erreur: Le fichier n'existe pas dans ce dossier !")
