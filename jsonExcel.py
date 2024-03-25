import json, openpyxl

class JsonVerExcel:
    "Classe pour transformer des données de JSON vers Excel"
    def __init__(self, nomFichierJSON,nomFichierExcel ):
        self._wb = openpyxl.Workbook() #Ouverture du fichier Excel
        self._ws = self._wb.active #Ouverture de la feuille
        self._nomFichierJSON = nomFichierJSON
        self._nomFichierExcel = nomFichierExcel
        self._listeLigne = []

        # Ouverture du fichier
        with open(self._nomFichierJSON,"r",encoding='utf-8') as read_file:
            data = json.load(read_file)
            #Extraction de l'entête
            self._dictionnaireVersListeCles(data[0])
            self._ecritureExcel(self._listeLigne, 1)
            compteur = 0
            for i in data: #Pour chaque item JSON
                self._dictionnaireVersListeDonnees(i, )
                self._ecritureExcel(self._listeLigne, compteur+2)
                compteur += 1

        self._wb.save(self._nomFichierExcel)  # Sauvegarde du fichier Excel
    #Fonction récursive pour les entêtes
    def _dictionnaireVersListeCles(self, dico):
        for cle in dico.keys():
            if isinstance(dico[cle], dict):
                self._dictionnaireVersListeCles(dico[cle])
            else:
                self._listeLigne.append(cle)
                print(cle)

    #Fonction récursive pour les données
    def _dictionnaireVersListeDonnees(self, dico):
        for cle in dico.keys():
            if isinstance(dico[cle], dict) :
                self._dictionnaireVersListeDonnees(dico[cle])
            else:
                self._listeLigne.append(dico[cle])

    def _ecritureExcel(self, liste, noLigne):
        c = 1 #Je ne peux pas utiliser l'index car il y a des répétitions dans les noms
        for i in liste:
            self._ws.cell(row=noLigne, column=c).value = i
            print(i)
            c += 1
        self._listeLigne.clear()


if __name__ == "__main__":
    JVE = JsonVerExcel("Airports.json", "aeroport.xlsx")
