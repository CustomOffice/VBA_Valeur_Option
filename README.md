# VBA_Valeur_Option
Renvois la valeur associé à la propriété Tag de l'option sélectionnée dans un groupe d'option

##Lien vers le site
http://customoffice.github.io/VBA_Valeur_Option/

## Instruction
- Soit créer un module dans votre projet vba et y copier/coller le code ci-dessous
- Soit télécharger le module (fichier *.bas) et l'inserer dans votre projet vba

##Code
```bash
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!TITRE : Récupération de l'option sélectionné                                                    !!!
'!!!DATE : 17.04.15                                                                                 !!!
'!!!                                                                                                !!!
'!!!DESCRIPTION : Renvois la valeur associé à la propriété Tag de l'option sélectionnée dans un 	!!!
'!!!groupe d'option           																		!!!
'!!!                                                                                                !!!
'!!!REGLES :                                                                                        !!!
'!!!- obj correspond au userform                                                                    !!!
'!!!- groupe est le nom du groupe d'option (propriété "GroupName"), par défaut "" dans VBA          !!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Function valeur_option(ByRef obj, Optional groupe As String = "")
    'Déclaration des variables
    Dim Ctrl As Control
    
    'Paramètre de personnalisation
    valeur_option = ""
    'Boucle sur tous les contrôles
    For Each Ctrl In obj.Controls
        'Vérifie qu'il s'agit d'un OptionButton
        If TypeOf Ctrl Is MSForms.OptionButton Then
            'Véfifie si l'OptionButton fait partie d'un groupe nommé "GR1"
             If Ctrl.GroupName = groupe Then
                'Renvoi le Tag de l'optionButton qui a la valeur True
                If Ctrl.Value = True Then
                    valeur_option = Ctrl.Tag
                    'Sort de la boucle (Il ne peut y a voir qu'une réponse à True)
                    Exit For
                End If
            End If
        End If
    Next
End Function
```
