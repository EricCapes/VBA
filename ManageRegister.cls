VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: ManageRegister
' Kind: Module de classe
' Purpose: Sauvegarde et lecture dans la base de registre
' Author: Eric VENTALON
' Date: 29/08/2019
' ------------------------------------------------------

'
' l 'instruction pour enregistrer quelque chose dans la base de registre est SaveSettings. Ce que vous sauvegardez avec SaveSettings s'installera
' automatiquement dans (Avec Windows 2000) :
'       HKEY_USERS \ S-4-6-23-9092171932-311282134-1221991238-2009 \ Software \ VB and VBA Program Settings
'
' Le chiffre en bleu est l'identification (fictive) de mon nom d'utilisateur. Avant d 'aller plus loin, et à titre d'exercice, essayez de trouver dans la base de
' registre (Démarrer/Exécuter : regedit), ce fameux endroit VB and VBA Program Settings. Dans cette clé (VB and VBA Program Settings) vont donc se trouver les clés
' et sous-clés de votre application.
'


Option Explicit

Public Sub saveRegister(ByVal sSection As String, ByVal sValue As String)
   
    SaveSetting appname:="certificat", section:="config", Key:=sSection, setting:=sValue
End Sub



Public Function getRegister(ByVal sSection As String) As String

    getRegister = GetSetting(appname:="certificat", section:="config", Key:=sSection)
    
End Function


