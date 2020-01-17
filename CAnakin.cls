VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAnakin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'
' Ajouter les 3 fonctions ci-dessous dans l'objet métiers s'ils n'existent pas
'
Public Function whereMissionUuid(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As Integer
   whereMissionUuid = 1
End Function

Public Function selectMissionUuid(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   ' C'est l'entrée du dictionnaire. Exemple Compte('ANS46654654645')
   ' Modifier le code ici
   selectMissionUuid = r_objArraySheet(lLigne, r_objHeader("Mission_UUID"))
End Function

Public Function getLineMissionUuid(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   ' C'est le contenue de l'entrée du dictionnaire . Exemple Compte('ANS46654654645') -> 525698.45
   ' Modifier le code ici
   getLineMissionUuid = CStr(lLigne)
End Function

