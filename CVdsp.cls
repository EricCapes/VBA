VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVdsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'
' Ajouter les 3 fonctions ci-dessous dans l'objet métiers s'ils n'existent pas
'
Public Function whereVdsp(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As Integer
   whereVdsp = 1
End Function

Public Function selectMission(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   selectMission = r_objArraySheet(lLigne, r_objHeader("Missions"))
End Function

Public Function getLineMission(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   getLineMission = CStr(lLigne)
End Function

