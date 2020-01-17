VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CServiceMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'
' Ajouter les 3 fonctions ci-dessous dans l'objet métiers s'ils n'existent pas
'
Public Function whereSm(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As Integer
   whereSm = 1
End Function

Public Function selectSm(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   selectSm = r_objArraySheet(lLigne, r_objHeader("ontology.order.orderId"))
End Function

Public Function getServiceMark(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   getServiceMark = r_objArraySheet(lLigne, r_objHeader("ontology.service.serviceMark"))
End Function


