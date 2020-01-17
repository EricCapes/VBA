VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHubPluging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: CHubPluging
' Kind: Module de classe
' Purpose: Classe métier des hubs plugins
' Author: XMXU139
' Date: 16/01/2020
' ------------------------------------------------------

Public Function whereHubPlugin(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As Integer
   whereHubPlugin = 1
End Function

Public Function selectExternalId(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   selectExternalId = r_objArraySheet(lLigne, r_objHeader("clientExternalId"))
End Function

Public Function getHubPlugin(ByRef r_objArraySheet() As Variant, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
   getHubPlugin = CStr(lLigne)
End Function

