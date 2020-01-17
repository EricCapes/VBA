VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetierTestArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: MetierTestArray
' Kind: Module de classe
' Purpose: Classe de test
' Author: XMXU139
' Date: 01/07/2019
' ------------------------------------------------------
Option Explicit
                                          
 
Private m_objStream        As New ManageLog             ' La gestion des traces
Private m_objIndex         As New ManageIndex           ' Gestion des index génériques
Private m_objArray         As New ManageArray

Private m_lLigne           As Long                      ' Numéro de ligne dans la feuille

' ----------------------------------------------------------------
' Procedure Name: beforeReadFile
' Purpose: Exécuter en premier avent de parcourir tout le tableau. Permet d'initialiser le tableau mémoire
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objArray (ManageArray): Référence de l'objet ManageArray
' Return Type: Long
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Public Function beforeReadFile(ByRef r_objArray As ManageArray) As Long
    
    On Error GoTo erreurGenerale
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String

    
    beforeReadFile = 2
    ' beforeExecFile = 2 Arrêt de lecture
    ' beforeExecFile = 0 Continuer à la ligne suivante
    '
    ' Initialise le manager de log et la feuille de travail
    Call m_objStream.InitManageLog
    '
    ' Affecte les entêtes
    Set m_objArray = r_objArray
        
    beforeReadFile = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
    
    Err.Raise lErr, "MetierTest.beforeReadFile(r_objArray)", sDescription
       
    
End Function

' ----------------------------------------------------------------
' Procedure Name: getHeader
' Purpose: Exécuter en deuxieme position. Permet de vérifier le nom des colonnes
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lLigne (Long): Numéro de position en ligne
' Return Type: Integer
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Public Function getHeader(lLigne As Long) As Integer
      
    On Error GoTo erreurGenerale
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
      
    getHeader = 0
    ' getHeader = 1 Passe à la ligne suivante
    ' getHeader = 2 Arrêt de lecture
        
    ' Commence en ligne
    m_lLigne = lLigne + 1
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, "MetierTest.getHeader(" & lLigne & ")", sDescription
      
End Function


' ----------------------------------------------------------------
' Procedure Name: getrow
' Purpose: Exécution en troisème position. Valeur des colonnes de la ligne du tableau
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lLigne (Long): Numéro de ligne
' Return Type: Integer
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Public Function getrow(lLigne As Long) As Integer

    On Error GoTo erreurGenerale
   
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
            
    If (lLigne Mod 500) = 0 Then
        Application.StatusBar = "Ligne numéro : " & lLigne
    End If
      
    getrow = 0
    ' getrow = 1 Passe directement à la ligne suivante
    ' getrow = 2 Arrêt de lecture
        
    
    On Error GoTo 0
    Exit Function

erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description + " : Ligne " & lLigne
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, , sDescription

End Function


' ----------------------------------------------------------------
' Procedure Name: verifRowFomat
' Purpose: Exécutéer en quatrieme position. Permet de vérifier les formats des valeurs
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lLigne (Long): Position en ligne dans le tableau
' Return Type: Integer
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Public Function verifRowFomat(lLigne As Long) As Integer
    
    ' verifRowFomat = 0 Exécute la procédure du SGBDR
    ' verifRowFomat = 1 Passe  à la ligne suivante
    ' verifRowFomat = 2 Arrêt de lecture
    
    verifRowFomat = 1   ' Dans ce cas je ne veux pas vider dans la fonction SGBDR
   
     
End Function
' ----------------------------------------------------------------
' Procedure Name: afterReadFile
' Purpose: Exécutéer en cinquième position des que le tableau est intégralement lu.
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Long
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Public Function afterReadFile() As Long
   
    afterReadFile = 0
    ' afterExecFile = 2 Arrêt de lecture
    
    '
    ' Suppresion des index
    Call m_objIndex.dropAllIndex
    Set m_objIndex = Nothing
    
        
End Function






Public Function selectValue(ByRef r_objArraySheet() As Variant, ByRef r_objIdxHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
    selectValue = CStr(r_objArraySheet(lLigne, r_objIdxHeader("ontology.refContrat")))
End Function

Public Function functionValue(ByRef r_objArraySheet() As Variant, ByRef r_objIdxHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
    functionValue = CStr(r_objArraySheet(lLigne, r_objIdxHeader("ontology.refClient")))
End Function

Public Sub functionValueLigne(ByRef r_objArraySheet() As Variant, ByRef r_objIdxHeader As Scripting.Dictionary, ByVal lLigne As Long, ByRef r_sTab() As String)
    r_sTab(UBound(r_sTab)) = CStr(lLigne)
End Sub

Public Function whereValue(ByRef r_objArraySheet() As Variant, ByRef r_objIdxHeader As Scripting.Dictionary, ByVal lLigne As Long) As Integer
    whereValue = 1
End Function

Public Function selectGroupByTdb(ByRef r_objArraySheet() As Variant, ByRef r_objIdxHeader As Scripting.Dictionary, ByVal lLigne As Long, ByVal sNameIndex) As String
        
    Select Case sNameIndex
    
        Case "idxL001"
            selectGroupByTdb = CStr(r_objArraySheet(lLigne, r_objIdxHeader("ontology.refContrat"))) + "#" + CStr(r_objArraySheet(lLigne, r_objIdxHeader("ontology.typeContenantACollecter")))
        
        Case "idxL002"
            selectGroupByTdb = r_objArraySheet(lLigne, r_objIdxHeader("ontology.refContrat"))
        
        Case "idxL003"
            selectGroupByTdb = r_objArraySheet(lLigne, r_objIdxHeader("ontology.refContrat"))
            
    End Select
        
    
End Function


' ----------------------------------------------------------------
' Procedure Name: sommationTdb
' Purpose: Fonction de sommation
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_ObjArraySheet (Variant()): Tableau variant
' Parameter r_objIdxHeader (Dictionary):
' Parameter lLigne (Long):
' Parameter sEtat (String):
' Parameter sNameIndex (String):
' Return Type: Double
' Author: XMXU139
' Date: 08/07/2019
' ----------------------------------------------------------------
Public Function sommationTdb(ByRef r_objArraySheet() As Variant, ByRef r_objIdxHeader As Scripting.Dictionary, ByVal lLigne As Long, ByVal sEtat As String, ByVal sNameIndex As String) As Double
    
    Select Case sNameIndex
    
        Case "idxL001"
            sommationTdb = 1#
        
        Case "idxL002"
            sommationTdb = 2#
        
        Case "idxL003"
            sommationTdb = 5#
            
    End Select

    
End Function


' ----------------------------------------------------------------
' Procedure Name: whereTdb
' Purpose: Fonction de critères
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objArraySheet (Variant()): Tableau variant d'une feuille
' Parameter r_objIdxHeader (Dictionary): Dictionnaire d'entête des colonnes
' Parameter lLigne (Long): Numéro de ligne
' Parameter sNameIndex (String): Nom de l'index
' Return Type: Integer
' Author: XMXU139
' Date: 08/07/2019
' ----------------------------------------------------------------
Public Function whereTdb(ByRef r_objArraySheet() As Variant, ByRef r_objIdxHeader As Scripting.Dictionary, ByVal lLigne As Long, ByVal sNameIndex As String) As Integer

    whereTdb = 0
       
    Select Case sNameIndex
    
        Case "idxL001"
            whereTdb = 1
        
        Case "idxL002"
            If r_objArraySheet(lLigne, r_objIdxHeader("ontology.typeContenantACollecter")) = "protectedCE30" Then
                whereTdb = 1
            End If
        
        Case "idxL003"
            If r_objArraySheet(lLigne, r_objIdxHeader("ontology.typeFluxACollecter")) = "bak" Then
                whereTdb = 1
            End If
            
    End Select
    
End Function
