VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetierTestFichier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: MetierTestFichier
' Kind: Module de classe
' Purpose: Classe de test
' Author: XMXU139
' Date: 01/07/2019
' ------------------------------------------------------

Option Explicit
                                          
 
Private m_objStream        As New ManageLog             ' La gestion des traces
Private m_objOnglet        As New ManageOnglet          ' La gestion des feuilles
Private m_objIndex         As New ManageIndex           ' Gestion des index génériques
Private m_objStat          As New manageStat            ' gestion des statistiques et de leur fichier

Private m_sSheetName       As String                    ' Le nom de la feuille

Private m_objHeaderFile    As New Scripting.Dictionary  ' Tableau des entêtes

' ----------------------------------------------------------------
' Procedure Name: beforeReadFile
' Purpose: Exécuter en premier avent de parcourir tout le fichier.
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Long
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Function beforeReadFile() As Long

    On Error GoTo erreurGenerale
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String

    
    beforeReadFile = 2
    ' beforeReadFile = 2 Arrêt de lecture
    ' beforeReadFile = 0 Poursuivre la lecture
    
    '
    ' Initialise le manager de log et la feuille de travail
    Call m_objStream.InitManageLog
    '
    ' Ouverture des statistiques
    If m_objStat.openStat("P_INPUT_VBAUNIT_ERR") = False Then
        Err.Raise 700, , "La paramètre P_INPUT_VBAUNIT_ERR n'est pas défini."
    End If
    '
    ' Mise en place des stats
    Call m_objStat.addStat(0, "...A la demande   : ")
    Call m_objStat.addStat(1, "...Désarchivage   : ")
    Call m_objStat.addStat(2, "...Abonnement     : ")
    Call m_objStat.addStat(3, "...Abonnement box : ")
    Call m_objStat.addStat(4, "...Divers         : ")
    
    beforeReadFile = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
    
    Call Err.Raise(lErr, "MetierTestFichier.beforeReadFile()", sDescription)
           
End Function
' ----------------------------------------------------------------
' Procedure Name : getHeader
' Purpose : Lecture des noms  des colonnes. Exécution en position n°2
'           et exécuté qu'une fois. Vérification des entêtes
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lLigne (Long): Numéro de ligne dans le fichier
' Parameter r_objColonne (String()): Nom des colonnes
' Return Type: Integer
' Author: Eric VENTALON
' Version: 01.00
' Date: 20/05/2019
' ----------------------------------------------------------------
Public Function getHeader(lLigne As Long, ByRef r_objColonne() As String) As Integer
    
    On Error GoTo erreurGenerale
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
  
    Dim iCompteur             As Integer
      
    getHeader = 0
    ' getHeader = 1 Passe à la ligne suivante
    ' getHeader = 2 Arrêt de lecture
     
    
    ' Mise en place des colonnes du fichier
    For iCompteur = 0 To UBound(r_objColonne)
        Call m_objHeaderFile.Add(r_objColonne(iCompteur), iCompteur + 1)
    Next iCompteur
                
            
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
    
    Call Err.Raise(lErr, "MetierTestFichier.getHeader()", sDescription)
      
End Function
' ----------------------------------------------------------------
' Procedure Name : getrow
' Purpose : Valeur des colonnes de la ligne du fichier. Exécution en position n°3
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lLigne (Long): Numéro de ligne
' Parameter r_objColonne (String()): Valeur pour chaque colonne
' Return Type: Integer
' Author: Eric VENTALON
' Version: 01.00
' Date: 20/05/2019
' ----------------------------------------------------------------
Public Function getrow(lLigne As Long, ByRef r_objColonne() As String) As Integer

    On Error GoTo erreurGenerale
    
  
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
            
    If (lLigne Mod 500) = 0 Then
        Application.StatusBar = "Lecture fichier TU : " & lLigne
    End If
      
    getrow = 0
    ' getrow = 1 Passe directement à la ligne suivante
    ' getrow = 2 Arrêt de lecture
    
    Select Case r_objColonne(5)
    Case "A la demande"
        Call m_objStat.sumStat(0)
    Case "Désarchivage"
        Call m_objStat.sumStat(1)
    Case "Abonnement"
        Call m_objStat.sumStat(2)
    Case "Abonnement box"
        Call m_objStat.sumStat(3)
    Case Else
        Call m_objStat.sumStat(4)
    End Select
        
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description & " : Ligne = " & lLigne
    sSource = Err.Source
        
    On Error GoTo 0
    
    Call Err.Raise(lErr, "MetierTestFichier.getrow()", sDescription)
    
End Function
' ----------------------------------------------------------------
' Procedure Name : verifRowFomat
' Purpose : Vérification du format des valeurs. Exécution en position n°4
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lLigne (Long): Numéro de ligne du fichier
' Parameter r_objColonne (String()): Valeur des colonnes du fichier
' Return Type: Integer
' Author: Eric VENTALON
' Version: 01.00
' Date: 20/05/2019
' ----------------------------------------------------------------
Public Function verifRowFomat(lLigne As Long, ByRef r_objColonne() As String) As Integer
     
    On Error GoTo erreurGenerale
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
      
    verifRowFomat = 1 ' Dans ce cas je ne veux pas vider dans la fonction SGBDR
    ' verifRowFomat =1 Passe  à la ligne suivante
    ' verifRowFomat = 2 Arrêt de lecture
    
    
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
    
    Err.Raise lErr, "MetierTestFichier.verifRowFomat()", sDescription
    
End Function
' ----------------------------------------------------------------
' Procedure Name : afterReadFile. Exécuter en position n°5
' Purpose : Excuter en fin de lecture du fichier
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Long
' Author: Eric VENTALON
' Date: 22/05/2019
' ----------------------------------------------------------------
Function afterReadFile() As Long
   
    afterReadFile = 0
    ' afterExecFile = 2 Arrêt de lecture
    '
    
    ' Affiche les statistiques
    Call m_objStream.writeInfo("...STATISTIQUES VBA UNIT")
    Call m_objStream.displayMessage
    Call m_objStat.displayStat(0, 4)
    '
    ' Suppresion des index
    Call m_objHeaderFile.RemoveAll
    Call m_objIndex.dropAllIndex
        
    Set m_objHeaderFile = Nothing
    Set m_objIndex = Nothing
    
        
End Function



' ----------------------------------------------------------------
' Procedure Name : selectValue
' Purpose : Extraire la clé primaire pour le dictionnaire
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sSheetName (String): Nom de la feuille
' Parameter r_objHeader (Dictionary): Index des colonnes
' Parameter lLigne (Long): Numéro de ligne
' Return Type: String
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Public Function selectValue(ByVal sSheetName As String, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String

    ' Extraire la valeur de l'article + Statut  commande
    selectValue = Trim(Sheets(sSheetName).Cells(lLigne, r_objHeader("id commande la poste")).Text) + Trim(Sheets(sSheetName).Cells(lLigne, r_objHeader("Statut  commande")).Text)
         
End Function

' ----------------------------------------------------------------
' Procedure Name: functionValue
' Purpose: Fonction de projection servant la PK
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sSheetName (String): Nom de la feuille
' Parameter r_objHeader (Dictionary): Nom du dictionnaire
' Parameter lLigne (Long): Numéro de ligne
' Return Type: String
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Public Function functionValue(ByVal sSheetName As String, ByRef r_objHeader As Scripting.Dictionary, ByVal lLigne As Long) As String
  
    functionValue = CStr(lLigne)
      
End Function

' ----------------------------------------------------------------
' Procedure Name: whereValue
' Purpose: Fonction des critères pour lindexation
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sSheetName (String): Nom de la feuille
' Parameter r_objHeader (Dictionary): Nom du dictionnaire
' Parameter lLigne (Long): Numéro de ligne
' Return Type: Integer
' Author: Eric VENTALON
' Date: 17/07/2019
' ----------------------------------------------------------------
Public Function whereValue( _
    ByVal sSheetName As String, _
    ByRef r_objHeader As Scripting.Dictionary, _
    ByVal lLigne As Long _
    ) As Integer
    
    whereValue = 1
    
End Function




' ----------------------------------------------------------------
' Procedure Name: getValueBeforeSavePresto
' Purpose: Permet de formatter la ligne avant son enregistrement
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_vPrimaryKey (Variant): la clé primaire du dico
' Parameter r_vDico (Variant): le contenu du dico déterminé par la PK
' Return Type: String
' Author: Eric VENTALON
' Date: 06/06/2019
' ----------------------------------------------------------------
Public Function getValueBeforeSavePresto(ByRef r_vPrimaryKey As Variant, ByRef r_vDico As Variant) As String
    
    Dim sPrimaryKey As String
    Dim sTuple As String
    
    sPrimaryKey = CStr(r_vPrimaryKey)
    sTuple = CStr(r_vDico)
    
    ' TODO Voir JIRA - MIGR-INT-05- Préparation fichier pour validation automatique des faisabilités
    getValueBeforeSavePresto = sTuple
        
End Function


