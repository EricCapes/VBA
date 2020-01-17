VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' ----------------------------------------------------------------
' Name: ManageIndex
' Help: https://app.gitbook.com/@eric-ventalon/s/vba-framework/
' Kind: Module de classe
' Purpose: Classe générique pour réaliser des group by avec count, sum, etc ...
'          Création d'index quelconques
'          Numéro d'erreur de 600 à 650
' Author: Eric VENTALON
' Date: 08/05/2019
' ----------------------------------------------------------------
Option Explicit

Private m_objGeneral      As New Scripting.Dictionary
Private m_objCustomIndex  As Object
' ----------------------------------------------------------------
' Procedure Name: createIndex
' Purpose: Creation d'un dictionnaire
' Procedure Kind: Function
' Procedure Access: Private
' Parameter sNomIndex (String): Nom de l'index
' Return Type: Boolean
' Author: Eric VENTALON
' Date: 16/07/2019
' ----------------------------------------------------------------
Private Function createIndex(ByVal sNomIndex As String) As Boolean
    
    createIndex = False
    
    Call deleteIndex(sNomIndex)
    Call m_objGeneral.Add(sNomIndex, New Scripting.Dictionary)
    
    createIndex = True
    
End Function


' ----------------------------------------------------------------
' Procedure Name: deleteIndex
' Purpose: Supprime un index par critère de son nom
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sNomIndex (String): Nom de l'index
' Return Type: Boolean
' Author: Eric VENTALON
' Date: 08/05/2019
' ----------------------------------------------------------------
Public Function deleteIndex(ByVal sNomIndex As String) As Boolean

    Dim objIndexLocal As Scripting.Dictionary
    
    deleteIndex = False
    
    If m_objGeneral.Exists(sNomIndex) Then
        Set objIndexLocal = m_objGeneral(sNomIndex)
        Call objIndexLocal.RemoveAll
        Call m_objGeneral.Remove(sNomIndex)
        deleteIndex = True
    End If
        
End Function

' ----------------------------------------------------------------
' Procedure Name: getIndex
' Purpose: Retourne un objet index par critere de son nom
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sNomIndex (String): Nom de l'index
' Return Type: Dictionary
' Author: Eric VENTALON
' Date: 08/05/2019
' ----------------------------------------------------------------
Public Function getIndex(sNomIndex As String) As Scripting.Dictionary
    If m_objGeneral.Exists(sNomIndex) Then
        Set getIndex = m_objGeneral(sNomIndex)
    Else
        Set getIndex = Nothing
    End If
End Function

' ----------------------------------------------------------------
' Procedure Name: dropAllIndex
' Purpose: Supprime tous les index
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Eric VENTALON
' Date: 08/05/2019
' ----------------------------------------------------------------
Public Sub dropAllIndex()

    Dim vKey As Variant
   
    For Each vKey In m_objGeneral.Keys()
        Call m_objGeneral(vKey).RemoveAll
    Next
    m_objGeneral.RemoveAll
    
End Sub


' ----------------------------------------------------------------
' Procedure Name : createIndexUniqueValues
' Purpose : Création d'un index avec des valeurs multiples
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objWorkBook (Workbook): La workbook cible
' Parameter sSheetName (String): La feuille cible
' Parameter sColNameCount (String): La colonne coble pour compter les lignes
' Parameter bClearIndex (Boolean): Remise à zéro de l index ou pas
' Parameter sIndexName (String): Nom de l'index
' Parameter sFunctionSelect (String): nom de la fonction pour le select
' Parameter sFunctionValues (String): Nom de la fonction pour le calcul
' Parameter sFunctionWhere (String): Nom de la fonction pour le where
' Return Type: Integer
' Author: Eric VENTALON
' Date: 26/05/2019
' ----------------------------------------------------------------
Public Function createIndexUniqueValues( _
    ByRef r_objWorkBook As Workbook, _
    ByVal sSheetName As String, _
    ByVal sColNameCount As String, _
    ByVal bClearIndex As Boolean, _
    ByVal sIndexName As String, _
    ByVal sFunctionSelect As String, _
    ByVal sFunctionValues As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
    
    On Error GoTo erreurGenerale
    
    Dim objTrace              As New ManageLog
    Dim objOnglet             As New ManageOnglet
    
    Dim objIdxHeader          As Scripting.Dictionary
    Dim objIndexLocal         As Scripting.Dictionary
    
    Dim sValeurPkFinale       As String
    Dim lMaxLine              As Long
    Dim lLigne                As Long
    Dim lStat(4)              As Long
    Dim sPosValeur            As String
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim iReturnVal            As Integer
    '
    ' 01 - Par défaut retour pésimiste
    createIndexUniqueValues = -1
    '
    ' 02 - Initialisation des lStats
    lStat(0) = 0: lStat(1) = 0: lStat(2) = 0: lStat(3) = 0
    '
    ' 03 - Initialisation du manager
    Call objTrace.InitManageLog
    '
    ' 04 - Mise en place de l'espace de travail et récupère les colonnes
    If r_objWorkBook Is Nothing Then
        r_objWorkBook = ThisWorkbook
    End If
    Call objOnglet.getColumnWithIndexPosition(r_objWorkBook, sSheetName, objIdxHeader)
    '
    ' 06 - Signaler le debut de l'indexation
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("DEBUT : Creation index " + sIndexName + " dans " + sSheetName + ", Valeurs : " + sFunctionValues + ", critères: " + sFunctionWhere)
    '
    ' 07 - Remise à zéro de l'index ?
    If bClearIndex = True Then Call createIndex(sIndexName)
    Set objIndexLocal = getIndex(sIndexName)
    '
    ' 08 - Si pas de classe métier
    If IsNull(m_objCustomIndex) Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    If m_objCustomIndex Is Nothing Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    '
    ' 08 - Parcourir toutes les lLignes du fichier
    r_objWorkBook.Sheets(sSheetName).Range(sColNameCount + "1").Select
    lMaxLine = Application.WorksheetFunction.CountA(Range(sColNameCount + ":" + sColNameCount))
    For lLigne = 2 To lMaxLine
        '
        ' 09 - Fonction critère ?
        If sFunctionWhere <> "" Then
            iReturnVal = CallByName(m_objCustomIndex, sFunctionWhere, VbMethod, sSheetName, objIdxHeader, lLigne)
        Else
            iReturnVal = 1
        End If
        '
        ' 10 - creation de l'index
        If iReturnVal = 1 Then
            '
            ' 11 - Extaction de la valeur de la PK composite
            sValeurPkFinale = CallByName(m_objCustomIndex, sFunctionSelect, VbMethod, sSheetName, objIdxHeader, lLigne)
            '
            ' 12 - Une fois la clé pk composite formée et existante ....
            If sValeurPkFinale <> "" Then
                '
                ' 13 - Est-ce que l'order ID existe ?
                If Not objIndexLocal.Exists(sValeurPkFinale) Then
                    sPosValeur = CallByName(m_objCustomIndex, sFunctionValues, VbMethod, sSheetName, objIdxHeader, lLigne)
                    Call objIndexLocal.Add(sValeurPkFinale, sPosValeur)
                    lStat(0) = lStat(0) + 1
                Else
                    lStat(1) = lStat(1) + 1
                End If
            Else
                lStat(2) = lStat(2) + 1
            End If
        End If
        
        If (lLigne Mod 500) = 0 Then
            Application.StatusBar = "Index: " + sIndexName + ", From: " + sSheetName + ", Fct: " + sFunctionValues + ", Where: " + sFunctionWhere + ", Progression: " & lLigne & " / " & lMaxLine & " : " & Format(lLigne / lMaxLine, "Percent")
        End If
    Next
    
    
    Call objTrace.writeInfo("...STATS - FROM: " + sSheetName + ", INDEX: " + sIndexName + ", FCT:" + sFunctionValues + ", WHERE: " + sFunctionWhere)
    Call objTrace.writeInfo("...AFFECTE DANS l'INDEX..: " & lStat(0))
    Call objTrace.writeInfo("...DOUBLES...............: " & lStat(1))
    Call objTrace.writeInfo("...PK VIDE...............: " & lStat(2))
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("FIN : Creation de l'index.")
    Call objTrace.displayMessage
    
    createIndexUniqueValues = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    Call objTrace.writeError("...Classe      : ManageIndex")
    Call objTrace.writeError("...Fonction    : createIndexUniqueValues")
    Call objTrace.writeError(".....Numéro      : " & lErr)
    Call objTrace.writeError(".....Explication : " & sDescription)
    Call objTrace.writeError(".....Source      : " & sSource)
    Call objTrace.writeError(".....lLigne       : " & lLigne)
        
    On Error GoTo 0
   
    Err.Raise lErr, , sDescription
End Function

' ----------------------------------------------------------------
' Procedure Name: createIndexMultiValue
' Purpose: Permet de realiser un dictionnaire en combinant des fonctions avec des critères
' Procedure Kind: Function
' Procedure Access: Public
'
' Parameter sSheetName (String):
' Parameter bClearIndex (Boolean):
' Parameter sIndexName (String):
' Parameter sFunctionGetValue (String):
' Parameter argFunctionGetPk (String):
' Parameter argFunctionCritereName (String):
'
' Return Type: Integer
' Author: Eric VENTALON
' Date: 08/05/2019
' ----------------------------------------------------------------
Public Function createIndexMultiValue( _
    ByRef r_objWorkBook As Workbook, _
    ByVal sSheetName As String, _
    ByVal sColNameCount As String, _
    ByVal bClearIndex As Boolean, _
    ByVal sIndexName As String, _
    ByVal sFunctionSelect As String, _
    ByVal sFunctionValues As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
    
    On Error GoTo erreurGenerale
    
    Dim objTrace              As New ManageLog
    Dim objOnglet             As New ManageOnglet
    
    Dim objIdxHeader          As Scripting.Dictionary
    Dim objIndexLocal         As Scripting.Dictionary
    
    Dim sValeurPkFinale       As String
    Dim lMaxLine              As Long
    Dim lLigne                As Long
    Dim lStat(4)              As Long
    Dim sPosValeur()          As String
    
    Dim lErr                 As Long
    Dim sDescription         As String
    Dim sSource              As String
    
    Dim iReturnVal            As Integer
    
    '
    ' 01 - Par défaut retour pésimiste
    createIndexMultiValue = -1
    '
    ' 02 - Initialisation des lStats
    lStat(0) = 0: lStat(1) = 0: lStat(2) = 0: lStat(3) = 0
    '
    ' 03 - Initialisation du manager
    Call objTrace.InitManageLog
    '
    ' 04 - Mise en place de l'espace de travail et récupère les colonnes
    If r_objWorkBook Is Nothing Then
        r_objWorkBook = ThisWorkbook
    End If
    Call objOnglet.getColumnWithIndexPosition(r_objWorkBook, sSheetName, objIdxHeader)
    '
    ' 05 - Signaler le debut de l'indexation
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("DEBUT : Creation index " + sIndexName + " dans " + sSheetName + ", Valeurs : " + sFunctionValues + ", critères: " + sFunctionWhere)
    '
    ' 06 - Remise à zéro de l'index ?
    If bClearIndex = True Then Call createIndex(sIndexName)
    Set objIndexLocal = getIndex(sIndexName)
    '
    ' 07 - Si pas de classe métier
    If IsNull(m_objCustomIndex) Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    If m_objCustomIndex Is Nothing Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    '
    ' 08 - Parcourir toutes les lLignes du fichier
    r_objWorkBook.Sheets(sSheetName).Range(sColNameCount + "1").Select
    lMaxLine = Application.WorksheetFunction.CountA(Range(sColNameCount + ":" + sColNameCount))
    For lLigne = 2 To lMaxLine
        '
        ' 09 - Fonction critère ?
        If sFunctionWhere <> "" Then
            iReturnVal = CallByName(m_objCustomIndex, sFunctionWhere, VbMethod, sSheetName, objIdxHeader, lLigne)
        Else
            iReturnVal = 1
        End If
        '
        ' 10 - creation de l'index
        If iReturnVal = 1 Then
            '
            ' 11 - Extaction de la valeur de la PK composite
            sValeurPkFinale = CallByName(m_objCustomIndex, sFunctionSelect, VbMethod, sSheetName, objIdxHeader, lLigne)
            '
            ' 12 - Une fois la clé pk composite formée et existante ....
            If sValeurPkFinale <> "" Then
                '
                ' Est-ce que l'order ID existe ?
                If Not objIndexLocal.Exists(sValeurPkFinale) Then
                    ReDim sPosValeur(0) As String
                    Call CallByName(m_objCustomIndex, sFunctionValues, VbMethod, sSheetName, objIdxHeader, lLigne, sPosValeur)
                    lStat(0) = lStat(0) + 1
                Else
                    sPosValeur = objIndexLocal(sValeurPkFinale)
                    ReDim Preserve sPosValeur(UBound(sPosValeur) + 1) As String
                    Call CallByName(m_objCustomIndex, sFunctionValues, VbMethod, sSheetName, objIdxHeader, lLigne, sPosValeur)
                    Call objIndexLocal.Remove(sValeurPkFinale)
                    lStat(1) = lStat(1) + 1
                End If
                '
                ' Ajoute le tout
                Call objIndexLocal.Add(sValeurPkFinale, sPosValeur)
            Else
                lStat(2) = lStat(2) + 1
            End If
        End If
        
        If (lLigne Mod 500) = 0 Then
            Application.StatusBar = "Index: " + sIndexName + ", From: " + sSheetName + ", Fct: " + sFunctionValues + ", Where: " + sFunctionWhere + ", Progression: " & lLigne & " / " & lMaxLine & " : " & Format(lLigne / lMaxLine, "Percent")
        End If
    Next
    
    
    Call objTrace.writeInfo("...STATISTIQUES - FROM: " + sSheetName + ", INDEX: " + sIndexName + ", FCT:" + sFunctionValues + ", WHERE: " + sFunctionWhere)
    Call objTrace.writeInfo("...INSERTION.............: " & lStat(0))
    Call objTrace.writeInfo("...MODIFICATION..........: " & lStat(1))
    Call objTrace.writeInfo("...PK VIDE...............: " & lStat(2))
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("FIN : Creation de l'index.")
    Call objTrace.displayMessage
    
    createIndexMultiValue = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    Call objTrace.writeError("...Classe......: ManageIndex")
    Call objTrace.writeError("...Fonction....: createIndexMultiValue")
    Call objTrace.writeError(".....Numéro.......: " & lErr)
    Call objTrace.writeError(".....Explication..: " & sDescription)
    Call objTrace.writeError(".....Source.......: " & sSource)
    Call objTrace.writeError(".....lLigne.......: " & lLigne)
    Call objTrace.displayMessage
        
    On Error GoTo 0
   
    Err.Raise lErr, , sDescription
End Function



' ----------------------------------------------------------------
' Procedure Name: createIndexSum
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objWorkBook (Workbook):
' Parameter sSheetName (String):
' Parameter sColNameCount (String):
' Parameter bClearIndex (Boolean):
' Parameter sListIndexName (String):
' Parameter sFunctionGroupBy (String):
' Parameter sFunctionAgregat (String):
' Parameter sFunctionWhere (String):
' Return Type: Integer
' Author: XMXU139
' Date: 03/07/2019
' ----------------------------------------------------------------
Public Function createIndexSum( _
    ByRef r_objWorkBook As Workbook, _
    ByVal sSheetName As String, _
    ByVal sColNameCount As String, _
    ByVal bClearIndex As Boolean, _
    ByVal sListIndexName As String, _
    ByVal sFunctionGroupBy As String, _
    ByVal sFunctionAgregat As String, _
    ByVal sFunctionWhere As String) As Integer
    
    On Error GoTo erreurGenerale
    
    Dim objTrace              As New ManageLog
    Dim objOnglet             As New ManageOnglet
    
    Dim objIdxHeader          As Scripting.Dictionary
    Dim objIndexLocal         As Scripting.Dictionary
    
    Dim sValeurPkFinale       As String
    Dim sIndexName()          As String
    Dim lMaxLine              As Long
    Dim lLigne                As Long
    Dim lMaxIndex             As Long
    Dim lCptIndex             As Long
    Dim lStat(4)              As Long
    Dim dPosValeur            As Double
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim iReturnVal            As Integer
    
    '
    ' 01 - Par défaut retour pésimiste
    createIndexSum = -1
    '
    ' 02 - Initialisation des lStats
    lStat(0) = 0: lStat(1) = 0: lStat(2) = 0: lStat(3) = 0
    '
    ' 03 - Initialisation du manager
    Call objTrace.InitManageLog
    '
    ' 04 - Mise en place de l'espace de travail et récupère les colonnes
    If r_objWorkBook Is Nothing Then
        r_objWorkBook = ThisWorkbook
    End If
    '
    ' 05 - Mise en  place des header
    Call objOnglet.getColumnWithIndexPosition(r_objWorkBook, sSheetName, objIdxHeader)
    '
    ' 06 - Signaler le debut de l'indexation
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("DEBUT : Creation createIndexSum dans " + sSheetName + ", Agregat : " + sFunctionAgregat + ", Regroupement : " + sFunctionGroupBy + ", critères: " + sFunctionWhere)
    '
    ' 07 - Remise à zéro des index ?
    If sListIndexName = "" Then
        Err.Raise 610, , "Veuillez déclarer une liste d'index."
    End If
    ' 07.01 - Extraction de la liste des index
    sIndexName = Split(sListIndexName, ",")
    lMaxIndex = UBound(sIndexName)
    If bClearIndex = True Then
        ' Remise à zéro
        For lCptIndex = 0 To lMaxIndex
            Call createIndex(sIndexName(lCptIndex))
        Next lCptIndex
    End If
    '
    ' 08 - Si pas de classe métier
    If IsNull(m_objCustomIndex) Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    If m_objCustomIndex Is Nothing Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    '
    ' 09 - Parcourir toutes les lLignes du fichier
    r_objWorkBook.Sheets(sSheetName).Range(sColNameCount + "1").Select
    lMaxLine = Application.WorksheetFunction.CountA(Range(sColNameCount + ":" + sColNameCount))
    For lLigne = 2 To lMaxLine
        '
        ' ... et parcourir tous les index
        For lCptIndex = 0 To lMaxIndex
            '
            ' Index encours
            Set objIndexLocal = getIndex(sIndexName(lCptIndex))
            '
            ' 10 - Fonction critère ?
            If sFunctionWhere <> "" Then
                iReturnVal = CallByName(m_objCustomIndex, sFunctionWhere, VbMethod, sSheetName, objIdxHeader, lLigne, sIndexName(lCptIndex))
            Else
                iReturnVal = 1
            End If
            '
            ' 11 - Extaction de la valeur de la PK composite
            sValeurPkFinale = CallByName(m_objCustomIndex, sFunctionGroupBy, VbMethod, sSheetName, objIdxHeader, lLigne, sIndexName(lCptIndex))
            '
            ' 12 - Une fois la clé pk composite formée et existante ....
            If sValeurPkFinale <> "" Then
                '
                ' Est-ce que l'order ID existe ?
                dPosValeur = 0
                If Not objIndexLocal.Exists(sValeurPkFinale) Then
                    If iReturnVal = 1 Then
                        dPosValeur = CallByName(m_objCustomIndex, sFunctionAgregat, VbMethod, sSheetName, objIdxHeader, lLigne, "I", sIndexName(lCptIndex))
                    End If
                    Call objIndexLocal.Add(sValeurPkFinale, dPosValeur)
                    lStat(0) = lStat(0) + 1
                Else
                    If iReturnVal = 1 Then
                        dPosValeur = CDbl(objIndexLocal(sValeurPkFinale)) + CallByName(m_objCustomIndex, sFunctionAgregat, VbMethod, sSheetName, objIdxHeader, lLigne, "U", sIndexName(lCptIndex))
                        Call objIndexLocal.Remove(sValeurPkFinale)
                        Call objIndexLocal.Add(sValeurPkFinale, dPosValeur)
                    End If
                    lStat(1) = lStat(1) + 1
                End If
            Else
                lStat(2) = lStat(2) + 1
            End If
            
        Next lCptIndex
        
        If (lLigne Mod 500) = 0 Then
            Application.StatusBar = "Index: From: " + sSheetName + ", Fct: " + sFunctionAgregat + ", Where: " + sFunctionWhere + ", Progression: " & lLigne & " / " & lMaxLine & " : " & Format(lLigne / lMaxLine, "Percent")
        End If
    Next
    
    
    Call objTrace.writeInfo("...STATISTIQUES")
    Call objTrace.writeInfo("...Insertion.............: " & lStat(0))
    Call objTrace.writeInfo("...Modification..........: " & lStat(1))
    Call objTrace.writeInfo("...PK vide...............: " & lStat(2))
    Call objTrace.writeInfo("")
    Call objTrace.writeInfo("FIN : Creation de l'index.")
    Call objTrace.displayMessage
    
    createIndexSum = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
  
    Call objTrace.writeError("...Classe      : ManageIndex")
    Call objTrace.writeError("...Fonction    : createIndexSum")
    Call objTrace.writeError(".....Numéro      : " & lErr)
    Call objTrace.writeError(".....Explication : " & sDescription)
    Call objTrace.writeError(".....Source      : " & sSource)
    Call objTrace.writeError(".....lLigne       : " & lLigne)
    Call objTrace.displayMessage
        
    On Error GoTo 0
   
    Err.Raise lErr, , sDescription
End Function
' ----------------------------------------------------------------
' Nom         : saveIntoModele
' Explication : Sauvegarde l'index dans un fichier texte
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFileName (String):
' Parameter sIndexName (String):
' Parameter sFunctionGetValue (String):
' Author: Eric VENTALON
' Date: 09/05/2019
' ----------------------------------------------------------------
Public Function SaveIntoModele(ByVal sFileName As String, ByVal sIndexName As String, Optional ByVal sFunctionGetValue As String = "getValueBeforeSave") As Integer

    On Error GoTo erreurGenerale

    Dim objFso          As Object
    Dim objFile         As Object
    Dim objParam        As New ManageParam
    Dim objLog          As New ManageLog
    Dim objDico         As Scripting.Dictionary
    Dim vValue          As Variant
    Dim sFullText       As String
    Dim sPath           As String
    
    Dim lErr          As Long
    Dim sDescription  As String
    Dim sSource       As String

    
    Call objLog.InitManageLog
    Call objParam.initManageParam(objLog)
    
    If IsNull(m_objCustomIndex) Or m_objCustomIndex Is Nothing Then
        Set m_objCustomIndex = Me
    End If
        
    SaveIntoModele = 1
    
    '
    ' Extraire le parametre du chemin
    '
    sPath = objParam.getParams("P_OUTPUT_MODELE")
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.CreateTextFile(sPath + sFileName)
    
    Set objDico = getIndex(sIndexName)
    If IsNull(objDico) Then Exit Function
    
    For Each vValue In objDico.Keys()
        sFullText = CallByName(m_objCustomIndex, sFunctionGetValue, VbMethod, vValue, objDico(vValue))
        sFullText = sFullText & vbCrLf
        objFile.Write sFullText
    Next vValue
    
    Call objFile.Close

    Set objFile = Nothing
    Set objFso = Nothing
    
    
    SaveIntoModele = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source

        
    On Error GoTo 0
   
     Err.Raise lErr, sSource + " - ManageIndex.SaveIntoModele() ", sDescription

End Function

' ----------------------------------------------------------------
' Procedure Name: SaveIntoFile
' Purpose: Sauvegarde l'index dans un fichier
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFileName (String): Nom du fichier avec extention ou pas
' Parameter sPathNamePara (String): La paramétre du répertoire
' Parameter sIndexName (String): Le nom de l'index
' Parameter sFunctionGetHeader (String): La fonction de génération des entêtes
' Parameter sFunctionGetValue (String): La fonction de presentation des lignes
' Return Type: Integer
' Author: Eric VENTALON
' Date: 06/06/2019
' ----------------------------------------------------------------
Public Function SaveIntoFile(ByVal sFileName As String, _
    ByVal sPathNamePara As String, _
    ByVal sIndexName As String, _
    ByVal sFunctionGetHeader As String, _
    ByVal sFunctionGetValue As String) As Integer

    On Error GoTo erreurGenerale

    Dim objFso          As Object
    Dim objFile         As Object
    Dim objParam        As New ManageParam
    Dim objLog          As New ManageLog
    Dim objDico         As Scripting.Dictionary
    Dim vValue          As Variant
    Dim sFullText       As String
    Dim sPath           As String
    
    Dim lErr          As Long
    Dim sDescription  As String
    Dim sSource       As String

    
    Call objLog.InitManageLog
    Call objParam.initManageParam(objLog)
    
    If IsNull(m_objCustomIndex) Or m_objCustomIndex Is Nothing Then
        Set m_objCustomIndex = Me
    End If
        
    SaveIntoFile = 1
    
    '
    ' Extraire le parametre du chemin
    '
    sPath = objParam.getParams(sPathNamePara)
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.CreateTextFile(sPath + sFileName)
    
    Set objDico = getIndex(sIndexName)
    If IsNull(objDico) Then Exit Function
    
    sFullText = CallByName(m_objCustomIndex, sFunctionGetHeader, VbMethod)
    sFullText = sFullText & vbCrLf
    objFile.Write sFullText
    
    For Each vValue In objDico.Keys()
        sFullText = CallByName(m_objCustomIndex, sFunctionGetValue, VbMethod, vValue, objDico(vValue))
        sFullText = sFullText & vbCrLf
        objFile.Write sFullText
    Next vValue
    
    Call objFile.Close

    Set objFile = Nothing
    Set objFso = Nothing
    
    SaveIntoFile = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " - ManageIndex.SaveIntoFile() ", sDescription

End Function
' ----------------------------------------------------------------
' Procedure Name : loadFromModele
' Purpose : Lecture du fichier csv et cration de l'index.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFileName (String):
' Parameter sIndexName (String):
' Parameter bClearIndex (Boolean):
' Parameter sFunctionGetValue (String):
' Return Type: Integer
' Author: Eric VENTALON
' Date: 10/05/2019
' ----------------------------------------------------------------
Public Function loadFromModele( _
    ByVal sFileName As String, _
    ByVal sIndexName As String, _
    Optional ByVal bClearIndex As Boolean = True, _
    Optional ByVal sFunctionGetValue As String = "getValueAfterLoad" _
    ) As Integer
    
    On Error GoTo erreurGenerale
    
    Dim objFso          As Object
    Dim objFile         As Object
    Dim objParam        As New ManageParam
    Dim objLog          As New ManageLog
    Dim objDico         As Scripting.Dictionary
    Dim sFullText       As String
    Dim sTabtext()      As String
    
    Dim lErr          As Long
    Dim sDescription  As String
    Dim sSource       As String
    
    Dim sPath As String
    
    Call objLog.InitManageLog
    Call objParam.initManageParam(objLog)
    
    If IsNull(m_objCustomIndex) Or m_objCustomIndex Is Nothing Then
        Set m_objCustomIndex = Me
    End If
        
    loadFromModele = 1
    
    '
    ' Extraire le parametre du chemin
    '
    sPath = objParam.getParams("P_OUTPUT_MODELE")
    Call objLog.writeInfo("")
    Call objLog.writeInfo("Chargement de l'index " + sIndexName + " venant de " + sPath + sFileName)
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.OpenTextFile(sPath + sFileName)
    
    If bClearIndex = True Then Call createIndex(sIndexName)
    Set objDico = getIndex(sIndexName)
    If IsNull(objDico) Then Exit Function
    
    Do While objFile.AtEndOfStream <> True
        sFullText = objFile.readLine
        ' La premiere colonne = la clé primaire
        ' La deuxime colonne = la valeur
        sTabtext = CallByName(m_objCustomIndex, sFunctionGetValue, VbMethod, sFullText)
        ' sTabtext = Split(sFullText, ";")
        If UBound(sTabtext) = 1 Then
            Call objDico.Add(sTabtext(0), sTabtext(1))
        End If
        If UBound(sTabtext) = 2 Then
            Call objDico.Add(sTabtext(0), sTabtext(1) + ";" + sTabtext(2))
        End If
    Loop
        
    Call objFile.Close

    Set objFile = Nothing
    Set objFso = Nothing
    
    Call objLog.displayMessage
    
    loadFromModele = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " - ManageIndex.loadFromModele() ", sDescription

End Function


' ----------------------------------------------------------------
' Procedure Name : linkObjectBuisness
' Purpose : Permet de lier votre classe metier avec celle de Manageindex
' Procedure Kind: Propriété (Set)
' Procedure Access: Public
' Parameter objNewValue (Object): Une instance de la classe métier
' Author: Eric VENTALON
' Date: 10/05/2019
' ----------------------------------------------------------------
Public Property Set linkObjectBuisness(ByVal objNewValue As Object)

    Set m_objCustomIndex = objNewValue

End Property


' ----------------------------------------------------------------
' Procedure Name: getValueAfterLoad
' Purpose: Découpe une ligne du fichier séparé par ; en tableau de string
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sLigneFichier (String): Ligne entiere du fichier
' Return Type: String()
' Author: Eric VENTALON
' Date: 16/07/2019
' ----------------------------------------------------------------
Public Function getValueAfterLoad(ByVal sLigneFichier As String) As String()
    getValueAfterLoad = Split(sLigneFichier, ";")
End Function

' ----------------------------------------------------------------
' Procedure Name : getValueBeforeSave
' Purpose : Extraire la ligne de l'index avant sauvegarde sous la forme clé ; valeur
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_vPrimaryKey (Variant): La clé primaire de l'index en variant
' Parameter r_vDico (Variant): Les valeurs associées en variant
' Return Type: String
' Author: Eric VENTALON
' Date: 21/05/2019
' ----------------------------------------------------------------
Public Function getValueBeforeSave(ByRef r_vPrimaryKey As Variant, ByRef r_vDico As Variant) As String
    

    Dim sPrimaryKey As String
    Dim sTuple As String
    
    sPrimaryKey = CStr(r_vPrimaryKey)
    sTuple = CStr(r_vDico)
    
    getValueBeforeSave = sPrimaryKey + ";" + sTuple
    

End Function
' ----------------------------------------------------------------
' Procedure Name: createIndexUniqueInArray
' Purpose: Creation d'un index à partir d'un tableau mémoire
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objArraySheet (Variant()): Le tableau
' Parameter r_objMetier (Object): La référence de l'objet métier
' Parameter r_objIdxHeader (Dictionary): La référence des entêtes
' Parameter bClearIndex (Boolean): True = Efface l'index, false préserve l'index
' Parameter sIndexName (String): Nom de l'index
' Parameter sFunctionSelect (String): Nom de la fonction de PK
' Parameter sFunctionValues (String): Nom de la fonction de la dépendance
' Parameter sFunctionWhere (String): Nom de la fonction de filtre
' Return Type: Integer
' Author: XMXU139
' Date: 02/07/2019
' ----------------------------------------------------------------
Public Function createIndexUniqueInArray( _
    ByRef r_objArraySheet() As Variant, _
    ByRef r_objMetier As Object, _
    ByRef r_objIdxHeader As Scripting.Dictionary, _
    ByVal bClearIndex As Boolean, _
    ByVal sIndexName As String, _
    ByVal sFunctionSelect As String, _
    ByVal sFunctionValues As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
    
    On Error GoTo erreurGenerale
    
    Dim objIndexLocal         As Scripting.Dictionary
    
    Dim sValeurPkFinale       As String
    Dim lMaxLine              As Long
    Dim lLigne                As Long
    Dim sPosValeur            As String
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim iReturnVal            As Integer
    
    '
    ' 01 - Par défaut retour pésimiste
    createIndexUniqueInArray = -1
    '
    ' 02 - Remise à zéro de l'index ?
    If bClearIndex = True Then Call createIndex(sIndexName)
    Set objIndexLocal = getIndex(sIndexName)
    '
    ' 03 - Si pas de classe métier
    If IsNull(r_objMetier) Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    If r_objMetier Is Nothing Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    '
    ' 04 - Parcourir toutes les lLignes du fichier
    lMaxLine = UBound(r_objArraySheet, 1)
    For lLigne = 1 To lMaxLine
        '
        ' 05 - Fonction critère ?
        If sFunctionWhere <> "" Then
            iReturnVal = CallByName(r_objMetier, sFunctionWhere, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne)
        Else
            iReturnVal = 1
        End If
        '
        ' 06 - creation de l'index
        If iReturnVal = 1 Then
            '
            ' 06.1 - Extaction de la valeur de la PK composite
            sValeurPkFinale = CallByName(r_objMetier, sFunctionSelect, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne)
            '
            ' 06.2 - Une fois la clé pk composite formée et existante ....
            If sValeurPkFinale <> "" Then
                '
                ' 06.1.1 - Est-ce que l'order ID existe ?
                If Not objIndexLocal.Exists(sValeurPkFinale) Then
                    sPosValeur = CallByName(r_objMetier, sFunctionValues, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne)
                    Call objIndexLocal.Add(sValeurPkFinale, sPosValeur)
                End If
            End If
        End If
        
        If (lLigne Mod 500) = 0 Then
            Application.StatusBar = "Index: " + sIndexName + ", Progression: " & lLigne & " / " & lMaxLine & " : " & Format(lLigne / lMaxLine, "Percent")
        End If
    Next
    
    createIndexUniqueInArray = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " - ManageIndex.createIndexSumInArray()", sDescription
    
    On Error GoTo 0
    Exit Function


End Function



' ----------------------------------------------------------------
' Procedure Name: createIndexSumInArray
' Purpose: Création d'un index avec une clé primaire déterminant une valeur
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objArraySheet (Variant()): Le tableau
' Parameter r_objMetier (Object): La référence de l'objet métier
' Parameter r_objIdxHeader (Dictionary): La référence des entêtes
' Parameter bClearIndex (Boolean): true :  nouvel index, false : réutilisation de l'index
' Parameter sListIndexName (String): Nom de l'index
' Parameter sFunctionGroupBy (String): Fonction de projection des colonnes servant pour la PK
' Parameter sFunctionAgregat (String): Fonction d'extraction des valeurs
' Parameter sFunctionWhere (String): Fonction critères
' Return Type: Integer
' Author: Eric VENTALON
' Date: 16/07/2019
' ----------------------------------------------------------------
Public Function createIndexSumInArray(ByRef r_objArraySheet() As Variant, ByRef r_objMetier As Object, ByRef r_objIdxHeader As Scripting.Dictionary, _
    ByVal bClearIndex As Boolean, ByVal sListIndexName As String, ByVal sFunctionGroupBy As String, ByVal sFunctionAgregat As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
        
    On Error GoTo erreurGenerale
    
    Dim objIndexLocal         As Scripting.Dictionary
    
    Dim sValeurPkFinale       As String
    Dim sIndexName()          As String
    Dim lMaxLine              As Long
    Dim lLigne                As Long
    Dim lMaxIndex             As Long
    Dim lCptIndex             As Long
    Dim dPosValeur            As Double
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim iReturnVal            As Integer
    
    '
    ' 01 - Par défaut retour pésimiste
    createIndexSumInArray = -1
    '
    ' 02 - Remise à zéro des index ?
    If sListIndexName = "" Then
        Err.Raise 610, , "Veuillez déclarer une liste d'index."
    End If
    ' 02.01 - Extraction de la liste des index
    sIndexName = Split(sListIndexName, ",")
    lMaxIndex = UBound(sIndexName)
    If bClearIndex = True Then
        ' Remise à zéro
        For lCptIndex = 0 To lMaxIndex
            Call createIndex(sIndexName(lCptIndex))
        Next lCptIndex
    End If
    '
    ' 03 - Si pas de classe métier
    If IsNull(r_objMetier) Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    If r_objMetier Is Nothing Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    '
    ' 04 - Parcourir toutes les lLignes du fichier
    lMaxLine = UBound(r_objArraySheet, 1)
    For lLigne = 1 To lMaxLine
        '
        ' ... et parcourir tous les index
        For lCptIndex = 0 To lMaxIndex
            '
            ' Index encours
            Set objIndexLocal = getIndex(sIndexName(lCptIndex))
            '
            ' A - Fonction critère ?
            If sFunctionWhere <> "" Then
                iReturnVal = CallByName(r_objMetier, sFunctionWhere, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne, sIndexName(lCptIndex))
            Else
                iReturnVal = 1
            End If
            '
            ' B - Extraction de la valeur de la PK composite
            sValeurPkFinale = CallByName(r_objMetier, sFunctionGroupBy, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne, sIndexName(lCptIndex))
            '
            ' C - Une fois la clé pk composite formée et existante ....
            If sValeurPkFinale <> "" And iReturnVal = 1 Then
                '
                ' Est-ce que l'order ID existe ?
                dPosValeur = 0
                If Not objIndexLocal.Exists(sValeurPkFinale) Then
                    dPosValeur = CallByName(r_objMetier, sFunctionAgregat, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne, "I", sIndexName(lCptIndex))
                Else
                    dPosValeur = CDbl(objIndexLocal(sValeurPkFinale)) + CallByName(r_objMetier, sFunctionAgregat, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne, "U", sIndexName(lCptIndex))
                    Call objIndexLocal.Remove(sValeurPkFinale)
                End If
                Call objIndexLocal.Add(sValeurPkFinale, dPosValeur)
                
            End If
            
        Next lCptIndex
        
        If (lLigne Mod 500) = 0 Then
            Application.StatusBar = "Index cumul : Progression: " & lLigne & " / " & lMaxLine & " : " & Format(lLigne / lMaxLine, "Percent")
        End If
        
    Next lLigne
    
        
    createIndexSumInArray = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, Err.Source + " : ManageIndex.createIndexSumInArray()", sDescription
    
End Function


' ----------------------------------------------------------------
' Procedure Name: createIndexMultiValueInArray
' Purpose: Création d'un index avec une clé primaire déterminant un tableau de valeur
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objArraySheet (Variant()): Le tableau
' Parameter r_objMetier (Object): La référence de l'objet métier
' Parameter r_objIdxHeader (Dictionary): La référence des entêtes
' Parameter bClearIndex (Boolean): true :  nouvel index, false : réutilisation de l'index
' Parameter sIndexName (String): Nom des index séparé par une virgule.
' Parameter sFunctionSelect (String): Fonction de projection des colonnes servant pour la PK
' Parameter sFunctionValues (String): Fonction d'extraction des valeurs
' Parameter sFunctionWhere (String): Fonction critères
' Return Type: Integer
' Author: Eric VENTALON
' Date: 16/07/2019
' ----------------------------------------------------------------
Public Function createIndexMultiValueInArray( _
    ByRef r_objArraySheet() As Variant, _
    ByRef r_objMetier As Object, _
    ByRef r_objIdxHeader As Scripting.Dictionary, _
    ByVal bClearIndex As Boolean, _
    ByVal sIndexName As String, _
    ByVal sFunctionSelect As String, _
    ByVal sFunctionValues As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
    
    On Error GoTo erreurGenerale
    
    Dim objIndexLocal         As Scripting.Dictionary
    
    Dim sValeurPkFinale       As String
    Dim lMaxLine              As Long
    Dim lLigne                As Long
    
    Dim sPosValeur()          As String
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim iReturnVal            As Integer
    '
    ' 01 - Par défaut retour pésimiste
    createIndexMultiValueInArray = -1
    '
    ' 02 - Remise à zéro de l'index ?
    If bClearIndex = True Then Call createIndex(sIndexName)
    Set objIndexLocal = getIndex(sIndexName)
    '
    ' 03 - Si pas de classe métier
    If IsNull(r_objMetier) Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    If r_objMetier Is Nothing Then
        Err.Raise 600, , "Veuillez déclarer une classe metier pour les critères, extractions, fonctions."
    End If
    '
    ' 04 - Parcourir toutes les lLignes du fichier
    lMaxLine = UBound(r_objArraySheet, 1)
    For lLigne = 1 To lMaxLine
        '
        ' 09 - Fonction critère ?
        If sFunctionWhere <> "" Then
            iReturnVal = CallByName(r_objMetier, sFunctionWhere, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne)
        Else
            iReturnVal = 1
        End If
        '
        ' 10 - creation de l'index
        If iReturnVal = 1 Then
            '
            ' 11 - Extaction de la valeur de la PK composite
            sValeurPkFinale = CallByName(r_objMetier, sFunctionSelect, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne)
            '
            ' 12 - Une fois la clé pk composite formée et existante ....
            If sValeurPkFinale <> "" Then
                '
                ' Est-ce que l'order ID existe ?
                If Not objIndexLocal.Exists(sValeurPkFinale) Then
                    ReDim sPosValeur(0) As String
                    Call CallByName(r_objMetier, sFunctionValues, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne, sPosValeur)
                Else
                    sPosValeur = objIndexLocal(sValeurPkFinale)
                    ReDim Preserve sPosValeur(UBound(sPosValeur) + 1) As String
                    Call CallByName(r_objMetier, sFunctionValues, VbMethod, r_objArraySheet, r_objIdxHeader, lLigne, sPosValeur)
                    Call objIndexLocal.Remove(sValeurPkFinale)
                End If
                '
                ' Ajoute le tout
                Call objIndexLocal.Add(sValeurPkFinale, sPosValeur)
            End If
        End If
        
        If (lLigne Mod 500) = 0 Then
            Application.StatusBar = "Index multi-values : Progression: " & lLigne & " / " & lMaxLine & " : " & Format(lLigne / lMaxLine, "Percent")
        End If
    Next
    
    createIndexMultiValueInArray = 0
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, Err.Source + " : ManageIndex.createIndexMultiValueInArray()", sDescription
End Function

