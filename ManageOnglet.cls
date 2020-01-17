VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageOnglet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: ManageOnglet
' Help: https://app.gitbook.com/@eric-ventalon/s/vba-framework/
' Kind: Module de classe
' Purpose: Utilitaire pour une feuille
' Author: Eric VENTALON
' Date: 28/06/2019
' ------------------------------------------------------
Option Explicit

'
'   _____                _     _                         _                  ______                  _   _ _
'  / ____|              | |   (_)                       | |                |  ____|                (_) | | | |
' | |  __    ___   ___  | |_   _    ___    _ __       __| |   ___   ___    | |__    ___   _   _   _  | | | |   ___   ___
' | | |_ |  / _ \ / __| | __| | |  / _ \  | '_ \     / _` |  / _ \ / __|   |  __|  / _ \ | | | | | | | | | |  / _ \ / __|
' | |__| | |  __/ \__ \ | |_  | | | (_) | | | | |   | (_| | |  __/ \__ \   | |    |  __/ | |_| | | | | | | | |  __/ \__ \
'  \_____|  \___| |___/  \__| |_|  \___/  |_| |_|    \__,_|  \___| |___/   |_|     \___|  \__,_| |_| |_| |_|  \___| |___/
'

'
' Liens sur un ensemble de feuille
Private m_objSetSheet               As New Scripting.Dictionary

Public Sub setSheet( _
    ByVal sKeySheet As String, _
    ByRef r_objWorkBook As Workbook, _
    ByVal sSheetName As String, _
    ByVal sColCount As String, _
    ByRef r_objClasseMetier As Object, _
    Optional lHearderColumn As Long = 1, _
    Optional sSeparateur As String = "", _
    Optional lHearderLine As Long = 1)
    
    On Error GoTo erreurGenerale
       
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
    Dim sNameColonne    As String
    
    Dim lLigneMax       As Long
    Dim lCpt            As Long
    Dim iIndexSheet     As Integer
    Dim objArray        As New ManageArray
    Dim objHeaderArray  As New Scripting.Dictionary
    '
    ' Vérifications
    If Trim(sKeySheet) = "" Then
        Err.Raise 800, "ManageOnglet.setSheet()", "La clé " + sKeySheet + " est vide."
    End If
    
    If m_objSetSheet.Exists(sKeySheet) Then
        Err.Raise 801, "ManageOnglet.setSheet()", "La clé " + sKeySheet + " est déjà existante."
    End If
    '
    ' Overture du fichier et active une feuille
    ' RG1 : Si le nom de la feuille n'est pas connu alors détecter uniquement la premiere feuille
    If sSheetName = "" Then
        sSheetName = r_objWorkBook.Sheets(1).name
    End If
    '
    ' Préparation de la feuille
    r_objWorkBook.Sheets(sSheetName).Select
    iIndexSheet = ActiveSheet.index
    '
    ' Calcule le nombre de ligne maxi
    lLigneMax = r_objWorkBook.Sheets(iIndexSheet).Range(sColCount & Rows.Count).End(xlUp).Row
        
    Application.DisplayAlerts = False
    '
    ' Séparateurs ; , t, ...
    If sSeparateur <> "" Then
        If sSeparateur = ";" Then
            r_objWorkBook.Sheets(iIndexSheet).Columns(1).TextToColumns DataType:=xlDelimited, Semicolon:=True
        ElseIf sSeparateur = "," Then
            r_objWorkBook.Sheets(iIndexSheet).Columns(1).TextToColumns DataType:=xlDelimited, comma:=True
        ElseIf sSeparateur = "t" Then
            r_objWorkBook.Sheets(iIndexSheet).Columns(1).TextToColumns DataType:=xlDelimited, Tab:=True
        ElseIf sSeparateur = " " Then
            r_objWorkBook.Sheets(iIndexSheet).Columns(1).TextToColumns DataType:=xlDelimited, Space:=True
        Else
            r_objWorkBook.Sheets(iIndexSheet).Columns(1).TextToColumns DataType:=xlDelimited, other:=True, OtherChar:=sSeparateur
        End If
    End If
    
    Application.DisplayAlerts = True
    '
    ' Mémorise les colonnes du fichier
    If lHearderLine > 0 Then
        Dim lHearderLineSauve As Long
        
        lHearderLineSauve = lHearderColumn
        sNameColonne = r_objWorkBook.Sheets(iIndexSheet).Cells(lHearderLine, lHearderLineSauve).Text
        lCpt = 1
        While sNameColonne <> ""
            Call objHeaderArray.Add(sNameColonne, lCpt)
            lHearderLineSauve = lHearderLineSauve + 1
            lCpt = lCpt + 1
            sNameColonne = r_objWorkBook.Sheets(iIndexSheet).Cells(lHearderLine, lHearderLineSauve).Text
        Wend
        
    End If
    '
    ' Affectation de l'objet feuille
    Call objArray.setSheet(r_objWorkBook, _
        sSheetName, _
        iIndexSheet, _
        sColCount, _
        r_objClasseMetier, _
        lHearderColumn, _
        sSeparateur, _
        objHeaderArray, _
        lHearderLine)
    '
    ' Ajout dans l'ensemble des feuilles
    Call m_objSetSheet.Add(sKeySheet, objArray)
       
    On Error GoTo 0
    Exit Sub
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageOnglet.setSheet()", sDescription
 
End Sub


Private Sub extractColLine(ByVal sZone As String, _
    ByRef r_sColLeft As String, _
    ByRef r_sColRight As String, _
    ByRef r_sLigneDebut As String)
    
    Dim iPos            As Integer
    Dim bCoteGauche    As Boolean
    Dim sCar            As String
      
    r_sColLeft = "A"
    r_sColRight = ""
    r_sLigneDebut = "1"
    
    If sZone = "" Then
        Exit Sub
    End If
    
    r_sColLeft = ""
    r_sLigneDebut = ""
    r_sLigneDebut = ""
    
    bCoteGauche = True
    
    sZone = UCase(sZone)
    For iPos = 1 To Len(sZone)
    
        sCar = Mid(sZone, iPos, 1)
        If sCar >= "A" And sCar <= "Z" Then
            If bCoteGauche Then
                r_sColLeft = r_sColLeft + sCar
            Else
                r_sColRight = r_sColRight + sCar
            End If
        ElseIf IsNumeric(sCar) Then
            If bCoteGauche Then
                r_sLigneDebut = r_sLigneDebut + sCar
            End If
        Else
            bCoteGauche = False
        End If
        
    Next iPos
    
    If r_sLigneDebut = "" Then r_sLigneDebut = "1"
   
End Sub


' sdropCol = "A" ou "AB" ou "no_commande" ou Numéro de facture"
Public Function dropColumn(ByVal sKeySheet As String, ByVal sdropCol As String) As Boolean
    
    On Error GoTo erreurGenerale
    
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
    
    Dim iSheetIndex     As Integer
    
    Dim objWkb          As Workbook
    Dim objArray        As ManageArray
    Dim objHeaderArray  As Scripting.Dictionary
            
    dropColumn = False
    
    If Not m_objSetSheet.Exists(sKeySheet) Then
        Err.Raise 651, "", "La clé " + sKeySheet + "n existe pas."
    End If
    
    Set objArray = m_objSetSheet(sKeySheet)
    
    Set objWkb = objArray.getWorkbook
    Set objHeaderArray = objArray.getHeader
    iSheetIndex = objArray.getSheetIndex
        
    ' Identification de la colonne par son nom de métier ?
    If objArray.getColumnMax > 0 Then
        If objHeaderArray.Exists(sdropCol) Then
            sdropCol = nomDeColonne(objHeaderArray(sdropCol))
        End If
    End If
        
    objWkb.Sheets(iSheetIndex).Range(sdropCol + ":" + sdropCol).EntireColumn.delete
    '
    ' recalcul des colonnes
    If objArray.getColumnMax > 0 Then
        Call getColumnWithIndexPosition(objWkb, Sheets(iSheetIndex).name, objHeaderArray)
        Call objArray.SetHeader(objHeaderArray)
    End If
    
    dropColumn = True
    
    On Error GoTo 0
    Exit Function
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageOnglet.dropColumn()", sDescription
 
End Function


Public Function insertColumn(ByVal sKeySheet As String, ByVal sInsertCol As String, Optional sNomHeader As String = "") As Boolean
    
    On Error GoTo erreurGenerale
    
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
    
        
    insertColumn = False
    
    Dim iSheetIndex     As Integer
    
    Dim objWkb          As Workbook
    Dim objArray        As ManageArray
    Dim objHeaderArray  As Scripting.Dictionary
            
    insertColumn = False
    
    If Not m_objSetSheet.Exists(sKeySheet) Then
        Err.Raise 652, "", "La clé " + sKeySheet + "n existe pas."
    End If
    
    Set objArray = m_objSetSheet(sKeySheet)
    Set objWkb = objArray.getWorkbook
    Set objHeaderArray = objArray.getHeader
    iSheetIndex = objArray.getSheetIndex
        
    objWkb.Sheets(iSheetIndex).Range(sInsertCol + "1").EntireColumn.insert
    If objArray.getColumnMax() Then
        objWkb.Sheets(iSheetIndex).Range(sInsertCol + "1").Value = sNomHeader
    End If
    '
    ' recalcul des colonnes
    If objArray.getColumnMax() Then
        Call getColumnWithIndexPosition(objWkb, Sheets(iSheetIndex).name, objHeaderArray)
        Call objArray.SetHeader(objHeaderArray)
    End If
    
    insertColumn = True
    
    On Error GoTo 0
    Exit Function
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageOnglet.insertColumn()", sDescription
 
End Function


' ----------------------------------------------------------------
' Procedure Name: checkToArray
' Purpose: Place la zone de la feuille en mémoire
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sKeySheet (String): Clé primaire de la zone
' Parameter sZone (String): La zone à mémoriser.
' Return Type: ManageArray
' Author: XMXU139
' Date: 15/12/2019
' ----------------------------------------------------------------
Public Function checkToArray(ByVal sKeySheet As String, Optional ByVal sZone As String = "", Optional ByVal bFermeture As Boolean = True) As ManageArray
    
    On Error GoTo erreurGenerale
    
    Dim sDescription    As String
    Dim sSource         As String
    Dim sColGauche      As String
    Dim sColDroite      As String
    Dim sNumLigne       As String
    
    Dim iSheetIndex     As Integer
    
    Dim lErr, lLigneMax As Long
    
    Dim lColDepHeader, lRowDepHeader As Long
    
    Dim objWkb          As Workbook
    Dim objArray        As ManageArray
    Dim objHeaderArray  As Scripting.Dictionary
    
    
    Set checkToArray = Nothing
   
    If Not m_objSetSheet.Exists(sKeySheet) Then
        Err.Raise 653, "", "La clé " + sKeySheet + "n existe pas."
    End If
    
    Set objArray = m_objSetSheet(sKeySheet)
    
    Set objWkb = objArray.getWorkbook
    Set objHeaderArray = objArray.getHeader
    iSheetIndex = objArray.getSheetIndex
    '
    ' La derniere ligne du tableau
    '
    lLigneMax = objWkb.Sheets(iSheetIndex).Range(objArray.getNameColumnCount() & Rows.Count).End(xlUp).Row
    '
    ' Récupére la position de départ des entêtes
    '
    lColDepHeader = objArray.getPosColumnHeader()
    If lColDepHeader = 0 Then lColDepHeader = 1
    lRowDepHeader = objArray.getPosRowHeader()
    '
    ' La zone à mémoriser est vide --> Donc déduction de la dimension de la matrice
    '
    If sZone = "" Then
        ' ------------------------------------------------------------------
        ' CAS 1 : Pas d'entête spécifié     -> Donne A1:AG5220
        ' CAS 2 : Entête spécifié           -> Donne A2:AG5220
        ' CAS 3 : Entête spécifié. Exemple colonne de départ 10, ligne de
        '         départ 15 -> Donne J15:AG5220
        ' ------------------------------------------------------------------
        '
        ' Zone à mémoriser
        '
        sZone = Me.nomDeColonne(CInt(lColDepHeader)) & lRowDepHeader + 1 & ":" + Me.nomDeColonne(CInt(objHeaderArray.Count())) & lLigneMax
    Else
        ' ------------------------------------------------------------------
        ' Cas R3:      --> R3:AG200
        ' Cas R3:AG:   --> R3:AG200
        ' Cas R3:AG200 --> R3:AG200
        ' ------------------------------------------------------------------
        If objHeaderArray.Count() > 0 Then
            '
            ' Extraction des proprietes de la zone. AAXX:AAXX
            Call extractColLine(sZone, sColGauche, sColDroite, sNumLigne)
            '
            ' Ligne de début + 1. A1 -> Donne A2
            sNumLigne = CStr(CLng(sNumLigne) + 1)
            ' A2: --> Donne A2:AG
            If sColDroite = "" Then
                sColDroite = Me.nomDeColonne(objHeaderArray.Count)
            End If
            '
            ' Formation de la zone de donnée à mémoriser sans les entêtes
            sZone = sColGauche + sNumLigne + ":" + sColDroite & lLigneMax
        End If
    End If
    '
    ' Mise en mémoire de la feuille
    Call objArray.sheetToArray(objWkb, iSheetIndex, sZone)
    '
    ' Fermeture du fichier
    If bFermeture Then
       Call objWkb.Close(False)
    End If
    '
    ' Retour OK
    Set checkToArray = objArray
    
    On Error GoTo 0
    Exit Function
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    Call objWkb.Close(False)
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : ManageOnglet.checkToArray()", sDescription
 
End Function




Public Sub dropArray(ByVal sKeySheet As String)
    
    On Error GoTo erreurGenerale
    
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
    Dim objArray        As ManageArray
    
    If Not m_objSetSheet.Exists(sKeySheet) Then
        Exit Sub
    End If
    
    Set objArray = m_objSetSheet(sKeySheet)
    
    Call objArray.dropArray
    Call m_objSetSheet.Remove(sKeySheet)
    
    
    
    On Error GoTo 0
    Exit Sub
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
       
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageOnglet.dropArray()", sDescription
 
End Sub


' ----------------------------------------------------------------
' Procedure Name: setAllEnableCalculation
' Purpose: Active ou désactive les optimisations EXCEL
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter bArgOnOff (Boolean): true ou false
' Parameter argWkb (Workbook): Objet Workbook
' Author: Eric VENTALON
' Date: 28/06/2019
' ----------------------------------------------------------------
Sub setAllEnableCalculation(ByVal bArgOnOff As Boolean, argWkb As Workbook)
    Dim objWs As Worksheet
    
    Application.ScreenUpdating = bArgOnOff
    Application.EnableEvents = bArgOnOff
    Application.CalculateBeforeSave = bArgOnOff
    Application.StatusBar = True
    Application.EnableEvents = bArgOnOff
    If bArgOnOff = False Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
    End If
    
    For Each objWs In argWkb.Worksheets
        objWs.EnableCalculation = bArgOnOff
    Next objWs

End Sub

Public Sub getColumInMem(ByRef r_objWorkBook As Workbook, _
    ByVal sNameSheet As String, _
    ByRef r_sTableName() As String, _
    Optional ByVal lLigne As Long = 1, _
    Optional ByVal sColBegin As String = "A", _
    Optional ByVal sColEnd As String = "")
    
    On Error GoTo erreurGenerale
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim lCol                  As Long
    Dim sValeur               As String
    Dim iCpt                  As Integer

    ReDim r_sTableName(2) As String
    
    ' Exemple A -> 1, C -> 3, etc ...
    lCol = CLng(toColNum(sColBegin))
    sValeur = Trim(r_objWorkBook.Worksheets(sNameSheet).Cells(lLigne, lCol).Text)
    iCpt = 1
    While sValeur <> ""
        r_sTableName(iCpt) = sValeur
        lCol = lCol + 1
        
        If sColEnd <> "" Then
            If lCol > CLng(toColNum(sColEnd)) Then
                sValeur = ""
            End If
        End If
        
        iCpt = iCpt + 1
        ReDim Preserve r_sTableName(iCpt) As String
        sValeur = Trim(r_objWorkBook.Worksheets(sNameSheet).Cells(lLigne, lCol).Text)
        
    Wend
        
    On Error GoTo 0
    
    Exit Sub
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    Err.Raise lErr, sSource + " : ManageOnglet.getColumInMem()", sDescription
    
    On Error GoTo 0
    
End Sub


Public Sub getColumnWithIndexPosition( _
    ByRef argWorkBook As Workbook, _
    ByVal argNameSheet As String, _
    ByRef argDico As Scripting.Dictionary, _
    Optional ByVal sZone As String = "")
    
    On Error GoTo erreurGenerale

    Dim iCpt                  As Integer
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    Dim sColonnePK()          As String
    
    Dim sColDroite            As String
    Dim sColGauche            As String
    Dim sNumLigne             As String
    
    '
    ' Place le nom des colonnes et la ligne de départ
    Call extractColLine(sZone, sColGauche, sColDroite, sNumLigne)
    ' Mémroise les colonnes
    Call getColumInMem(argWorkBook, argNameSheet, sColonnePK, sNumLigne, sColGauche, sColDroite)
    '
    ' Initialise le dico si celui-ci est vide
    If Not (argDico Is Nothing) Then
        argDico.RemoveAll
        Set argDico = Nothing
    End If
    
    Set argDico = New Scripting.Dictionary
    
    For iCpt = 1 To UBound(sColonnePK)
        Call argDico.Add(sColonnePK(iCpt), iCpt)
    Next iCpt
    
    On Error GoTo 0
    Exit Sub
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    Err.Raise lErr, sSource + " :  ManageOnglet.getColumnWithIndexPosition()", "@@" + sColonnePK(iCpt) + "@@ : " + sDescription
    
    On Error GoTo 0


End Sub


' ----------------------------------------------------------------
' Procedure Name: getNumberColonneOf
' Purpose: Retourne le rang de la colonne par rapport à son nom d'entête
' Procedure Kind: Function
' Procedure Access: Public
' Parameter argName (String): Nom d'entête
' Return Type: Long
' Author: Eric VENTALON
' Date: 16/07/2019
' ----------------------------------------------------------------
Public Function getNumberColonneOf(ByVal argName As String) As Long
    Dim colonne As Long
    Dim valeur As String
    Dim trouve As Boolean
    
    getNumberColonneOf = -1
    trouve = False
    colonne = 1
    valeur = Trim(Cells(1, colonne).Text)
    While valeur <> "" And trouve = False
        If Trim(UCase(valeur)) = Trim(UCase(argName)) Then
            trouve = True
            getNumberColonneOf = colonne
        Else
            colonne = colonne + 1
            valeur = Trim(Cells(1, colonne).Text)
        End If
    Wend
End Function
' ----------------------------------------------------------------
' Procedure Name: toColNum
' Purpose: A partir de la position de la colonne dans l'alphabet retourne sa position en chiffre. Exemple A donne 1
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sColonne (String): A ou B ou AA etc ...
' Return Type: Long
' Author: Eric VENTALON
' Date: 29/06/2019
' ----------------------------------------------------------------
Public Function toColNum(sColonne As String) As Long
    toColNum = Range(sColonne & 1).Column
End Function
' ----------------------------------------------------------------
' Procedure Name: nomDeColonne
' Purpose: Retourne le nom de la colonne par rapport à son rang. Exemple : colonne 2 donne B
' Procedure Kind: Function
' Procedure Access: Public
' Parameter iPosColonne (Integer): Position en nombre de la colonne
' Return Type: String
' Author: Eric VENTALON
' Date: 29/06/2019
' ----------------------------------------------------------------
Public Function nomDeColonne(iPosColonne As Integer) As String
    On Error GoTo erreurGenerale
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    nomDeColonne = Cells(1, iPosColonne).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    nomDeColonne = CStr(Left(nomDeColonne, Len(nomDeColonne) - 1))
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    Err.Raise lErr, sSource + " :  ManageOnglet.nomDeColonne()", sDescription
    
    On Error GoTo 0
End Function




Public Sub writeStepNameIntoModeOp(ByVal argStepName As String, ByVal argLogAppli As ManageLog)
        
    Dim sSheetName, sParamLastDate, sParamNameFile, sParamLastNumber, sParamNameFileFinal As String
    Dim numLigne As Long
    Dim paramAppli As New ManageParam
      
    Dim objIndex                As New ManageIndex
    Dim objLibMetier            As New Scripting.Dictionary
            
    sParamLastDate = Format(Now(), "yyyyMMdd-hhmmss")
      
    Call argLogAppli.writeInfo("")
    Call argLogAppli.writeInfo("DEBUT : Sauvegarde le fichier")
    '
    ' Initialise les paramètres
    '
    If paramAppli.initManageParam(argLogAppli) < 0 Then
        Call argLogAppli.writeError("...Erreur de paramètres.")
        Call argLogAppli.writeInfo("Fin : Sauvegarde le fichier")
        Exit Sub
    End If
    '
    ' Récupérer le nom du fichier
    '
    sParamNameFile = paramAppli.getParams("P_SYSTEM_FILENAME")
    If sParamNameFile = "" Then
        Call argLogAppli.writeError("...Le nom du fichier générique n'existe pas. Veuillez le definir dans les paramètres.")
        Call argLogAppli.writeError("...Paramètre : P SYSTEM FILENAME")
        Call argLogAppli.writeInfo("Fin : Sauvegarde le fichier")
        Exit Sub
    End If
    '
    ' Récupérer le numéro d'ordre
    '
    sParamLastNumber = paramAppli.getParams("P_SYSTEM_LASTNUM")
    If sParamLastNumber = "" Then
        sParamLastNumber = "01"
    Else
        sParamLastNumber = Format(CLng(sParamLastNumber), "00")
    End If
    '
    ' Génération du fichier Final
    '
    sParamNameFileFinal = Replace(sParamNameFile, "AAAAMMJJ_HHMMSS", sParamLastDate)
    sParamNameFileFinal = Replace(sParamNameFileFinal, "XX", sParamLastNumber)
    '
    ' Ecrire dans l'onglet paramètre la date et le numéro d'ordre
    '
    numLigne = paramAppli.getNumberLineOf("P_SYSTEM_LASTDATE")
    Worksheets("PARAMETRES").Cells(numLigne, 4).Value = sParamLastDate
      
    numLigne = paramAppli.getNumberLineOf("P_SYSTEM_LASTNUM")
    Worksheets("PARAMETRES").Cells(numLigne, 4).Value = sParamLastNumber
      
    Call argLogAppli.writeInfo("Fin : Sauvegarde le fichier")
    Call argLogAppli.displayMessage
    '
    ' Ecrire dans le menu général l'étape et la date
    '
    sSheetName = "MODE OP"
   
    numLigne = 6
    While Worksheets(sSheetName).Cells(numLigne, 11).Value <> ""
        numLigne = numLigne + 1
    Wend
    Worksheets(sSheetName).Cells(numLigne, 11).Value = argStepName
    Worksheets(sSheetName).Cells(numLigne, 12).Value = Format(Now(), "dd/MM/yyyy-hh:mm:ss")
    '
    ' Charge le fichier des libellés métiers
    '
    Call objIndex.loadFromModele("process_metier.csv", "refMetier")
    Set objLibMetier = objIndex.getIndex("refMetier")
    If Not objLibMetier Is Nothing Then
        If objLibMetier.Exists(argStepName) Then
            Worksheets(sSheetName).Cells(numLigne, 13).Value = objLibMetier(argStepName)
        End If
    End If
      
    '
    ' Sauver le fichier
    '
    Application.CalculateBeforeSave = False
    ActiveWorkbook.SaveAs filename:=ActiveWorkbook.Path + "\" + sParamNameFileFinal, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
      
End Sub


Public Sub writeStepNameOnly(ByVal argStepName As String)
        
    Dim sSheetName As String
    Dim numLigne As Long
      
    Dim objIndex                As New ManageIndex
    Dim objLibMetier            As New Scripting.Dictionary
            
    ' Ecrire dans le menu général l'étape et la date
    '
    sSheetName = "MODE OP"
   
    numLigne = 6
    While Worksheets(sSheetName).Cells(numLigne, 11).Value <> ""
        numLigne = numLigne + 1
    Wend
    Worksheets(sSheetName).Cells(numLigne, 11).Value = argStepName
    Worksheets(sSheetName).Cells(numLigne, 12).Value = Format(Now(), "dd/MM/yyyy-hh:mm:ss")
    '
    ' Charge le fichier des libellés métiers
    '
    Call objIndex.loadFromModele("process_metier.csv", "refMetier")
    Set objLibMetier = objIndex.getIndex("refMetier")
    If Not objLibMetier Is Nothing Then
        If objLibMetier.Exists(argStepName) Then
            Worksheets(sSheetName).Cells(numLigne, 13).Value = objLibMetier(argStepName)
        End If
    End If
      
      
End Sub

' ----------------------------------------------------------------
' Procedure Name: feuilleExiste
' Purpose: Vérifie l'existance d'un onglet
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFeuilleAVerifier (String): Nom de l'onglet
' Return Type: Boolean
' Author: XMXU139
' Date: 15/12/2019
' ----------------------------------------------------------------
Public Function feuilleExiste(sFeuilleAVerifier As String) As Boolean

    On Error GoTo erreurGenerale
    
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim objFeuille            As Worksheet
        
    feuilleExiste = False
    
    For Each objFeuille In Worksheets
        
        If objFeuille.name = sFeuilleAVerifier Then
            feuilleExiste = True
            Exit For
        End If
        
    Next objFeuille
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
    
    Err.Raise lErr, "ManageOnglet.feuilleExiste(" & sFeuilleAVerifier & ")", sDescription

End Function


' ----------------------------------------------------------------
' Procedure Name: openFileDialogSelectFile
' Purpose: Ouverture d'une fenêtre de sélection de fichier
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sPath (String): Répertoire de base pour la sélection
' Parameter sTitle (String): Titre de la fenêtre
' Return Type: String
' Author: XMXU139
' Date: 15/12/2019
' ----------------------------------------------------------------
Public Function openFileDialogSelectFile(ByVal sPath As String, ByVal sTitle As String) As String
    
    On Error GoTo erreurGenerale
     
    '   Descriptif extensions fichiers
    With Application.FileDialog(msoFileDialogFilePicker)
        'Définit un titre pour la boîte de dialogue
        .Title = sTitle
        'Autorise la multi-sélection
        .AllowMultiSelect = False
 
        'Définit un nom de fichier par défaut
        '.InitialFileName = "Nom document.doc"
 
        'Efface les filtres existants.
        .Filters.Clear
        'Définit une liste de filtres pour le champ "Type de fichiers".
        .Filters.Add "Classeurs Excel", "*.xls; *.xlsx; *.xlsm"
        'La valeur 1 permet de spécifier que ce filtre est positionner en premier dans la liste de choix.
        .Filters.Add "Fichiers Textes", "*.csv; *.text", 1
        'Définit le filtre qui s'affiche par défaut dans le champ "Type de fichiers".
        .FilterIndex = 1
 
        'Indique le type d'affichage dans la boîte de dialogue (exemple visualisation des propriétés)
        .InitialView = msoFileDialogViewProperties
        .InitialFileName = sPath
        'Affiche la boîte de dialogue
        .Show
 
        openFileDialogSelectFile = .SelectedItems(1)
 
    End With
 
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

        
    On Error GoTo 0
    
    openFileDialogSelectFile = ""

End Function

