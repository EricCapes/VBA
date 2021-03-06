VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: ManageArray
' help: https://app.gitbook.com/@eric-ventalon/s/vba-framework/
' Kind: Module de gestion des tableau m�moire
' Author: Eric VENTALON
' Date: 10/07/2019
' ------------------------------------------------------
Option Explicit
'
' Liens
'
Private m_objClasseMetier           As Object
Private m_objWorkBook               As Workbook
Private m_objHeaderArray            As New Scripting.Dictionary  ' Dictionnaire des nom des colonnes
'
' Attributs
'
Private m_iIndexSheet               As Integer                   ' Index de la feuille
Private m_lEnteteCol                As Long                      ' Position des ent�tes en colonne
Private m_lEnteteRow                As Long                      ' Position des ent�tes en ligne
Private m_sSheetName                As String
Private m_sColCount                 As String
Private m_sSeparateur               As String
'
' La feuille en m�moire
'
Private m_vFileToArray()            As Variant                   ' Le tableau m�moire des feuilles il commence en ligne 1 et non 0
Private m_lLigneMax                 As Long
'
' Gestion des index
'
Private m_objIndex                  As New ManageIndex

' ----------------------------------------------------------------
' Procedure Name: setValue
' Purpose: Affecte une valeur dans le tableau
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter lPosLigne (Long): Position de la ligne
' Parameter lPosColonne (Long): Position de la colonne
' Parameter vValue (Variant): Une valeur
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Sub setValue(ByVal lPosLigne As Long, ByVal lPosColonne As Long, ByVal vValue As Variant)
    m_vFileToArray(lPosLigne, lPosColonne) = vValue
End Sub


' ----------------------------------------------------------------
' Procedure Name: getValue
' Purpose: Retourne une valeur du tableau
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lPosLigne (Long): Position en ligne
' Parameter lPosColonne (Long): Position en colonne
' Return Type: Variant
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getValue(ByVal lPosLigne As Long, ByVal lPosColonne As Long) As Variant
    getValue = m_vFileToArray(lPosLigne, lPosColonne)
End Function

' ----------------------------------------------------------------
' Procedure Name: addLine
' Purpose: Ajoute des lignes au tableau
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter lNbrLigneEnPlus (Long): Nombre de ligne en plus
' Author: XMXU139
' Date: 15/07/2019
' Maintenance: 05/11/2019 - ERV - Cas de diminution du tableau
'
' ----------------------------------------------------------------
Public Sub addLine(Optional ByVal lNbrLigneEnPlus As Long = 1)
    Dim lMaxCol As Long
    Dim lFirst As Long
    Dim lLast As Long
    
    Dim vOuterArray() As Variant
    
    m_lLigneMax = UBound(m_vFileToArray, 1)
    lMaxCol = UBound(m_vFileToArray, 2)
    
    ' 05/11/2019 - ERV - Cas de diminution du tableau
    ReDim vOuterArray(1 To m_lLigneMax + lNbrLigneEnPlus, 1 To lMaxCol)
    
    ' 05/11/2019 - ERV - Cas de diminution du tableau
    If UBound(vOuterArray, 1) < m_lLigneMax Then
       m_lLigneMax = UBound(vOuterArray, 1)
    End If
    
    
    For lFirst = 1 To m_lLigneMax
        For lLast = 1 To lMaxCol
            vOuterArray(lFirst, lLast) = m_vFileToArray(lFirst, lLast)
        Next lLast
    Next lFirst
    
    m_vFileToArray = vOuterArray
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: getWorkbook
' Purpose: Retourne le workbook associ� au tableau
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Workbook
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getWorkbook() As Workbook
    Set getWorkbook = m_objWorkBook
End Function

' ----------------------------------------------------------------
' Procedure Name: getSheetIndex
' Purpose: Retourne le nom de la feuille associ� au tableau
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Integer
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getSheetIndex() As Integer
    getSheetIndex = m_iIndexSheet
End Function

' ----------------------------------------------------------------
' Procedure Name: getPosColumnHeader
' Purpose: Position de d�part en colonne des ent�tes
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Boolean
' Author: XMXU139
' Date: 15/07/2019
' Maintenance: ERV - 15/12/2019 - 01.00 - Cr�ation
'
' ----------------------------------------------------------------
Public Function getPosColumnHeader() As Long
    getPosColumnHeader = m_lEnteteCol
End Function

' ----------------------------------------------------------------
' Procedure Name: getPosRowHeader
' Purpose: Position de d�part en ligne des ent�tes
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Boolean
' Author: XMXU139
' Date: 15/07/2019
' Maintenance: ERV - 15/12/2019 - 01.00 - Cr�ation
'
' ----------------------------------------------------------------
Public Function getPosRowHeader() As Long
    getPosRowHeader = m_lEnteteRow
End Function

' ----------------------------------------------------------------
' Procedure Name: getHeader
' Purpose: Retourne le dictionnaire des ent�tes
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Dictionary
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getHeader() As Scripting.Dictionary
    Set getHeader = m_objHeaderArray
End Function


' ----------------------------------------------------------------
' Procedure Name: SetHeader
' Purpose: Initialise une collection d'ent�te
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter r_objHeader (Dictionary): Les ent�tes
' Author: Eric  VENTALON
' Date: 04/11/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Sub SetHeader(ByRef r_objHeader As Scripting.Dictionary)
    Set m_objHeaderArray = r_objHeader
End Sub

' ----------------------------------------------------------------
' Procedure Name: getNameColumnCount
' Purpose: Retourne le nom de la colonne qui sert de compteur de ligne
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: String
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getNameColumnCount() As String
    getNameColumnCount = m_sColCount
End Function

' ----------------------------------------------------------------
' Procedure Name: getColumnMax
' Purpose: Retourne le nombre de colonne du tableau
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Long
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getColumnMax() As Long
    getColumnMax = UBound(m_vFileToArray, 2)
End Function

' ----------------------------------------------------------------
' Procedure Name: getLineMax
' Purpose: Retourne le nombre de ligne du tableau
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Long
' Author: XMXU139
' Date: 15/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getLineMax() As Long
    getLineMax = UBound(m_vFileToArray, 1)
End Function
' ----------------------------------------------------------------
' Procedure Name: sheetToArray
' Purpose: Place en m�moire une feuille
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter r_objWorkBook (Workbook): R�f�rence du workbook
' Parameter iIndexSheet (Integer): Index de la feuille
' Parameter sZone (String): Zone � renseigner
' Author: Eric VENTALON
' Date: 16/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Sub sheetToArray(ByRef r_objWorkBook As Workbook, ByVal iIndexSheet As Integer, ByVal sZone As String)
    m_vFileToArray = r_objWorkBook.Sheets(iIndexSheet).Range(sZone).Value
End Sub

' ----------------------------------------------------------------
' Procedure Name: setSheet
' Purpose: Pr�pare une feuille pour la mise en place en m�moire
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter r_objWorkBook (Workbook): R�f�rence du workbook
' Parameter sSheetName (String): Nom de la feuille
' Parameter iIndexSheet (Integer): Index de la feuille
' Parameter sColCount (String): Nom de la colonne pour le comptage des lignes. Exemple A
' Parameter r_objClasseMetier (Object): La r�f�rence de la classe m�tier
' Parameter bHearderColumn (Boolean): true : Ent�te, false : pas d'ent�te
' Parameter sSeparateur (String): le s�parateur pour transformer une ligne entiere en colonne
' Parameter r_objHeaderArray (Dictionary): Le dictionnaire des ent�tes
' Author: Eric VENTALON
' Date: 16/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Sub setSheet(ByRef r_objWorkBook As Workbook, _
    ByVal sSheetName As String, _
    ByVal iIndexSheet As Integer, _
    ByVal sColCount As String, _
    ByRef r_objClasseMetier As Object, _
    ByVal lHearderColumn As Long, _
    ByVal sSeparateur As String, _
    ByRef r_objHeaderArray As Scripting.Dictionary, _
    ByVal lHearderRow As Long)
    
    On Error GoTo erreurGenerale
       
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
    '
    ' Affecte les attributs
    '
    m_lEnteteCol = lHearderColumn   ' Position en colonne du tableau
    m_lEnteteRow = lHearderRow
    m_iIndexSheet = iIndexSheet        ' Index de la feuille
    m_sSheetName = sSheetName          ' Nom de la feuille des donn�es
    m_sColCount = sColCount            ' La colonne pour compter les lignes
    m_sSeparateur = sSeparateur        ' Le s�parateur
    '
    ' Attache les liens
    '
    Set m_objHeaderArray = r_objHeaderArray
    Set m_objClasseMetier = r_objClasseMetier
    Set m_objWorkBook = r_objWorkBook
    
    On Error GoTo 0
    Exit Sub
   
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageArray.setSheet()", sDescription
 
End Sub


' ----------------------------------------------------------------
' Procedure Name: arrayToSheet
' Purpose: Place une zone de la feuille dans un tableau et ferme ou pas le document
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter r_objWorkBook (Workbook): La r�f�rence du workbook
' Parameter sSheetName (String): Le nom de la feuille
' Parameter r_objOnglet (ManageOnglet): Un objet de type ManageOnglet
' Parameter sZone (String): La zone � definir ou vide. Exemple "" ou "A3:R89" ou "A4:" c'est la taille du tableau
' Author: Eric VENTALON
' Date: 16/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Sub arrayToSheet(ByRef r_objWorkBook As Workbook, ByVal sSheetName As String, ByRef r_objOnglet As ManageOnglet, Optional ByVal sZone As String = "")
    
    On Error GoTo erreurGenerale
    
    Dim lErr                As Long
    Dim lLigneMax           As Long
    Dim lLigneCpt           As Long
    Dim sDescription        As String
    Dim sSource             As String
    
    '
    ' Calcule le nombre de ligne maxi + 1 avec les ent�tes. Si pas de ligne => 1 + 1 = 2
    lLigneMax = UBound(m_vFileToArray, 1)
    '
    ' La zone � m�moriser
    If sZone = "" Then
        If m_lEnteteCol = False Then
            sZone = "A1:" + r_objOnglet.nomDeColonne(getColumnMax) & lLigneMax
        Else
            sZone = "A2:" + r_objOnglet.nomDeColonne(m_objHeaderArray.Count) & lLigneMax + 1
        End If
    ElseIf Right(sZone, 1) = ":" And m_lEnteteCol = 1 Then
        '
        ' Extraire le chiffre de A45: -> 45
        sSource = "0"
        For lLigneCpt = 1 To Len(sZone)
            If IsNumeric(Mid(sZone, lLigneCpt, 1)) Then
                sSource = sSource + Mid(sZone, lLigneCpt, 1)
            End If
        Next lLigneCpt
        sZone = sZone + r_objOnglet.nomDeColonne(getColumnMax) & lLigneMax + CLng(sSource) - 1
    End If
    
    r_objWorkBook.Sheets(sSheetName).Range(sZone) = m_vFileToArray
    
    On Error GoTo 0
    Exit Sub
   
erreurGenerale:
    
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageArray.arrayToSheet()", sDescription
End Sub

' ----------------------------------------------------------------
' Procedure Name: dropArray
' Purpose: Suppression du tableau,  des index et des ent�tes
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Eric VENTALON
' Date: 16/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Sub dropArray()
    Call m_objHeaderArray.RemoveAll
    Erase m_vFileToArray
    Call m_objIndex.dropAllIndex
End Sub


' ----------------------------------------------------------------------------
' Procedure Name: SaveIntoFile
' Purpose: Sauvegarde le tableau dans un fichier csv avec s�parateur
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sFileName (String): Nom complet du fichier
' Author: Eric VENTALON
' Date: 23/07/2019
' Maintenance: 05/11/2019 - ERV - Suppression des 2 caract�res de fin de ligne
'
' ----------------------------------------------------------------------------
Public Sub SaveIntoFile(ByVal sFileName As String)

    On Error GoTo erreurGenerale

    Dim objFso           As Object
    Dim objFile          As Object
    Dim vKey             As Variant
    Dim sFullText        As String
    
    Dim lErr             As Long
    Dim sDescription     As String
    Dim sSource          As String
    Dim iEntete          As Integer
    
    Dim lLast, lMaxCol   As Long
    Dim lFirst, lMaxLig  As Long
    '
    ' Pr�paration du fichier
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.CreateTextFile(sFileName, True)
    '
    ' Enregistrer l'ent�te
    sFullText = ""
    iEntete = 0
    If m_lEnteteCol Then
        For Each vKey In m_objHeaderArray.Keys()
            sFullText = sFullText + CStr(vKey) + ";"
        Next vKey
        sFullText = Mid(sFullText, 1, Len(sFullText) - 2)
        sFullText = sFullText & vbCrLf
        objFile.Write sFullText
        iEntete = 1
    End If
    '
    ' Enregistre  tout le tableau
    lMaxLig = getLineMax
    lMaxCol = getColumnMax
    
    For lFirst = 1 To lMaxLig
        sFullText = ""
        For lLast = 1 To lMaxCol
            sFullText = sFullText + CStr(m_vFileToArray(lFirst, lLast)) + ";"
        Next lLast
        ' Suppression de cette ligne
        ' sFullText = Mid(sFullText, 1, Len(sFullText) - 2)
        sFullText = sFullText & vbCrLf
        objFile.Write sFullText
    Next lFirst
        
    Call objFile.Close

    Set objFile = Nothing
    Set objFso = Nothing
    
    On Error GoTo 0
    Exit Sub
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + "ManageArray.SaveIntoFile()", sDescription

End Sub

' ----------------------------------------------------------------
' Procedure Name: loadFromFile
' Purpose: Charger un fichier csv - TODO : TOUT REVOIR POUR LES ENTETES
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sFileName (String): Nom complet du fichier
' Parameter bClearArray (Boolean): True = Efface le tableau, false = Ajoute � la suite
' Parameter lIncrement (Long): Nombre de ligne en plus dans le tableau
' Author: Eric VENTALON
' Date: 23/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Sub loadFromFile(ByVal sFileName As String, Optional ByVal bClearArray As Boolean = True, Optional ByVal lIncrement As Long = 15000)
    
    On Error GoTo erreurGenerale
    
    Dim objFso          As Object
    Dim objFile         As Object
    
    Dim sFullText       As String
    Dim sTabtext()      As String
    
    Dim lCptFile        As Long
    Dim lLigne          As Long
    
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
        
    If m_objHeaderArray.Count = 0 Then
        Err.Raise 800, "ManageArray.loadFromFile()", "L'ent�te est obligatoire"
    End If
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.OpenTextFile(sFileName)
        
    m_lEnteteCol = True
    m_sSeparateur = ";"
    
    '
    ' Compter le nombre de ligne pour cr�er le tableau
    '
    lCptFile = 0
    Do While objFile.AtEndOfStream <> True
        objFile.readLine
        lCptFile = lCptFile + 1
    Loop
    objFile.Close
    
    If bClearArray = True Then
        Erase m_vFileToArray
        ReDim m_vFileToArray(1 To lCptFile, 1 To m_objHeaderArray.Count) As Variant
        m_lLigneMax = lIncrement
    End If
    '
    ' Ouverture du fichier
    '
    lCptFile = 1
    Set objFile = objFso.OpenTextFile(sFileName)
    Do While objFile.AtEndOfStream <> True
        
        sFullText = objFile.readLine
        sTabtext = Split(sFullText, m_sSeparateur)
        For lLigne = 0 To UBound(sTabtext) - 1
            m_vFileToArray(lCptFile, lLigne + 1) = sTabtext(lLigne)
        Next lLigne
        
        lCptFile = lCptFile + 1
        
        If lCptFile > m_lLigneMax Then
            Call addLine(lIncrement)
        End If
        
    Loop
    Call objFile.Close

    Set objFile = Nothing
    Set objFso = Nothing
    
    On Error GoTo 0
    Exit Sub
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Call Err.Raise(lErr, sSource, sDescription)

End Sub

' ----------------------------------------------------------------
' Procedure Name: createIndexUnique
' Purpose: Cr�ation d'un index unique du tableau m�moire
' Procedure Kind: Function
' Procedure Access: Public
' Parameter bClearIndex (Boolean): true : nouvel index, false : r�utilisation de l'index
' Parameter sIndexName (String): Nom de l'index
' Parameter sFunctionSelect (String): Fonction de projection de la cl� primaire de l'index
' Parameter sFunctionValues (String): Fonction d'extraction des valeurs
' Parameter sFunctionWhere (String): Fonction de crit�res
' Return Type: Integer
' Author: Eric VENTALON
' Date: 16/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function createIndexUnique( _
    ByVal bClearIndex As Boolean, _
    ByVal sIndexName As String, _
    ByVal sFunctionSelect As String, _
    ByVal sFunctionValues As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
    
    On Error GoTo erreurGenerale
       
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
            
    createIndexUnique = -1
    
    createIndexUnique = m_objIndex.createIndexUniqueInArray(m_vFileToArray, m_objClasseMetier, m_objHeaderArray, bClearIndex, sIndexName, sFunctionSelect, sFunctionValues, sFunctionWhere)
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, Err.Source + " - ManageArray.createIndexUnique()", sDescription
    
End Function

' ----------------------------------------------------------------
' Procedure Name: createIndexSum
' Purpose: Cr�ation d'un index de sommation du tableau
' Procedure Kind: Function
' Procedure Access: Public
' Parameter bClearIndex (Boolean): true : cr�ation, false ajouter � l'existant
' Parameter sIndexName (String): Nom des index s�par� par une virgule
' Parameter sFunctionSelect (String): Fonction de s�lection
' Parameter sFunctionValues (String): Fonction de valeur
' Parameter sFunctionWhere (String): Fonction de crit�re
' Return Type: Integer
' Author: Eric VENTALON
' Date: 16/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function createIndexSum( _
    ByVal bClearIndex As Boolean, _
    ByVal sIndexName As String, _
    ByVal sFunctionSelect As String, _
    ByVal sFunctionValues As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
    
    On Error GoTo erreurGenerale
       
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
        
    createIndexSum = -1
    
    createIndexSum = m_objIndex.createIndexSumInArray(m_vFileToArray, m_objClasseMetier, m_objHeaderArray, bClearIndex, sIndexName, sFunctionSelect, sFunctionValues, sFunctionWhere)
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, Err.Source + " - ManageArray.createIndexSum()", sDescription
End Function

' ----------------------------------------------------------------
' Procedure Name: createIndexMultiValues
' Purpose: Creation d'un index avec cl� unique d�terminant plusieurs valeurs situ� dans des lignes diff�rentes
' Procedure Kind: Function
' Procedure Access: Public
' Parameter bClearIndex (Boolean): True : Cr�ation de l'index, False : Ajout dans l'index existant
' Parameter sIndexName (String): Nom de la feuille � indexer
' Parameter sFunctionSelect (String): Fonction qui d�termine la PK
' Parameter sFunctionValues (String): Fonction qui d�termine les valeurs � extraire pour une ligne
' Parameter sFunctionWhere (String): Fonction qui filtre la ligne en cours
' Return Type: Integer
' Author: XMXU139
' Date: 16/07/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function createIndexMultiValues( _
    ByVal bClearIndex As Boolean, _
    ByVal sIndexName As String, _
    ByVal sFunctionSelect As String, _
    ByVal sFunctionValues As String, _
    Optional ByVal sFunctionWhere As String = "") As Integer
    
    On Error GoTo erreurGenerale
       
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
        
    createIndexMultiValues = -1
    
    createIndexMultiValues = m_objIndex.createIndexMultiValueInArray(m_vFileToArray, m_objClasseMetier, m_objHeaderArray, bClearIndex, sIndexName, sFunctionSelect, sFunctionValues, sFunctionWhere)
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " - ManageArray.createIndexMultiValues()", sDescription
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
' Maintenance:
'
' ----------------------------------------------------------------
Public Function getIndex(sNomIndex As String) As Scripting.Dictionary
    Set getIndex = m_objIndex.getIndex(sNomIndex)
End Function



' ----------------------------------------------------------------
' Procedure Name: delete
' Purpose: Suppression physique des lignes du tableau
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFonctionCritere (String): M�thode de suppression
' Parameter r_objMetier (Object): Classe m�tier
' Return Type: Boolean
' Author: Eric  VENTALON
' Date: 04/11/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function delete(ByVal sFonctionCritere As String, Optional ByRef r_objMetier As Object = Nothing) As Boolean
    
    On Error GoTo erreurGenerale
        
    Dim lMaxCol As Long
    Dim lFirst As Long
    Dim lLast As Long
    
    Dim lErr                  As Long
    Dim sDescription          As String
    Dim sSource               As String
    
    Dim objMetier As Object
    
    Dim vOuterArray() As Variant
        
    delete = False
    
    '
    ' Il faut qu'une classe m�tier soit d�clar�e
    '
    If Not r_objMetier Is Nothing Then
        Set objMetier = r_objMetier
    ElseIf Not m_objClasseMetier Is Nothing Then
        Set objMetier = m_objClasseMetier
    Else
        Err.Raise 800, , "Veuillez affecter une classe m�tier."
    End If
    '
    ' V�rifer que la fonction existe
    '
    sFonctionCritere = Trim(sFonctionCritere)
    If sFonctionCritere = "" Then
        Err.Raise 801, , "Veuillez affecter une fonction de filtre."
    End If
    
    m_lLigneMax = UBound(m_vFileToArray, 1)
    lMaxCol = UBound(m_vFileToArray, 2)
    
    ReDim vOuterArray(1 To m_lLigneMax, 1 To lMaxCol)
    
    For lFirst = 1 To m_lLigneMax
        If CallByName(objMetier, sFonctionCritere, VbMethod, m_vFileToArray, m_objHeaderArray, lFirst) = 1 Then
            For lLast = 1 To lMaxCol
                vOuterArray(lFirst, lLast) = m_vFileToArray(lFirst, lLast)
            Next lLast
        End If
    Next lFirst
    
    m_vFileToArray = vOuterArray
    
    delete = True
    
    On Error GoTo 0
    Exit Function
    
erreurGenerale:

    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Err.Raise lErr, sSource + " : " + Err.Source + " - ManageArray.delete()", sDescription
    
    
End Function

' ----------------------------------------------------------------
' Procedure Name: iterator
' Purpose: Parcourir tout le tableau
' Procedure Kind: Function
' Procedure Access: Public
' Parameter r_objMetier (Object): Classe m�tier
' Parameter sBeforeRead (String): M�thode avant lecture du tableau
' Parameter sReadRow (String): M�thode pendant la lecture du tableau
' Parameter sAfterRead (String): M�thode apr�s lecture du tableau
' Return Type: Integer
' Author: Eric  VENTALON
' Date: 04/11/2019
' Maintenance:
'
' ----------------------------------------------------------------
Public Function iterator( _
    ByRef r_objMetier As Object, _
    ByVal sBeforeRead As String, _
    ByVal sReadRow As String, _
    ByVal sAfterRead As String) As Integer
    
    On Error GoTo erreurGenerale
    
    Dim lResultFct          As Long
    Dim lCptLigne           As Long
    Dim lErr                As Long
    Dim lLigneMax           As Long
    
    Dim sDescription        As String
    Dim sSource             As String
    
    '
    ' 01 - Par d�faut retour p�simiste
    iterator = -1
    '
    ' 02 - Si pas de classe m�tier
    If IsNull(r_objMetier) Then
        Err.Raise 600, , "Veuillez d�clarer une classe m�tier pour les crit�res, extractions, fonctions."
    End If
    If r_objMetier Is Nothing Then
        Err.Raise 600, , "Veuillez d�clarer une classe metier pour les crit�res, extractions, fonctions."
    End If
    '
    ' 03 - Ev�nement avant lecture du tableau
    ' Retour code 2 - Arr�t de l'iterateur
    If sBeforeRead <> "" Then
        lResultFct = CallByName(r_objMetier, sBeforeRead, VbMethod, m_objHeaderArray)
        If lResultFct = 2 Then
            Exit Function
        End If
    Else
        lResultFct = 0
    End If
    '
    ' 04 - Parcourir tout le tableau
    lCptLigne = 1
    lLigneMax = UBound(m_vFileToArray, 1)
    Do While lCptLigne <= lLigneMax And lResultFct <> 2
        ' Si en premiere ligne et une ent�te existe alors ...
        ' Retour code 2 - Arr�t de l'iterateur
        If sReadRow <> "" Then
            lResultFct = CallByName(r_objMetier, sReadRow, VbMethod, lCptLigne, m_vFileToArray)
        Else
            lResultFct = 0
        End If
        lCptLigne = lCptLigne + 1
    Loop
    
    If lResultFct = 2 Then
        Exit Function
    End If
    '
    ' 05 - Ev�nement apres lecture du tableau
    ' Retour code 2 - Arr�t de l'iterateur m�me si c'est la fin
    If sAfterRead <> "" Then
        lResultFct = CallByName(r_objMetier, sAfterRead, VbMethod)
        If lResultFct = 2 Then
            Exit Function
        End If
    Else
        iterator = 0
    End If
   
    On Error GoTo 0
    Exit Function
   
erreurGenerale:
    
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
        
    On Error GoTo 0
   
    Call Err.Raise(lErr, sSource + " : ManageArray.iterator()", sDescription)
   
End Function




