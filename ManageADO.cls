VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: ManageADO
' Kind: Module de classe
' Purpose: Gestion des connexions
' Author: XMXU139
' Date: 13/09/2019
' ------------------------------------------------------
Option Explicit

Private m_objDb As New ADODB.Connection
Private m_objRecordset As ADODB.Recordset
Private m_objSchema As New Scripting.Dictionary

' ----------------------------------------------------------------
' Procedure Name: openDatabaseAccess2007
' Purpose: Ouverture de la base de données
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFullPath (String): Chemin absolue du fichier
' Return Type: Long
' Author: XMXU139
' Date: 18/09/2019
' ----------------------------------------------------------------
Public Function openDatabaseAccess2007(ByVal sFullPath As String) As Long

    On Error GoTo erreurSQL
   
    openDatabaseAccess2007 = -1

    m_objDb.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & sFullPath

    openDatabaseAccess2007 = 0
   
    On Error GoTo 0
    Exit Function
   
erreurSQL:

    Dim lErr As Long
    Dim sDescription As String
    Dim sSource As String
   
    If Not (m_objDb Is Nothing) Then
        If m_objDb.State = adStateOpen Then
            Call closeDatabase
        End If
    End If
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.openDatabaseAccess2007()", sDescription

End Function

' ----------------------------------------------------------------
' Procedure Name: openXlsFile
' Purpose: Ouverture d'un fichier XLS
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFullPath (String): Nom complet du fichier XLS
' Return Type: Long
' Author: XMXU139
' Date: 13/09/2019
' ----------------------------------------------------------------
Public Function openXlsFile(ByVal sFullPath As String) As Long

    On Error GoTo erreurSQL
   
    Dim lErr            As Long
    
    Dim sDescription    As String
    Dim sSource         As String
    Dim sDbName         As String
   
    openXlsFile = -1
   
    If isFileOpen(sFullPath) Then
        MsgBox "Impossible d'analyser les données. Le fichier XLS contenant PHARE et GESTPRET doit être fermé.", vbOKOnly Or vbCritical, "COREM ANALYSE"
        Exit Function
    End If

    sDbName = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sFullPath & ";Extended Properties=""Excel 12.0;HDR=YES;"""
    m_objDb.Open sDbName

    openXlsFile = 0
   
    On Error GoTo 0
    Exit Function
   
erreurSQL:
   
    If Not (m_objDb Is Nothing) Then
        If m_objDb.State = adStateOpen Then
            Call closeDatabase
        End If
    End If
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.openXlsFile()", sDescription

End Function

' ----------------------------------------------------------------
' Procedure Name: closeDatabase
' Purpose: Termine la connexion oledb
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 13/09/2019
' ----------------------------------------------------------------
Public Sub closeDatabase()
    On Error Resume Next
    
    m_objDb.Close
End Sub

' ----------------------------------------------------------------
' Procedure Name: execSqlIntoSheet
' Purpose: Execute une requête SQL  (type select) et retourne ses valeurs
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sNomFeuille (String): Nom de la feuille de données
' Parameter sPosition (String): Position dans la feuille de données. Exemple A3
' Parameter sSql (String): Requête SQL complète
' Author: XMXU139
' Date: 13/09/2019
' ----------------------------------------------------------------
Public Sub execSqlIntoSheet(ByVal sNomFeuille As String, ByVal sPosition As String, ByVal sSql As String, Optional bDisplayHeader As Boolean = False, Optional lPosHeaderCol As Long = 0, Optional lPosHeaderRow As Long = 0)

    On Error GoTo erreurSQL
    
    Dim iCols               As Integer
    Dim lErr                As Long
    Dim sDescription        As String
    Dim sSource             As String
    
    Dim objRecord           As ADODB.Recordset
    
    Set objRecord = m_objDb.Execute(sSql)
    
    If bDisplayHeader Then
        If lPosHeaderCol = 0 Then
            Call Err.Raise(800, "execSqlIntoSheet", "Pas de position en colonne pour les entêtes.")
        End If
        If lPosHeaderRow = 0 Then
            Call Err.Raise(800, "execSqlIntoSheet", "Pas de position en ligne pour les entêtes.")
        End If
        For iCols = 0 To objRecord.Fields.Count - 1
            Sheets(sNomFeuille).Cells(lPosHeaderRow, iCols + lPosHeaderCol).Value = objRecord.Fields(iCols).name
        Next
    End If
    
    Sheets(sNomFeuille).Range(sPosition).CopyFromRecordset objRecord
    
    Call objRecord.Close
    
    On Error GoTo 0
    Exit Sub
    
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    If Not (objRecord Is Nothing) Then
        If objRecord.State = adStateOpen Then
            objRecord.Close
        End If
    End If
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.execSQL(" + sSql + ")", sDescription

    
End Sub
      
' ----------------------------------------------------------------
' Procedure Name: execSqlCmd
' Purpose: Exécude une commande SQL
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sSql (String): la requête SQL
' Author: XMXU139
' Date: 18/09/2019
' ----------------------------------------------------------------
Public Sub execSqlCmd(ByVal sSql As String)
    On Error GoTo erreurSQL
    
    Dim lErr                As Long
    Dim sDescription        As String
    Dim sSource             As String
    
    Dim objCom              As New ADODB.Command
    
    objCom.ActiveConnection = m_objDb
    objCom.CommandTimeout = 5000
    objCom.CommandType = adCmdText
    
    objCom.CommandText = sSql
    objCom.Execute
            
    On Error GoTo 0
    Exit Sub
    
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.execSqlCmd()", sDescription
End Sub

' ----------------------------------------------------------------
' Procedure Name: execProcStock
' Purpose: Execution procédure
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sProcStock (String): Nom de la procédure stockée
' Author: XMXU139
' Date: 27/09/2019
' ----------------------------------------------------------------
Public Sub execProcStock(ByVal sProcStock As String)
    On Error GoTo erreurSQL
    
    Dim lErr                As Long
    Dim sDescription        As String
    Dim sSource             As String
    
    Dim objCom              As New ADODB.Command
    
    objCom.ActiveConnection = m_objDb
    objCom.CommandTimeout = 5000
    objCom.CommandType = adCmdStoredProc
    objCom.CommandText = sProcStock
    
    objCom.Execute
    
                
    On Error GoTo 0
    Exit Sub
    
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.execSqlCmd()", sDescription
End Sub
' ----------------------------------------------------------------
' Procedure Name: beginBatch
' Purpose: Ouverture de la transaction
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sTableName (String): Nom de la table à alimenter dans la transaction
' Author: XMXU139
' Date: 18/09/2019
' ----------------------------------------------------------------
Public Sub beginBatch(ByVal sTableName As String)
    On Error GoTo erreurSQL
    
    Dim lErr                As Long
    Dim sDescription        As String
    Dim sSource             As String
        
    '
    Set m_objRecordset = New ADODB.Recordset
    m_objRecordset.CursorLocation = adUseClient
    
    Call m_objRecordset.Open(sTableName, m_objDb, adOpenForwardOnly, adLockBatchOptimistic, adCmdTable)
    
    On Error GoTo 0
    Exit Sub
    
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    m_objRecordset.Close
   
    On Error GoTo 0
      
    Err.Raise Number:=lErr, Source:=sSource + " : ManageADO.beginBatchInsert(" + sTableName + ")", Description:=sDescription

    
End Sub

' ----------------------------------------------------------------
' Procedure Name: insertBatch
' Purpose: Insertion des données dans une transaction
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter r_objFields (Variant()): La liste des champs
' Parameter r_objValues (Variant()): La liste des valeurs
' Author: XMXU139
' Date: 18/09/2019
' ----------------------------------------------------------------
Public Sub insertBatch(ByRef r_objFields() As Variant, ByRef r_objValues() As Variant)
    On Error GoTo erreurSQL
    
    Dim lErr                As Long
    Dim sDescription        As String
    Dim sSource             As String
    
    Call m_objRecordset.AddNew(r_objFields, r_objValues)
    
    On Error GoTo 0
    Exit Sub
    
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    Call m_objRecordset.CancelBatch
    Call m_objRecordset.Close
           
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.insertBatch()", sDescription

    
End Sub

' ----------------------------------------------------------------
' Procedure Name: commitBatch
' Purpose: Validation de toutes les insertions
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 18/09/2019
' ----------------------------------------------------------------
Public Sub commitBatch()
    On Error GoTo erreurSQL
    
    Dim lErr                As Long
    Dim sDescription        As String
    Dim sSource             As String
    
    m_objRecordset.UpdateBatch
    
    On Error GoTo 0
    Exit Sub
    
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    Call m_objRecordset.Close
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.execSQL()", sDescription

    
End Sub

' ----------------------------------------------------------------
' Procedure Name: endBatch
' Purpose: Fin du batch
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 18/09/2019
' ----------------------------------------------------------------
Public Sub endBatch()

    On Error GoTo erreurSQL
    
    Dim lErr                As Long
    Dim sDescription        As String
    Dim sSource             As String
    
    m_objRecordset.Close
    
    On Error GoTo 0
    Exit Sub
    
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
    
    Call m_objRecordset.CancelBatch
    Call m_objRecordset.Close
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.execSQL()", sDescription
End Sub
    


' ----------------------------------------------------------------
' Procedure Name: isFileOpen
' Purpose: Vérifie un fichier est ouvert
' Procedure Kind: Function
' Procedure Access: Private
' Parameter sFileName (String): Nom du fichier complet à tester
' Author: XMXU139
' Date: 13/09/2019
' ----------------------------------------------------------------
Private Function isFileOpen(sFileName As String)
    Dim iFilenum As Integer, iErrnum As Integer

    On Error Resume Next   ' Turn error checking off.
    
    iFilenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open sFileName For Input Lock Read As #iFilenum
    Close iFilenum          ' Close the file.
    iErrnum = Err           ' Save the error number that occurred.
    
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case iErrnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
            isFileOpen = False

            ' Error number for "Permission Denied."
            ' File is already opened by another user.
        Case 70
            isFileOpen = True

            ' Another error occurred.
        Case Else
            Error iErrnum
    End Select

End Function


' ----------------------------------------------------------------
' Procedure Name: getStructureTablesColonnes
' Purpose: Recherche toutes les tables avec leurs colonnes
' Procedure Kind: Sub
' Procedure Access: Public
' Author: XMXU139
' Date: 18/09/2019
' ----------------------------------------------------------------
Public Sub getStructureTablesColonnes()
    
    Dim objTable        As ADODB.Recordset
    Dim objField        As ADODB.Recordset
    Dim objType         As New Collection
    Dim objStruct       As Scripting.Dictionary
        
   
    Dim lErr            As Long
    Dim sDescription    As String
    Dim sSource         As String
    
    Dim sNomTable       As String
    
       
    objType.Add Key:="0", Item:="adEmpty"
    objType.Add Key:="2", Item:="adSmallInt"
    objType.Add Key:="3", Item:="adInteger"
    objType.Add Key:="4", Item:="adSingle"
    objType.Add Key:="5", Item:="adDouble"
    objType.Add Key:="6", Item:="adCurrency"
    objType.Add Key:="7", Item:="adDate"
    objType.Add Key:="8", Item:="adBSTR"
    objType.Add Key:="9", Item:="adIDispatch"
    objType.Add Key:="10", Item:="adError"
    objType.Add Key:="11", Item:="adBoolean"
    objType.Add Key:="12", Item:="adVariant"
    objType.Add Key:="13", Item:="adIUnknown"
    objType.Add Key:="14", Item:="adDecimal"
    objType.Add Key:="16", Item:="adTinyInt"
    objType.Add Key:="17", Item:="adUnsignedTinyInt"
    objType.Add Key:="18", Item:="adUnsignedSmallInt"
    objType.Add Key:="19", Item:="adUnsignedInt"
    objType.Add Key:="20", Item:="adBigInt"
    objType.Add Key:="21", Item:="adUnsignedBigInt"
    
    objType.Add Key:="64", Item:="adFileTime"
    objType.Add Key:="72", Item:="adGUID"
    objType.Add Key:="128", Item:="adBinary"
    objType.Add Key:="129", Item:="adChar"
    objType.Add Key:="130", Item:="adWChar"
    objType.Add Key:="131", Item:="adNumeric"
    objType.Add Key:="132", Item:="adUserDefined"
    
    objType.Add Key:="133", Item:="adDBDate"
    objType.Add Key:="134", Item:="adDBTime"
    objType.Add Key:="135", Item:="adDBTimeStamp"
    objType.Add Key:="136", Item:="adChapter"
    objType.Add Key:="138", Item:="adPropVariant"
    objType.Add Key:="139", Item:="adVarNumeric"
    objType.Add Key:="200", Item:="adVarChar"
    
    objType.Add Key:="201", Item:="adLongVarChar"
    objType.Add Key:="202", Item:="adVarWChar"
    objType.Add Key:="203", Item:="adLongVarWChar"
    objType.Add Key:="204", Item:="adVarBinary"
    objType.Add Key:="205", Item:="adLongVarBinary"
    '
    ' Liste des tables et de leurs colonnes avec leur type, celle par défault et la colonne facultative
    '
    Set objTable = m_objDb.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    Do While Not objTable.EOF
        If objTable.Fields("TABLE_TYPE") = "TABLE" Then
            sNomTable = objTable.Fields("TABLE_NAME")
            If Not m_objSchema.Exists(sNomTable) Then
                Call m_objSchema.Add(sNomTable, New Scripting.Dictionary)
            End If
            Set objStruct = m_objSchema(sNomTable)
            '
            ' Pour tous les champs
            Set objField = m_objDb.OpenSchema(adSchemaColumns, Array(Empty, Empty, "" & sNomTable, Empty))
            Do While Not objField.EOF
                Call objStruct.Add(objField.Fields("ORDINAL_POSITION").Value, objField.Fields("COLUMN_NAME").Value)
                Call objField.MoveNext
            Loop
            Call objField.Close
        End If
        ' Passe  à la table suivante
        objTable.MoveNext
    Loop
    objTable.Close
             
    On Error GoTo 0
    Exit Sub
    '
    '
erreurSQL:
   
    lErr = Err.Number
    sDescription = Err.Description
    sSource = Err.Source
   
    On Error GoTo 0
      
    Err.Raise lErr, sSource + " : ManageADO.getStructureTablesColonnes()", sDescription
End Sub
'
'
'
Public Function getColumnsFromTable(ByVal sTableName As String) As String
    
    Dim objStruct       As Scripting.Dictionary
    Dim iCpt            As Integer
    Dim sColonne        As String
    
    getColumnsFromTable = ""
    If Not m_objSchema.Exists(sTableName) Then
        Call getStructureTablesColonnes
        If Not m_objSchema.Exists(sTableName) Then
            Exit Function
        End If
    End If
        
    Set objStruct = m_objSchema(sTableName)
    For iCpt = 1 To objStruct.Count
        sColonne = Trim(objStruct(iCpt))
        If InStr(sColonne, " ") > 0 Then
            sColonne = "[" + sColonne + "]"
        End If
        getColumnsFromTable = getColumnsFromTable + sColonne + ","
    Next iCpt
    
    getColumnsFromTable = Left(getColumnsFromTable, Len(getColumnsFromTable) - 1)
    
        
End Function
    
Public Function getColumnsInArrayFromTable(ByVal sTableName As String) As Variant()
    
    Dim objStruct       As Scripting.Dictionary
    Dim iCpt            As Integer
    Dim sColonne        As String
    Dim vTableauCol()   As Variant
    
    If Not m_objSchema.Exists(sTableName) Then
        Call getStructureTablesColonnes
        If Not m_objSchema.Exists(sTableName) Then
            Exit Function
        End If
    End If
        
    Set objStruct = m_objSchema(sTableName)
    
    ReDim vTableauCol(objStruct.Count - 1) As Variant
    
    For iCpt = 1 To objStruct.Count
        sColonne = Trim(objStruct(iCpt))
        'If InStr(sColonne, " ") > 0 Then
        '   sColonne = "[" + sColonne + "]"
        'End If
        vTableauCol(iCpt - 1) = sColonne
    Next iCpt
    
    getColumnsInArrayFromTable = vTableauCol
    
        
End Function
    
    


Public Sub listeStructureTablesColonnes()
    
    Dim rsTable As ADODB.Recordset
    Dim rsField As ADODB.Recordset
    Dim indexType As New Collection
    Dim ligneFind As Long
    Dim ligneFindMem As Long
    Dim bTrouve As Boolean
    Dim valeurTrouve As String
    
    Dim ligne As Long
        
    indexType.Add Key:="0", Item:="adEmpty"
    indexType.Add Key:="2", Item:="adSmallInt"
    indexType.Add Key:="3", Item:="adInteger"
    indexType.Add Key:="4", Item:="adSingle"
    indexType.Add Key:="5", Item:="adDouble"
    indexType.Add Key:="6", Item:="adCurrency"
    indexType.Add Key:="7", Item:="adDate"
    indexType.Add Key:="8", Item:="adBSTR"
    indexType.Add Key:="9", Item:="adIDispatch"
    indexType.Add Key:="10", Item:="adError"
    indexType.Add Key:="11", Item:="adBoolean"
    indexType.Add Key:="12", Item:="adVariant"
    indexType.Add Key:="13", Item:="adIUnknown"
    indexType.Add Key:="14", Item:="adDecimal"
    indexType.Add Key:="16", Item:="adTinyInt"
    indexType.Add Key:="17", Item:="adUnsignedTinyInt"
    indexType.Add Key:="18", Item:="adUnsignedSmallInt"
    indexType.Add Key:="19", Item:="adUnsignedInt"
    indexType.Add Key:="20", Item:="adBigInt"
    indexType.Add Key:="21", Item:="adUnsignedBigInt"
    
    indexType.Add Key:="64", Item:="adFileTime"
    indexType.Add Key:="72", Item:="adGUID"
    indexType.Add Key:="128", Item:="adBinary"
    indexType.Add Key:="129", Item:="adChar"
    indexType.Add Key:="130", Item:="adWChar"
    indexType.Add Key:="131", Item:="adNumeric"
    indexType.Add Key:="132", Item:="adUserDefined"
    
    indexType.Add Key:="133", Item:="adDBDate"
    indexType.Add Key:="134", Item:="adDBTime"
    indexType.Add Key:="135", Item:="adDBTimeStamp"
    indexType.Add Key:="136", Item:="adChapter"
    indexType.Add Key:="138", Item:="adPropVariant"
    indexType.Add Key:="139", Item:="adVarNumeric"
    indexType.Add Key:="200", Item:="adVarChar"
    
    indexType.Add Key:="201", Item:="adLongVarChar"
    indexType.Add Key:="202", Item:="adVarWChar"
    indexType.Add Key:="203", Item:="adLongVarWChar"
    indexType.Add Key:="204", Item:="adVarBinary"
    indexType.Add Key:="205", Item:="adLongVarBinary"
    
    
    Sheets("SGBDR").Select
    Selection.ClearContents
    Sheets("SGBDR").Range("A1").Select
    '
    ' Mise en place des entêtes
    '
    Sheets("SGBDR").Cells(1, 1).Value = "TABLE"
    Sheets("SGBDR").Cells(1, 2).Value = "COLONNE"
    Sheets("SGBDR").Cells(1, 3).Value = "TYPE"
    Sheets("SGBDR").Cells(1, 4).Value = "DEFAULT"
    Sheets("SGBDR").Cells(1, 5).Value = "FACULTATIF"
    Sheets("SGBDR").Cells(1, 6).Value = "PRIMARY KEY"
    Sheets("SGBDR").Cells(1, 7).Value = "ALTERNATIVE KEY"
    Sheets("SGBDR").Cells(1, 8).Value = "FOREIGN KEY"
    '
    ' Liste des tables et de leurs colonnes avec leur type, celle par défault et la colonne facultative
    '
    ligne = 2
    Set rsTable = m_objDb.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    Do While Not rsTable.EOF
        
        Set rsField = m_objDb.OpenSchema(adSchemaColumns, Array(Empty, Empty, "" & rsTable("TABLE_NAME")))
        Do While Not rsField.EOF
            Sheets("SGBDR").Cells(ligne, 1).Value = rsTable("TABLE_NAME")
            Sheets("SGBDR").Cells(ligne, 2).Value = rsField("COLUMN_NAME")
            Sheets("SGBDR").Cells(ligne, 3).Value = indexType(CStr(rsField("DATA_TYPE")))
            Sheets("SGBDR").Cells(ligne, 4).Value = rsField("COLUMN_DEFAULT")
            Sheets("SGBDR").Cells(ligne, 5).Value = rsField("IS_NULLABLE")
            rsField.MoveNext
            ligne = ligne + 1
        Loop
        rsField.Close
        
        rsTable.MoveNext
    Loop
    rsTable.Close
    '
    ' Recherche des clés primaires
    '
    ligne = 2
    Set rsTable = m_objDb.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    Do While Not rsTable.EOF
        Set rsField = m_objDb.OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, "" & rsTable("TABLE_NAME")))
        '
        ' Se positionner sur la table
        ligneFind = 2
        bTrouve = False
        valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 1).Value)
        While bTrouve = False And valeurTrouve <> ""
            If rsTable("TABLE_NAME") = valeurTrouve Then
                bTrouve = True
            Else
                ligneFind = ligneFind + 1
                valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 1).Value)
            End If
        Wend
        If bTrouve = True Then
            '
            '
            ligneFindMem = ligneFind
            Do While Not rsField.EOF
                '
                ' Localise la colonne
                bTrouve = False
                valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 2).Value)
                While bTrouve = False And valeurTrouve <> ""
                    If rsField("COLUMN_NAME") = valeurTrouve Then
                        bTrouve = True
                        Sheets("SGBDR").Cells(ligneFind, 6).Value = "PK"
                    Else
                        ligneFind = ligneFind + 1
                        valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 2).Value)
                    End If
                Wend
                ligneFind = ligneFindMem
                '
                '
                rsField.MoveNext
                ligne = ligne + 1
            Loop
            rsField.Close
        End If
        
        rsTable.MoveNext
    Loop
    rsTable.Close
    '
    ' Recherche des clés étrangéres
    '
    ' adSchemaIndexes
    ' AdSchemaReferentialConstraints
    ' adSchemaConstraintColumnUsage
    ' adSchemaCheckConstraints
    '
    Set rsTable = m_objDb.OpenSchema(adSchemaForeignKeys, Array(Empty, Empty, Empty))
    Do While Not rsTable.EOF
        '
        ' Se positionner sur la table
        ligneFind = 2
        bTrouve = False
        valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 1).Value)
        While bTrouve = False And valeurTrouve <> ""
            If rsTable("FK_TABLE_NAME") = valeurTrouve Then
                bTrouve = True
            Else
                ligneFind = ligneFind + 1
                valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 1).Value)
            End If
        Wend
        If bTrouve = True Then
            '
            ' Localise la colonne
            bTrouve = False
            valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 2).Value)
            While bTrouve = False And valeurTrouve <> ""
                If rsTable("FK_COLUMN_NAME") = valeurTrouve Then
                    bTrouve = True
                    Sheets("SGBDR").Cells(ligneFind, 7).Value = "FK"
                    Sheets("SGBDR").Cells(ligneFind, 8).Value = rsTable("PK_TABLE_NAME")
                    Sheets("SGBDR").Cells(ligneFind, 9).Value = rsTable("PK_COLUMN_NAME")
                Else
                    ligneFind = ligneFind + 1
                    valeurTrouve = Trim(Sheets("SGBDR").Cells(ligneFind, 2).Value)
                End If
            Wend
        End If
        
        rsTable.MoveNext
    Loop
    rsTable.Close

    
    
     
    On Error GoTo 0
    Exit Sub
    '
    '
erreurSQL:

    Dim numErr As Long
    Dim strDescription As String
    Dim strSource As String
   
    numErr = Err.Number
    strDescription = Err.Description
    strSource = Err.Source
   
    On Error GoTo 0
      
    Err.Raise numErr, strSource + "-" + "ManageADO.listeStructureTablesColonnes()", strDescription
End Sub
