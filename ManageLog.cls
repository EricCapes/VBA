VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Help: https://app.gitbook.com/@eric-ventalon/s/vba-framework/
Private logSheetName  As String
Private numLigne As Long
Private listeDesMessages As Collection

Private Function getLastLine() As Long

    getLastLine = 2
    
    While Trim(Worksheets(logSheetName).Cells(getLastLine, 1).Value) <> ""
        getLastLine = getLastLine + 1
    Wend
    
End Function

Public Sub InitManageLog()
    Set listeDesMessages = New Collection
    logSheetName = "JOURNALISATION"
    If Not isExistSheetName() Then
       Sheets.Add(After:=Worksheets(1)).name = logSheetName
       numLigne = 2
    Else
       numLigne = getLastLine()
    End If
    
End Sub
Private Function isExistSheetName() As Boolean

On Error GoTo SiErreur
    Dim Feuille As Worksheet
        
    isExistSheetName = False
    For Each Feuille In Worksheets
        If Feuille.name = logSheetName Then
             isExistSheetName = True
             Exit Function
         End If
    Next Feuille
    
    Exit Function

SiErreur:

    isExistSheetName = False
    
End Function


Public Sub writeInfo(argMessage As String)
   
   listeDesMessages.Add Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss") + "@INFO@" + argMessage
   

End Sub

Public Sub writeError(argMessage As String)
   
  listeDesMessages.Add Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss") + "@ERREUR@" + argMessage

End Sub


Public Sub writeDebug(argMessage As String)
   
  listeDesMessages.Add Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss") + "@DEBUG@" + argMessage
  
End Sub

Public Sub displayMessage()
    Dim WrdArray() As String
    
    numLigne = getLastLine()
    For Each unMessage In listeDesMessages
    
        WrdArray = Split(unMessage, "@")
        Worksheets(logSheetName).Cells(numLigne, 1).Value = WrdArray(0)
        Worksheets(logSheetName).Cells(numLigne, 2).Value = WrdArray(1)
        Worksheets(logSheetName).Cells(numLigne, 3).Value = WrdArray(2)
        
        numLigne = numLigne + 1
    Next
    
    Set listeDesMessages = Nothing
    Set listeDesMessages = New Collection
    
End Sub

