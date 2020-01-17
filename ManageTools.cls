VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Function pre_traitement(chaine As String)
    acc = Array("é", "è", "ê", "à", "ï")
    repl = Array("e", "e", "e", "a", "e", "i")
    For i = 0 To UBound(acc)
        chaine = Replace(chaine, acc(i), repl(i))
    Next
    chaine = UCase(chaine)
    pre_traitement = chaine
End Function


Public Function ressemblance(ByVal s1 As String, ByVal s2 As String) As Single
'Calcul la similarité (de [0 à 1]) entre deux chaines d'après l'algorithme de Damerau-Levenshtein
'références : http://en.wikipedia.org/wiki/Damerau%E2%80%93Levenshtein_distance
'             http://mwh.geek.nz/2009/04/26/python-damerau-levenshtein-distance/
'             http://www-igm.univ-mlv.fr/~lecroq/seqcomp/node2.html
'Remarques  : Préparer les chaines car les comparaisons sont binaires : UCase(), Trim(),...
'Philben v1.0 - Free to Use
    Const cFacteur As Long = &H100&, cMaxLen As Long = 256&   'Longueur maxi autorisée des chaines analysées
    Dim l1 As Long, l2 As Long, c1 As Long, c2 As Long
    Dim r() As Integer, rp() As Integer, rpp() As Integer, i As Integer, j As Integer
    Dim c As Integer, x As Integer, y As Integer, z As Integer, f1 As Integer, f2 As Integer
    Dim dls As Single, ac1() As Byte, ac2() As Byte
    s1 = pre_traitement(s1)
    s2 = pre_traitement(s2)
    l1 = Len(s1): l2 = Len(s2)
    If l1 > 0 And l1 <= cMaxLen And l2 > 0 And l2 <= cMaxLen Then
        ac1 = s1: ac2 = s2   'conversion des chaines en tableaux de bytes
        'Initialise la ligne précédente (rp) de la matrice
        ReDim rp(0 To l2)
        For i = 0 To l2: rp(i) = i: Next i
        For i = 1 To l1
            'Initialise la ligne courante de la matrice
            ReDim r(0 To l2): r(0) = i
            'Calcul le CharCode du caractère courant de la chaine
            f1 = (i - 1) * 2: c1 = ac1(f1 + 1) * cFacteur + ac1(f1)
            For j = 1 To l2
                f2 = (j - 1) * 2: c2 = ac2(f2 + 1) * cFacteur + ac2(f2)
                c = -(c1 <> c2)   'Cout : True = -1 => c = 1
                'suppression, insertion, substitution
                x = rp(j) + 1: y = r(j - 1) + 1: z = rp(j - 1) + c
                If x < y Then
                    If x < z Then r(j) = x Else r(j) = z
                Else
                    If y < z Then r(j) = y Else r(j) = z
                End If
                'transposition
                If i > 1 And j > 1 And c = 1 Then
                    If c1 = ac2(f2 - 1) * cFacteur + ac2(f2 - 2) And c2 = ac1(f1 - 1) * cFacteur + ac1(f1 - 2) Then
                        If r(j) > rpp(j - 2) + c Then r(j) = rpp(j - 2) + c
                    End If
                End If
            Next j
            'Reculer d'un niveau la ligne précédente (rp) et courante (r)
            rpp = rp: rp = r
        Next i
        'Calcul la similarité via la distance entre les chaines r(l2)
        If l1 >= l2 Then dls = 1 - r(l2) / l1 Else dls = 1 - r(l2) / l2
    ElseIf l1 > cMaxLen Or l2 > cMaxLen Then
        dls = -1   'indique un dépassement de longueur de chaine
    ElseIf l1 = 0 And l2 = 0 Then
        dls = 1   'cas particulier
    End If
    ressemblance = dls * 100
End Function

Public Function func_DelAllSpace(str_ParamString As String)
  Dim str_TempString As String
  Dim i As Integer
 
  str_TempString = LTrim(str_ParamString)
  str_TempString = RTrim(str_TempString)
 
  i = InStr(1, str_TempString, "  ")
 
  While i <> 0
    str_TempString = Replace(str_TempString, "  ", " ")
    i = InStr(1, str_TempString, "  ")
    DoEvents
  Wend
 
  func_DelAllSpace = str_TempString
 
End Function

