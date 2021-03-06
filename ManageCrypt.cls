VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManageCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'D�finition des constantes utiles pour la fonction Crypter et deCrypter (Cl� et nombre d'it�rations de la fonction maximum)
Private Const Clef              As String = "toto"
Private Const NBROTATIONSMAX    As Long = 11

Public Function Crypter(ByVal cha�neACrypter As String)
'---------------------------------------------------------------------------------------
' Procedure &#58; Crypter
' Cr��e le  &#58; lundi 18 juil 2005 18&#58;51
' Auteur    &#58; Maxence HUBICHE
' Site      &#58; http&#58;//mhubiche.developpez.com
' Objet     &#58; Crypter la cha�ne cha�neACrypter en fonction d'une clef et de la m�thode
'               de Vigen�re
'---------------------------------------------------------------------------------------
'
    Dim sLettres    As String
    Dim lCompteur   As Long
    Dim lLongueur   As Long
    Dim lBoucle     As Long
    
    
    'D�finition de la longueur de la cha�ne � crypter et de la cha�ne de r�sultat
    lLongueur = Len(cha�neACrypter)
    sLettres = String(lLongueur, Chr(0))
'Trucs de codage
'Vous ne le saviez peut-�tre pas, mais, lorsque vous �crivez ceci &#58;
'Dim s As String
's = "A"
's = s & "B"
'Vous utilisez �norm�ment de ressources syst�me (enfin ... tout est relatif) car, il y a cr�ation de la variable s pour stocker "A", puis pour ajouter le "B", il va falloir que le syst�me recr�e une variable contenant les 2 caract�res ; alors que si on avait �crit ceci &#58;
's = "xx"
'Mid(s, 1, 1) = "A"
'Mid(s, 2, 1) = "B"
'nous utilisons toujours la m�me variable, et nous la modifions au fur et � mesure qu'on modifie la position d'un caract�re. Cette deuxi�me �criture demande donc moins de traitements au proc�sseur et est donc plus rapide au niveau traitement. C'est donc la solution que j'ai favoris� dans le code de la fonction.


    'Boucler en fonction du nombre de rotations attendues
    For lBoucle = 1 To NBROTATIONSMAX
        'boucler pour chaque caract�re de la cha�ne initiale
        For lCompteur = 1 To lLongueur
            'Remplacer le caract�re de la cha�ne � crypter par le caract�re correspondant �
            '   le reste de
            '       la valeur ascii du caract�re � crypter
            '       plus
            '       la valeur ascii du caract�re correspondant dans la cl�, multipli� par la longueur de la cl�
            '   quand on le divise par 256
            Mid(sLettres, lCompteur, 1) = Chr((asc(Mid(cha�neACrypter, lCompteur, 1)) + _
                    (asc(Mid(Clef, (lCompteur Mod Len(Clef)) + 1, 1)) * lLongueur)) Mod 256)
        'recommencer
        Next
        'r�affecter la cha�ne � crypter par le r�sultat trouv� pour pouvoir recommencer une it�ration
        cha�neACrypter = sLettres
    'Nouvelle it�ration
    Next
    'Renvoyer le r�sultat final
    Crypter = sLettres
End Function

Public Function deCrypter(ByVal cha�neAdeCrypter As String)
'---------------------------------------------------------------------------------------
' Procedure &#58; deCrypter
' Cr��e le  &#58; jeudi 8 septembre 2005
' Auteur    &#58; Muhad'hib (d'apr�s Maxence HUBICHE)
' Site      &#58;
' Objet     &#58; deCrypter la cha�ne cha�neAdeCrypter en fonction d'une clef et de la m�thode
'               de Vigen�re
'---------------------------------------------------------------------------------------
'
    Dim sLettres    As String
    Dim lCompteur   As Long
    Dim lLongueur   As Long
    Dim lBoucle     As Long
    
    
    'D�finition de la longueur de la cha�ne � crypter et de la cha�ne de r�sultat
    lLongueur = Len(cha�neAdeCrypter)
    sLettres = String(lLongueur, Chr(0))


    'Boucler en fonction du nombre de rotations attendues
    For lBoucle = 1 To NBROTATIONSMAX
        'boucler pour chaque caract�re de la cha�ne initiale
        For lCompteur = 1 To lLongueur
            'Remplacer le caract�re de la cha�ne � crypter par le caract�re correspondant �
            '   le reste de
            '       la valeur ascii du caract�re � crypter
            '       plus
            '       la valeur ascii du caract�re correspondant dans la cl�, multipli� par la longueur de la cl�
            '   quand on le divise par 256
            Dim reste As Integer
            reste = asc(Mid(cha�neAdeCrypter, lCompteur, 1))
            
            
            
            
            
            Mid(sLettres, lCompteur, 1) = Chr((((asc(Mid(cha�neAdeCrypter, lCompteur, 1)) - _
                    (asc(Mid(Clef, (lCompteur Mod Len(Clef)) + 1, 1)) * lLongueur)) Mod 256) + 256) Mod 256)
        'recommencer
        Next
        'r�affecter la cha�ne � crypter par le r�sultat trouv� pour pouvoir recommencer une it�ration
        cha�neAdeCrypter = sLettres
    'Nouvelle it�ration
    Next
    'Renvoyer le r�sultat final
    deCrypter = sLettres
End Function

Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc
        
    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim Bytes() As Byte
        
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.Getbytes_4(sIn)
    Bytes = oMD5.ComputeHash_2((TextToHash))
 
    If bB64 = True Then
       MD5 = ConvToBase64String(Bytes)
    Else
       MD5 = ConvToHexString(Bytes)
    End If
        
    Set oT = Nothing
    Set oMD5 = Nothing

End Function

Public Function SHA1(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '40 Hex:   da39a3ee5e6...etc
    '28 Base-64:   2jmj7l5rSw0yVb...etc
    
    Dim oT As Object, oSHA1 As Object
    Dim TextToHash() As Byte
    Dim Bytes() As Byte
            
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    Bytes = oSHA1.ComputeHash_2((TextToHash))
        
    If bB64 = True Then
       SHA1 = ConvToBase64String(Bytes)
    Else
       SHA1 = ConvToHexString(Bytes)
    End If
            
    Set oT = Nothing
    Set oSHA1 = Nothing
    
End Function

Public Function SHA256(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '64 Hex:   e3b0c44298f...etc
    '44 Base-64:   47DEQpj8HBSa+/...etc
    
    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, Bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    Bytes = oSHA256.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA256 = ConvToBase64String(Bytes)
    Else
       SHA256 = ConvToHexString(Bytes)
    End If
    
    Set oT = Nothing
    Set oSHA256 = Nothing
    
End Function

Public Function SHA384(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '96 Hex:   38b060a751ac...etc
    '64 Base-64:   OLBgp1GsljhM2T...etc
    
    Dim oT As Object, oSHA384 As Object
    Dim TextToHash() As Byte, Bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    Bytes = oSHA384.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA384 = ConvToBase64String(Bytes)
    Else
       SHA384 = ConvToHexString(Bytes)
    End If
    
    Set oT = Nothing
    Set oSHA384 = Nothing
    
End Function

Public Function SHA512(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '128 Hex:   cf83e1357eefb8bd...etc
    '88 Base-64:   z4PhNX7vuL3xVChQ...etc
    
    Dim oT As Object, oSHA512 As Object
    Dim TextToHash() As Byte, Bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    
    TextToHash = oT.Getbytes_4(sIn)
    Bytes = oSHA512.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA512 = ConvToBase64String(Bytes)
    Else
       SHA512 = ConvToHexString(Bytes)
    End If
    
    Set oT = Nothing
    Set oSHA512 = Nothing
    
End Function

Public Function StrToSHA512Salt(ByVal sIn As String, ByVal sSecretKey As String, _
                           Optional ByVal b64 As Boolean = False) As String
    'Returns a sha512 STRING HASH in function name, modified by the parameter sSecretKey.
    'This hash differs from that of SHA512 using the SHA512Managed class.
    'HMAC class inputs are hashed twice;first input and key are mixed before hashing,
    'then the key is mixed with the result and hashed again.
    
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SecretKey() As Byte
    Dim Bytes() As Byte
    
    'Test results with both strings empty:
    '128 Hex:    b936cee86c9f...etc
    '88 Base-64:   uTbO6Gyfh6pd...etc
    
    'create text and crypto objects
    Set asc = CreateObject("System.Text.UTF8Encoding")
    
    'Any of HMACSHAMD5,HMACSHA1,HMACSHA256,HMACSHA384,or HMACSHA512 can be used
    'for corresponding hashes, albeit not matching those of Managed classes.
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA512")

    'make a byte array of the text to hash
    Bytes = asc.Getbytes_4(sIn)
    'make a byte array of the private key
    SecretKey = asc.Getbytes_4(sSecretKey)
    'add the private key property to the encryption object
    enc.Key = SecretKey

    'make a byte array of the hash
    Bytes = enc.ComputeHash_2((Bytes))
    
    'convert the byte array to string
    If b64 = True Then
       StrToSHA512Salt = ConvToBase64String(Bytes)
    Else
       StrToSHA512Salt = ConvToHexString(Bytes)
    End If
    
    'release object variables
    Set asc = Nothing
    Set enc = Nothing

End Function

Private Function ConvToBase64String(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Private Function ConvToHexString(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function
