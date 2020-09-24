Attribute VB_Name = "MSNProtocol"
Const IMSPROT = "MSNP2" 'Protocole d'échange avec le serveur
Const SECURITYPCK = "MD5" 'Protocole cryptographie
Const DSIP = "messenger.hotmail.com" 'Serveur principal
Const PORT = 1863 'Port pour se connecter

Const DS_TRID = 1

Const cmdXFR = "XFR"

Const ChrCode = "Ã¡Ã Ã¢Ã¤Ã©Ã¨ÃªÃ«Ã­Ã¬Ã®Ã¯Ã³Ã²Ã´Ã¶ÃºÃ¹Ã»Ã¼"
Const ChrNorm = "áàâäéèêëíìîïóòôöúùûü"
Const cOffline = 0
Const cConnected = 1
Const cVer = 2
Const cInf = 3
Const cUsr = 4
Const cHash = 5



Public Function Split(ByVal aText As String, OutputStr() As String, ByVal Delimiter As String) As Integer
    Dim pos As Integer, lastpos As Integer
    Dim argCounter As Integer, originalTxt As String
    originalstr = aText
    lastpos = 1


    If InStr(1, aText, Delimiter) = 0 Then
        ReDim OutputStr(1) As String
        OutputStr(1) = aText
        Split = 1
        Exit Function
    End If


    Do
        pos = InStr(1, aText, Delimiter)
        argCounter = argCounter + 1
        ReDim Preserve OutputStr(argCounter) As String


        If pos = 0 Then
            OutputStr(argCounter) = aText
            Exit Do
        End If
        OutputStr(argCounter) = Left(aText, pos - 1)
        aText = Mid(aText, pos + Len(Delimiter))
        lastpos = 1 ' Len(originalstr) - Len(Text)
    Loop
    Split = argCounter
End Function

Function Clean(tOrig As String) As String

'Espaces
Clean = Remplacer(tOrig, "%20", " ")

For n = 1 To Len(ChrNorm)
    Clean = Remplacer(tOrig, Mid(ChrCode, (2 * n) - 1, 2), Mid(ChrNorm, n, 1))
Next n

Clean = tOrig
End Function
Function DeClean(tOrig As String, cSpc As Boolean) As String

Dim n As Integer

If cSpc = True Then
    DeClean = Remplacer(tOrig, " ", "%20")
End If

For n = 1 To Len(ChrNorm)
    DeClean = Remplacer(tOrig, Mid(ChrNorm, n, 1), Mid(ChrCode, (2 * n) - 1, 2))
Next n

End Function

Public Function Remplacer(sData As String, sSubstring As String, sNewsubstring) As String

Dim i As String
Dim lSub As Long
Dim lData As Long
i = 1

lSub = Len(sSubstring)
lData = Len(sData)

Do
i = InStr(i, sData, sSubstring)
If i = 0 Then
    Remplacer = sData
    Exit Function
Else
    sData = Mid(sData, 1, i - 1) & sNewsubstring & Mid(sData, i + lSub, lData)
    'i = i + lSub
End If
Loop Until i > Len(sData) 'lData
Remplacer = sData
End Function



