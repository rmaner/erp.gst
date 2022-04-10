Attribute VB_Name = "CipherDecipher"
Private Const CipherShift = 53

Public Function Cipher(ByVal str As String) As String
    On Error Resume Next
    Dim strCipher As String
    strCipher = ""
    For i = 1 To Len(str)
        strCipher = strCipher & Chr(Asc(Mid(str, i, 1)) + CipherShift)
    Next
    Cipher = strCipher
End Function

Public Function Decipher(ByVal str As String) As String
    On Error Resume Next
    Dim strDecipher As String
    strDecipher = ""
    For i = 1 To Len(str)
        strDecipher = strDecipher & Chr(Asc(Mid(str, i, 1)) - CipherShift)
    Next
    Decipher = strDecipher
End Function

