Attribute VB_Name = "Phone"
Dim bQuit As Boolean

Public Sub Dial(PhNum As String)
    bQuit = False
    If PhNum = "" Then
        MsgBox "No number to dial", vbOKOnly, "Sorry"
        Exit Sub
    End If
    OutDialCode = mdiOne.sckGo.GReadINI("[OutDialCode]")
    If Not MSComm1.PortOpen Then OpenCom
    msgUITS "Dialing " & OutDialCode & PhNum
    MSComm1.Output = "AT DT" & OutDialCode & PhNum & vbCrLf
    ReadModem
End Sub

Public Sub OpenCom()
    MSComm1.CommPort = Val(mdiOne.sckGo.GReadINI("[ModemComPort]"))
    MSComm1.Settings = "14400,N,8,1"
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
End Sub

Public Sub ReadModem()
    Dim InString As String
    Do While Not bQuit
        If MSComm1.InBufferCount Then
            InString = InString & MSComm1.Input
            If InStr(InString, vbCrLf) Then
                DoEvents
                If InStr(InString, "BUSY") Then
                    bQuit = True
                    msgUITS "Phone busy"
                End If
                If InStr(InString, "NO") Then
                    bQuit = True
                    msgUITS "Phone busy"
                End If
                InString = ""
            End If
        End If
        DoEvents
    Loop
    HangUp
End Sub

Public Sub HangUp()
    If MSComm1.PortOpen = False Then Exit Sub
    bQuit = True
    MSComm1.PortOpen = False
    msgUITS "CALL ENDED!"
End Sub
