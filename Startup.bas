Attribute VB_Name = "Startup"
Public Const CodeKey = "uits_gst"
Public Const ArrayLength = 60
Public Const CONNS = 25
Public Const FetchMax = "Top 150"
Public Const Pos1X = 50
Public Const Pos1Y = 600

Public XID As String
Public XREF As Long
Public RunningNews As String
Public ItemID As Long
Public SaleInvoiceTerms, SaleChallanTerms, SaleOrderTerms As String
Public SaleReturnTerms As String
Public PurchaseTerms As String
Public PurchaseReturnTerms As String

Public AppExpiryDate As Date

Public GUID As String
Public UserRights As Integer
Public UnifiedEntry As String

Public conn As ADODB.Connection
Public comm(CONNS) As ADODB.Command
Public recs(CONNS) As ADODB.Recordset

Public DatabaseServer As String
Public DCN As String
Public RecLocation As Integer
Public sArray(ArrayLength) As String
Public sArrayN(ArrayLength) As Integer
Public sSQL(CONNS) As String

'Public iniFile As String
Public defaultInitCatalog As String
Public initCatalog As String
Public login As String
Public passwd As String
Public Provider As String
Public PSI As String

Public MSComm1 As Control
Public bLoad As Boolean
Public bLoadMessage As String
Public frmSelectitemHeight As Long
Public frmShowHeight As Long

'PPSS Forms
Public Sale As New clsSale
Public SaleReturn As New clsSaleReturn
Public Purchase As New clsPurchase
Public PurchaseReturn As New clsPurchaseReturn
Public StockTransferIN As New clsStockTransferIN
Public StockTransferOUT As New clsStockTransferOUT

'PRT Forms
Public PYMT As New clsPymt
Public RCPT As New clsRcpt

'Messenger Forms
Public MSN As New clsMessenger

Public Sub Main()
    ''' THIS IS THE ENTRY POINT OF THIS PROJECT
    MainEntry
End Sub

Public Sub MainEntry()
    Set conn = New ADODB.Connection
    'iniFile = App.Path & "\uits.ini"
    
    DCN = ""
    App.Title = CompanyName
    AppExpiryDate = CDate("15-AUG-2040")
    
    SetCompanyInfo
    CreateTermsAndConditions
    
    Provider = "SQLOLEDB.1"
    PSI = "True"
    defaultInitCatalog = mdiOne.sckGo.GReadINI("[initCatalog]")
    DatabaseServer = mdiOne.sckGo.GReadINI("[DatabaseServer]")
    login = mdiOne.sckGo.GReadINI("[login]")
    passwd = Decipher(mdiOne.sckGo.GReadINI("[passwd]"))
        
    bLoad = True
    If Date > AppExpiryDate And Format(Time, "hhnnss") < "130000" Then
        bLoad = False: bLoadMessage = "A. Software configuration error! Call your software vendor for maintenance."
    Else
        bLoad = ConnectToDatabase("master", DatabaseServer)
        If bLoad = False Then bLoadMessage = "B. Software configuration DB error! Call your software vendor for maintenance."
    End If
    
    If bLoad = True Then
        mdiOne.Show
    Else
        MsgBox bLoadMessage, vbOKOnly + vbCritical + vbMsgBoxSetForeground
    End If
End Sub

Public Sub dbOpen(i As Integer)
    'Set & Execute command
    Set comm(i) = New ADODB.Command: comm(i).ActiveConnection = conn: comm(i).CommandType = adCmdText
    
    'set & open recordset
    Set recs(i) = New ADODB.Recordset: recs(i).CursorLocation = adUseClient: recs(i).CursorType = adOpenDynamic
    recs(i).LockType = adLockOptimistic: recs(i).Open sSQL(i), conn
End Sub

Public Sub dbClose(i As Integer)
    If IsEmpty(recs(i)) Then recs(i).Close
    Set recs(i) = Nothing: Set comm(i) = Nothing
End Sub

Public Sub ClearsArray(Optional i As Integer)
    Erase sArray
End Sub

Public Sub FillsArray(i As Integer)
    If recs(i).EOF <> True Then
        For k = 0 To recs(0).FIELDS.Count - 1
            If Not IsNull(recs(i).FIELDS(k)) Then sArray(k) = recs(i).FIELDS(k)
        Next
    Else
        For k = 0 To recs(0).FIELDS.Count - 1
            sArray(k) = "X"
        Next
    End If
End Sub

Public Function ReadFont(ByVal S As String, ByVal i As Integer) As String
    fnt = Split(mdiOne.sckGo.GReadINI(S), ":")
    If UBound(fnt) >= i Then
        ReadFont = fnt(i)
    Else
        ReadFont = ""
    End If
End Function

Public Sub MsgMdiBox(msg As String, Optional X As Integer)
    msgUITS (msg)
End Sub

Public Function ConnectToDatabase(ByVal cInitCatalog As String, ByVal cDatabaseServer As String) As Boolean
    On Error Resume Next
    initCatalog = cInitCatalog
    DCN = ""
    DCN = DCN & "Provider=" & Provider & ";"
    DCN = DCN & "Persist Security Info=" & PSI & ";"
    DCN = DCN & "Initial Catalog=" & cInitCatalog & ";"
    DCN = DCN & "Data Source=" & cDatabaseServer & ";"
    DCN = DCN & "User ID=" & login & ";"
    DCN = DCN & "Password=" & passwd & ";"
    
    'Set & Open connection
    conn.Close
    conn.ConnectionString = DCN
    conn.Open
    If conn.State = 1 Then
        ConnectToDatabase = True
    Else
        MsgBox Error & vbCrLf & "Unable to connect to the database. Contact administrator!", vbOKOnly + vbCritical
        frmDatabaseConnection.Show vbModal
        If conn.State = 1 Then
            ConnectToDatabase = True
        Else
            ConnectToDatabase = False
        End If
    End If
End Function

Public Sub msgUITS(ByVal msg As String)
        'MsgBox msg, vbOKOnly
        mdiOne.StatusBar1.Panels(2).Text = msg
End Sub

Public Sub OutBoundRule()
        Const NET_FW_PROFILE2_DOMAIN = 1
        Const NET_FW_PROFILE2_PRIVATE = 2
        Const NET_FW_PROFILE2_PUBLIC = 4
        
        Dim fwPolicy2
        Set fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")
        fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_DOMAIN) = False
        fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PRIVATE) = False
        fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PUBLIC) = False
End Sub

Public Function QT(ByVal str As String) As String
    QT = Chr(39) & str & Chr(39)
    'QT = mdiOne.sckGo.InQuote(str)
End Function

Private Sub CreateTermsAndConditions()
    Terms = "Terms & Conditions:-"
    Terms = Terms & vbCrLf & "1. Payment/other terms as per agreement."
    Terms = Terms & vbCrLf & "2. All disputes subject to Company HO State Jurisdiction."
    Terms = Terms & vbCrLf & "3. Goods once sold will not be taken back."
    Terms = Terms & vbCrLf & "4. Our responsibility ceases when goods leave our godown."
    SaleInvoiceTerms = Terms
End Sub
