VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJournals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const PPSSJrEntryCount = 8

Public Function PPSSJournalEntries(ByVal Support As String, ByVal DBRef As Long, ByVal MemoFormat As String) As Boolean
    Dim cData(5) As New clsData
    Dim B As Boolean
    Dim DrAc(PPSSJrEntryCount), CrAc(PPSSJrEntryCount), Narration As String
    Dim Table, AdjustmentTable, AdjustmentMemoFormat, Party, SD, Transporter, Karter, Post, Status As String
    Dim DrCrAmt(PPSSJrEntryCount) As Double
    Dim BillDate, AuthDate As Date
    Dim AuthDelay As Integer
    Dim AdjustmentType As Integer
    
    Select Case UCase(Support)
        Case "SALE": Table = "SMAIN"
        Case "PURCHASE": Table = "PMAIN"
        Case "SALERETURN": Table = "SRETURNMAIN"
        Case "PURCHASERETURN": Table = "PRETURNMAIN"
        Case "TIN": Table = "TINMAIN"
        Case "TOUT": Table = "TOUTMAIN"
    End Select
    
    B = True
    cData(1).dbOpen "DELETE JOURNAL WHERE MemoRef=" & QT(Format(DBRef, MemoFormat)), 1
    cData(1).dbOpen "SELECT * FROM " & Table & " WHERE DBRef=" & DBRef, 1
    If cData(1).recs.RecordCount = 1 Then
        B = True
    Else
        B = False
    End If
    
    If B = True Then
        BillDate = CDate(cData(1).recs!DBDate)
        Status = cData(1).recs!Status
        Party = cData(1).recs!id
        Transporter = cData(1).recs!TID
        Karter = cData(1).recs!KID
        Post = cData(1).recs!PID
        
        AuthDate = CDate(cData(1).recs!DBDate)
        Narration = Left(cData(1).recs!id & ";" & cData(1).recs!Name & ";" & cData(1).recs!State & "|" & cData(1).recs!Comments, 149)
        DrCrType = UCase(Status)
        
        If DrCrType = "CASH" Then Party = "R0002"
        
        Select Case UCase(Support)
            Case "PURCHASE":
                DrAc(0) = "N0011": CrAc(0) = Party:         DrCrAmt(0) = Val(cData(1).recs!NetAmount)     'PARTY
                DrAc(4) = "X0001": CrAc(4) = "N0001":       DrCrAmt(4) = Val(cData(1).recs!NetCGSTAmt)
                DrAc(5) = "X0002": CrAc(5) = "N0002":       DrCrAmt(5) = Val(cData(1).recs!NetSGSTAmt)
                DrAc(6) = "X0003": CrAc(6) = "N0003":       DrCrAmt(6) = Val(cData(1).recs!NetIGSTAmt)
            
            Case "PURCHASERETURN":
                DrAc(0) = Party: CrAc(0) = "N0012":         DrCrAmt(0) = Val(cData(1).recs!NetAmount)     'PARTY
                DrAc(4) = "N0001": CrAc(4) = "X0001":       DrCrAmt(4) = Val(cData(1).recs!NetCGSTAmt)
                DrAc(5) = "N0002": CrAc(5) = "X0002":       DrCrAmt(5) = Val(cData(1).recs!NetSGSTAmt)
                DrAc(6) = "N0003": CrAc(6) = "X0003":       DrCrAmt(6) = Val(cData(1).recs!NetIGSTAmt)
            
            
            Case "SALE":
                DrAc(0) = Party: CrAc(0) = "N0013":         DrCrAmt(0) = Val(cData(1).recs!NetAmount)     'PARTY
                DrAc(4) = "N0001": CrAc(4) = "X0001":       DrCrAmt(4) = Val(cData(1).recs!NetCGSTAmt)
                DrAc(5) = "N0002": CrAc(5) = "X0002":       DrCrAmt(5) = Val(cData(1).recs!NetSGSTAmt)
                DrAc(6) = "N0003": CrAc(6) = "X0003":       DrCrAmt(6) = Val(cData(1).recs!NetIGSTAmt)
            
            Case "SALERETURN":
                DrAc(0) = "N0014": CrAc(0) = Party:         DrCrAmt(0) = Val(cData(1).recs!NetAmount)     'PARTY
                DrAc(4) = "X0001": CrAc(4) = "N0001":       DrCrAmt(4) = Val(cData(1).recs!NetCGSTAmt)
                DrAc(5) = "X0002": CrAc(5) = "N0002":       DrCrAmt(5) = Val(cData(1).recs!NetSGSTAmt)
                DrAc(6) = "X0003": CrAc(6) = "N0003":       DrCrAmt(6) = Val(cData(1).recs!NetIGSTAmt)
            
            Case "TOUT":
                DrAc(0) = Party: CrAc(0) = "N0015":         DrCrAmt(0) = Val(cData(1).recs!NetAmount)     'PARTY
                DrAc(4) = "N0001": CrAc(4) = "X0001":       DrCrAmt(4) = Val(cData(1).recs!NetCGSTAmt)
                DrAc(5) = "N0002": CrAc(5) = "X0002":       DrCrAmt(5) = Val(cData(1).recs!NetSGSTAmt)
                DrAc(6) = "N0003": CrAc(6) = "X0003":       DrCrAmt(6) = Val(cData(1).recs!NetIGSTAmt)
            
            Case "TIN":
                DrAc(0) = "N0016": CrAc(0) = Party:         DrCrAmt(0) = Val(cData(1).recs!NetAmount)     'PARTY
                DrAc(4) = "X0001": CrAc(4) = "N0001":       DrCrAmt(4) = Val(cData(1).recs!NetCGSTAmt)
                DrAc(5) = "X0002": CrAc(5) = "N0002":       DrCrAmt(5) = Val(cData(1).recs!NetSGSTAmt)
                DrAc(6) = "X0003": CrAc(6) = "N0003":       DrCrAmt(6) = Val(cData(1).recs!NetIGSTAmt)
        End Select
        
            DrAc(1) = "N0010": CrAc(1) = Transporter:   DrCrAmt(1) = Val(cData(1).recs!GRAmount)    'TRANSPORTATION/FREIGHT
            DrAc(2) = "N0011": CrAc(1) = Karter:        DrCrAmt(2) = Val(cData(1).recs!KAmount)     'CARTAGE
            DrAc(3) = "N0012": CrAc(3) = Post:          DrCrAmt(3) = Val(cData(1).recs!Postage)     'POSTAGE
            
            DrAc(7) = "N0004": CrAc(7) = "X0004":       DrCrAmt(7) = Val(cData(1).recs!NetCessAmt)
        
        
        For i = 0 To PPSSJrEntryCount - 1
            If DrCrAmt(i) <> 0 Then B = B And HasAccount(DrAc(i)) And HasAccount(CrAc(i))
        Next
        
        If Not (Status = "CASH" Or Status = "CREDIT" Or Status = "CHALLAN") Then B = False
        If B = True Then
            For i = 0 To PPSSJrEntryCount - 1
                If DrCrAmt(i) <> 0 Then MakeJournalEntry BillDate, DrAc(i), CrAc(i), DrCrAmt(i), Narration, Format(DBRef, MemoFormat), AuthDate
            Next
        End If
    End If
    PPSSJournalEntries = B
End Function

Public Function PRTJournalEntries(ByVal Table As String, ByVal Serial As Long, ByVal MemoFormat As String) As Boolean
    Dim JCon(5) As New clsData
    Dim B As Boolean
    Dim DrAc, CrAc, Party, MoneyAccount, Mode, Narration As String
    Dim MemoDate, AuthDate As Date
    
    JCon(0).dbOpen "DELETE JOURNAL WHERE MemoRef=" & QT(Format(Serial, MemoFormat))
    JCon(0).dbOpen "SELECT * FROM " & Table & " WHERE Serial=" & Serial
    
    B = True
    If Not JCon(0).recs.EOF Then
        MemoDate = CDate(JCon(0).recs!Date)
        Party = JCon(0).recs!id
        DrCrAmt = Val(JCon(0).recs!Amount)
        AuthDate = CDate(JCon(0).recs!Date)
        Narration = Trim(JCon(0).recs!Narration)
        Mode = JCon(0).recs!Mode
        MoneyAccount = JCon(0).recs!Mode
    Else
        B = False
    End If
    
    Select Case UCase(Table)
        Case "PMT":
            DrAc = Party: CrAc = MoneyAccount: DrCrAmt = DrCrAmt
        Case "RCT":
            DrAc = MoneyAccount: CrAc = Party: DrCrAmt = DrCrAmt
    End Select
    
    B = B And HasAccount(DrAc) And HasAccount(CrAc)
    
    If B = True Then
        MakeJournalEntry MemoDate, DrAc, CrAc, DrCrAmt, Narration, Format(Serial, MemoFormat), AuthDate
    End If
    PRTJournalEntries = B
End Function

Public Function VoucherToJournalEntries(ByVal Serial As Long, ByVal MemoFormat As String) As Boolean
    Dim JCon(5) As New clsData
    Dim B As Boolean
    Dim DrAc, CrAc, Narration As String
    Dim MemoDate, AuthDate As Date
    
    JCon(0).dbOpen "DELETE JOURNAL WHERE MemoRef=" & QT(Format(Serial, MemoFormat))
    JCon(0).dbOpen "SELECT * FROM VOUCHERS WHERE Serial=" & Serial
    
    B = True
    If Not JCon(0).recs.EOF Then
        MemoDate = CDate(JCon(0).recs!Date)
        DrAc = JCon(0).recs!DrID
        CrAc = JCon(0).recs!CrID
        DrCrAmt = Val(JCon(0).recs!Amount)
        AuthDate = CDate(JCon(0).recs!Date)
        Narration = Trim(JCon(0).recs!Narration)
    Else
        B = False
    End If
    
    B = B And HasAccount(DrAc) And HasAccount(CrAc)
    
    If B = True Then
        MakeJournalEntry MemoDate, DrAc, CrAc, DrCrAmt, Narration, Format(Serial, MemoFormat), AuthDate
    End If
    VoucherToJournalEntries = B
End Function



Public Function TRFJournalEntries(ByVal Table As String, ByVal Serial As Long, ByVal MemoFormat As String) As Boolean
    Dim JCon(5) As New clsData
    Dim B As Boolean
    Dim DrAc, CrAc, Narration As String
    Dim MemoDate, AuthDate As Date
    
    JCon(0).dbOpen "DELETE JOURNAL WHERE MemoRef=" & QT(Format(Serial, MemoFormat))
    JCon(0).dbOpen "SELECT * FROM " & Table & " WHERE Serial=" & Serial
    
    B = True
    If Not JCon(0).recs.EOF Then
        MemoDate = CDate(JCon(0).recs!Date)
        DrAc = JCon(0).recs!DrID
        CrAc = JCon(0).recs!CrID
        DrCrAmt = Val(JCon(0).recs!Amount)
        AuthDate = CDate(JCon(0).recs!Date)
        Narration = Table
    Else
        B = False
    End If
    
    B = B And HasAccount(DrAc) And HasAccount(CrAc)
    If B = True Then
        MakeJournalEntry MemoDate, DrAc, CrAc, DrCrAmt, Narration, Format(Serial, MemoFormat), AuthDate
    End If
    TRFJournalEntries = B
End Function

Public Function MakeJournalEntry(ByVal jDate As Date, ByVal DrAc As String, ByVal CrAc As String, ByVal Amount As Double, ByVal Narration As String, ByVal MemoRef As String, ByVal AuthDate As Date) As Boolean
    ' THIS FUNCTION MAKES NEW JOURNAL ENTRY/OR UPDATES EXISTING ENTRY FOR AMOUNT AND NARRATION/ IF DRAC & CRAC ARE INVALID ACCOUNTS THEN IT RETURNS 0
    
    Dim cData As New clsData
    Dim DrName, CrName, msg As String
    Dim CanMakeJournalEntry As Boolean
    
    CanMakeJournalEntry = True
    
    cData.dbOpen "SELECT * FROM appview_ALLACCOUNTS where ID=" & QT(DrAc), 1
    If cData.recs.RecordCount < 1 Then
        CanMakeJournalEntry = False
    Else
        DrName = cData.recs!Name
    End If

    cData.dbOpen "SELECT * FROM appview_ALLACCOUNTS where ID=" & QT(CrAc), 1
    If cData.recs.RecordCount < 1 Then
        CanMakeJournalEntry = False
    Else
        CrName = cData.recs!Name
    End If

    If CanMakeJournalEntry Then
        SQ = "INSERT INTO JOURNAL (Date, DrAC, DrName, CrAC, CrName, Amount, Narration, UserID, MemoRef, AuthDate) VALUES ("
        SQ = SQ & QT(Format(jDate, "dd-MMM-yy")) & ","
        SQ = SQ & QT(DrAc) & "," & QT(DrName) & ","
        SQ = SQ & QT(CrAc) & "," & QT(CrName) & ","
        SQ = SQ & Amount & "," & QT(Narration) & ","
        SQ = SQ & QT(GUID) & "," & QT(MemoRef) & "," & QT(Format(AuthDate, "DD-MMM-YY HH:MM:SS"))
        SQ = SQ & ")"
        cData.dbOpen SQ, 1
        
        msg = MemoRef & " : " & Format(AuthDate, "dd-MMM-yy") & " : " & Format(Amount, "##,###.00") & " :  " & CrName & " ->" & DrName
        mdiOne.StatusBar1.Panels(2).Text = msg
        DoEvents
    End If
    MakeJournalEntry = CanMakeJournalEntry
End Function
