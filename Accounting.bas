Attribute VB_Name = "Accounting"
Public ErrorMsg As String
Public Const NoJrEntries = 5

Public Function RamjeeRound(ByVal N As Double) As Double
    RamjeeRound = Int(-N / 10) * -10
End Function

Public Function CircularDiscount(ByVal DiscStr As String) As Double
    Dim disc As Double
    a = Split(DiscStr, "+")
    If UBound(a) >= 0 Then disc = Val(a(0))
    For i = 1 To UBound(a)  ' CIRCULAR DISCOUNT CALCULATION EXAMPLE 20+10+10+10+5
        X = disc: Y = Val(a(i))
        disc = X + (Y - X * Y / 100)
    Next
    CircularDiscount = disc
End Function

Public Function FinancialYear(ByVal d As Date) As String
    Dim mth As Integer
    Dim fyear As String
    fyear = ""
    If IsDate(d) Then
        mth = Month(d)
        Select Case mth
            Case 4 To 12
                fyear = Right(Year(d), 2) & "-" & Right(Year(d) + 1, 2)
            Case 1 To 3
                fyear = Right(Year(d) - 1, 2) & "-" & Right(Year(d), 2)
        End Select
    End If
    FinancialYear = fyear
End Function

Public Function GetOrderStatus(ByVal TBL As String, ByVal DBRef As Long) As String
    sSQL(0) = "SELECT Status FROM " & TBL & " WHERE DBRef=" & DBRef
    dbOpen (0)
    If recs(0).RecordCount = 1 Then
        GetOrderStatus = recs(0)!Status
    Else
        GetOrderStatus = ""
    End If
    dbClose (0)
End Function

Public Function HasAccount(ByVal AcID As String) As Boolean
    sSQL(0) = "SELECT * FROM appview_AllAccounts where ID=" & Chr(39) & AcID & Chr(39)
    dbOpen (0)
    If recs(0).RecordCount < 1 Then
        msgUITS AcID & " has no or blocked account."
        HasAccount = False
    Else
        HasAccount = True
    End If
    dbClose (0)
End Function

Public Function MakeJournalEntry(ByVal jDate As Date, ByVal DrAc As String, ByVal CrAc As String, ByVal Amount As Double, ByVal Narration As String, ByVal MemoRef As String, ByVal AuthDate As Date) As Boolean
    ' THIS FUNCTION MAKES NEW JOURNAL ENTRY
    ' OR UPDATES EXISTING ENTRY FOR AMOUNT AND NARRATION
    ' IF DRAC & CRAC ARE INVALID ACCOUNTS THEN IT RETURNS 0
    
    Dim DrName, CrName As String
    Dim CanMakeJournalEntry As Boolean
    
    CanMakeJournalEntry = True
    
    sSQL(0) = "SELECT * FROM appview_AllAccounts where ID=" & Chr(39) & DrAc & Chr(39): dbOpen (0)
    If recs(0).RecordCount < 1 Then
        CanMakeJournalEntry = False
    Else
        DrName = recs(0)!Name
    End If
    dbClose (0)

    sSQL(0) = "SELECT * FROM appview_AllAccounts where ID=" & Chr(39) & CrAc & Chr(39): dbOpen (0)
    If recs(0).RecordCount < 1 Then
        CanMakeJournalEntry = False
    Else
        CrName = recs(0)!Name
    End If
    dbClose (0)

    If CanMakeJournalEntry Then
        SQ = "INSERT INTO JOURNAL (Date, DrAC, DrName, CrAC, CrName, Amount, Narration, UserID, MemoRef, AuthDate) VALUES ("
        SQ = SQ & QT(Format(jDate, "MM-dd-yyyy")) & ","
        SQ = SQ & QT(DrAc) & "," & QT(DrName) & ","
        SQ = SQ & QT(CrAc) & "," & QT(CrName) & ","
        SQ = SQ & Amount & "," & QT(Narration) & ","
        SQ = SQ & QT(GUID) & "," & QT(MemoRef) & "," & QT(Format(AuthDate, "MM-dd-yyyy"))
        SQ = SQ & ")"
        sSQL(0) = SQ
        dbOpen (0): dbClose (0)
        mdiOne.StatusBar1.Panels(2).Text = MemoRef & " : " & Format(AuthDate, "dd-MMM-yy") & " : " & Format(Amount, "##,###.00") & " :  " & CrName & " ->" & DrName
    End If
    MakeJournalEntry = CanMakeJournalEntry
End Function

Public Function WhatIsLedgerBalance(ByVal id As String, Optional dt As Variant) As Currency
    If IsMissing(dt) Then dt = Now
    sSQL(3) = "appproc_LEDGERBAL " & QT(id) & ", " & QT(Format(dt, "DD-MMM-YYYY"))
    dbOpen (3)
    If recs(3).RecordCount = 1 Then
        WhatIsLedgerBalance = recs(3)!BAL
    Else
        WhatIsLedgerBalance = 0
    End If
    dbClose (3)
    sSQL(3) = ""
End Function

Public Function CreditNoteAdjustmentDate(CNDate As Date) As Date
    yr = Year(CNDate)
    If CNDate >= CDate("1-9-" & str(yr)) And CNDate <= CDate("14-11-" & str(yr)) Then CreditNoteAdjustmentDate = CDate("15-1-" & str(yr + 1))
    If CNDate >= CDate("1-4-" & str(yr)) And CNDate <= CDate("14-6-" & str(yr)) Then CreditNoteAdjustmentDate = CDate("15-8-" & str(yr))
    CreditNoteAdjustmentDate = CDate("02-10-1857")
End Function

Public Function MakeMultipleJournalEntries(ByVal Support As String, ByVal DBRef As Long, ByVal MemoFormat As String) As Boolean
    Dim B As Boolean
    Dim DrAc(NoJrEntries), CrAc(NoJrEntries), Narration As String
    Dim Table, Party, SD, Transporter, Karter, Post, Status As String
    Dim DrCrAmt(NoJrEntries) As Double
    Dim BillDate, AuthDate As Date
    
    Select Case UCase(Support)
        Case "SALE": Table = "SMAIN"
        Case "PURCHASE": Table = "PMAIN"
        Case "SALERETURN": Table = "SRETURNMAIN"
        Case "PURCHASERETURN": Table = "PRETURNMAIN"
    End Select
    sSQL(0) = "DELETE JOURNAL WHERE MemoRef=" & QT(Format(DBRef, MemoFormat)): dbOpen (0): dbClose (0)
    
    sSQL(0) = "SELECT * FROM " & Table & " WHERE DBRef=" & DBRef
    dbOpen (0)
    If recs(0).RecordCount = 1 Then
        B = True
        ClearsArray (0): FillsArray (0)
    Else
        B = False
    End If
    dbClose (0)
    
    If B = True Then
        BillDate = CDate(sArray(6))
        Status = sArray(2)
        Party = sArray(7)
        SD = sArray(29)
        Transporter = sArray(14)
        Karter = sArray(22)
        Post = sArray(26)
        Narration = Support
        DrCrType = UCase(Status)
        If DrCrType = "CASH" Then Party = "R0002"
        If DrCrType = "CHALLAN" Then
            AuthDate = DateAdd("D", ChallanDelay, BillDate)
        Else
            AuthDate = BillDate
        End If
    
        Select Case UCase(Support)
            Case "SALE":
            DrAc(0) = Party: CrAc(0) = "N003": DrCrAmt(0) = Val(sArray(48)) 'PARTY
            DrAc(1) = "N009": CrAc(1) = SD: DrCrAmt(1) = Val(sArray(31))    'SD
            Case "PURCHASE":
            DrAc(0) = "N001": CrAc(0) = Party: DrCrAmt(0) = Val(sArray(48))
            DrAc(1) = "N009": CrAc(1) = SD: DrCrAmt(1) = Val(sArray(31))
            Case "SALERETURN":
            DrAc(0) = "N004": CrAc(0) = Party: DrCrAmt(0) = Val(sArray(48))
            DrAc(1) = SD: CrAc(1) = "N009": DrCrAmt(1) = Val(sArray(31))
            Case "PURCHASERETURN":
            DrAc(0) = Party: CrAc(0) = "N002": DrCrAmt(0) = Val(sArray(48))
            DrAc(1) = SD: CrAc(1) = "N009": DrCrAmt(1) = Val(sArray(31))
        End Select
        DrAc(2) = "N008": CrAc(2) = Transporter: DrCrAmt(2) = Val(sArray(35))   'FREIGHT
        DrAc(3) = "N006": CrAc(3) = Karter: DrCrAmt(3) = Val(sArray(25))        'CARTAGE
        DrAc(4) = "N016": CrAc(4) = Post: DrCrAmt(4) = Val(sArray(28))          'POSTAGE
        
        For i = 0 To NoJrEntries - 1
            If DrCrAmt(i) <> 0 Then B = B And HasAccount(DrAc(i)) And HasAccount(CrAc(i))
        Next
        
        If Val(sArray(40)) = 0 Then     'Blocked for Zero Amount Commodity
            B = False: msgUITS "Failed 0 amount transaction for " & Format(OrderRef, MemoFormat)
        End If
            
        If B = True Then
            For i = 0 To NoJrEntries - 1
                If DrCrAmt(i) <> 0 Then
                    MakeJournalEntry BillDate, DrAc(i), CrAc(i), DrCrAmt(i), Narration, Format(DBRef, MemoFormat), AuthDate
                    Debug.Print BillDate, DrAc(i), CrAc(i), DrCrAmt(i), Narration, Format(DBRef, MemoFormat), AuthDate
                End If
            Next
        End If
    End If
    MakeMultipleJournalEntries = B
End Function

Public Function MakeSingleJournalEntry(ByVal Table As String, ByVal Serial As Long, ByVal MemoFormat As String) As Boolean
    Dim B As Boolean
    Dim DrAc, CrAc, Party, MoneyAccount, Mode, Narration As String
    Dim DrCrAmt As Double
    Dim MemoDate, AuthDate As Date
    
    sSQL(0) = "SELECT * FROM " & Table & " WHERE Serial=" & Serial
    dbOpen (0): ClearsArray (0): FillsArray (0): dbClose (0)
    sSQL(0) = "DELETE JOURNAL WHERE MemoRef=" & QT(Format(Serial, MemoFormat)): dbOpen (0): dbClose (0)
        
    MemoDate = CDate(sArray(1))
    Party = sArray(2)
    DrCrAmt = Val(sArray(4))
    Narration = Table
    Mode = sArray(5)
    If Mode = "CASH" Then
        MoneyAccount = "R0002"
        AuthDate = DateAdd("D", 0, MemoDate)
    Else
        MoneyAccount = "R0001"
        AuthDate = DateAdd("D", BankDelay, MemoDate)
    End If

    Select Case UCase(Table)
        Case "PMT":
            DrAc = Party: CrAc = MoneyAccount: DrCrAmt = DrCrAmt
        Case "RCT":
            DrAc = MoneyAccount: CrAc = Party: DrCrAmt = DrCrAmt
    End Select
    
    B = True
    B = B And HasAccount(DrAc) And HasAccount(CrAc)
    
    If B = True Then
        MakeJournalEntry MemoDate, DrAc, CrAc, DrCrAmt, Narration, Format(Serial, MemoFormat), AuthDate
        Debug.Print MemoDate, DrAc, CrAc, DrCrAmt, Narration, Format(Serial, MemoFormat), AuthDate
    End If
    MakeSingleJournalEntry = B
End Function
