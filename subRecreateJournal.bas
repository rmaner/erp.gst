Attribute VB_Name = "subRecreateJournal"
Private cJr As New clsJournals

Public Sub ReCreateJournalEntries()
    Dim MaxDBRef As Long
    Dim MaxSerial As Long
    
    If MsgBox("Recreate journal entries...", vbYesNo + vbCritical) = vbYes Then
        sSQL(0) = "TRUNCATE TABLE JOURNAL": dbOpen (0): dbClose (0)
        MsgBox "Existing journal entries deleted...", vbOKOnly
        
        'SALE/SALERETURN/PURCHASE/PURCHASERETURN
        sSQL(0) = "SELECT isnull(MAX(DBRef),0) as DBRef FROM SMAIN": dbOpen (0): MaxDBRef = recs(0)!DBRef: dbClose (0)
        For i = 1 To MaxDBRef
            cJr.PPSSJournalEntries "SALE", i, "\S\A0"
        Next
        sSQL(0) = "SELECT isnull(MAX(DBRef),0) as DBRef  FROM SRETURNMAIN": dbOpen (0): MaxDBRef = recs(0)!DBRef: dbClose (0)
        For i = 1 To MaxDBRef
            cJr.PPSSJournalEntries "SALERETURN", i, "\S\R0"
        Next
        sSQL(0) = "SELECT isnull(MAX(DBRef),0) as DBRef  FROM PMAIN": dbOpen (0): MaxDBRef = recs(0)!DBRef: dbClose (0)
        For i = 1 To MaxDBRef
            cJr.PPSSJournalEntries "PURCHASE", i, "\P\U0"
        Next
        sSQL(0) = "SELECT isnull(MAX(DBRef),0) as DBRef  FROM PRETURNMAIN": dbOpen (0): MaxDBRef = recs(0)!DBRef: dbClose (0)
        For i = 1 To MaxDBRef
            cJr.PPSSJournalEntries "PURCHASERETURN", i, "\P\R0"
        Next
        sSQL(0) = "SELECT isnull(MAX(DBRef),0) as DBRef  FROM TOUTMAIN": dbOpen (0): MaxDBRef = recs(0)!DBRef: dbClose (0)
        For i = 1 To MaxDBRef
            cJr.PPSSJournalEntries "TOUT", i, "\T\O0"
        Next
        sSQL(0) = "SELECT isnull(MAX(DBRef),0) as DBRef  FROM TINMAIN": dbOpen (0): MaxDBRef = recs(0)!DBRef: dbClose (0)
        For i = 1 To MaxDBRef
            cJr.PPSSJournalEntries "TIN", i, "\T\I0"
        Next
    
        'PMT/RCT
        sSQL(0) = "SELECT isnull(MAX(Serial),0) as Serial  FROM PMT": dbOpen (0): MaxSerial = recs(0)!Serial: dbClose (0)
        For i = 1 To MaxSerial
            cJr.PRTJournalEntries "PMT", i, "\P\T0"
        Next
        sSQL(0) = "SELECT isnull(MAX(Serial),0)  as Serial FROM RCT": dbOpen (0): MaxSerial = recs(0)!Serial: dbClose (0)
        For i = 1 To MaxSerial
            cJr.PRTJournalEntries "RCT", i, "\R\T0"
        Next
        sSQL(0) = "SELECT isnull(MAX(Serial),0)  as Serial FROM VOUCHERS": dbOpen (0): MaxSerial = recs(0)!Serial: dbClose (0)
        For i = 1 To MaxSerial
            cJr.VoucherToJournalEntries i, "\V\R0"
        Next
        
        MsgBox ErrorMsg, vbOKOnly + vbExclamation
    End If
End Sub

