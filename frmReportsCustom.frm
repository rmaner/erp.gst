VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmReportsCustom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Reports ..."
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
   Icon            =   "frmReportsCustom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   12420
      TabIndex        =   0
      Top             =   8130
      Width           =   12480
      Begin VB.PictureBox Picture2 
         Height          =   405
         Left            =   5790
         ScaleHeight     =   345
         ScaleWidth      =   4335
         TabIndex        =   5
         Top             =   -15
         Width           =   4395
         Begin VB.CommandButton cmdQuit 
            Caption         =   "&Quit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2895
            TabIndex        =   8
            Top             =   0
            Width           =   1440
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            TabIndex        =   7
            Top             =   0
            Width           =   1440
         End
         Begin VB.CommandButton cmdSaveGrid 
            Caption         =   "SaveGrid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   15
            TabIndex        =   6
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.ComboBox cmbReportHead 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   30
         Width           =   3720
      End
      Begin VB.ComboBox cmbReportString 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2910
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   30
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5820
         TabIndex        =   3
         Top             =   45
         Width           =   1350
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   3  'Align Left
      Height          =   8130
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12435
      _cx             =   21934
      _cy             =   14340
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12632319
      ForeColorSel    =   64
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   4
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   4
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   50
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmReportsCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(5) As New clsData

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    cmbReportHead.AddItem "LedgerBalances Parties"
    cmbReportHead.AddItem "LedgerBalances Publishers Distributors"
    cmbReportHead.AddItem "LedgerBalances Parties-NonZero"
    cmbReportHead.AddItem "LedgerBalances Publishers Distributors-NonZero"
    cmbReportHead.AddItem "LedgerBalances ALL"
    cmbReportHead.AddItem "STOCK STATEMENT"
    cmbReportHead.AddItem "STOCK REPORT"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW
        For i = 0 To flxReport.COLS - 1
            flxReport.ColHidden(i) = False
        Next
    End If
    If KeyCode = 191 And Shift = vbCtrlMask Then  '/ hide one
        flxReport.ColHidden(flxReport.Col) = True
        If flxReport.Col < flxReport.COLS - 2 Then flxReport.Col = flxReport.Col + 1
    End If
End Sub

Private Sub cmbReportHead_Click()
    Select Case cmbReportHead.ListIndex
        Case 0: LedgerBalancesParties
        Case 1: LedgerBalancesPD
        Case 2: LedgerBalancesParties_NZ
        Case 3: LedgerBalancesPD_NZ
        Case 4: LedgerBalancesAll
        Case 5: StockStatement
        Case 6: StockReport
    End Select
End Sub

Private Sub flxReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = MyCAPTION & FlxSum(Me.flxReport)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then SaveGrid Me.flxReport
End Sub

Private Sub cmdSaveGrid_Click()
    flxReport.Row = 0
    If MsgBox("Do you wish to save the grid?", vbYesNo + vbQuestion) = vbYes Then
        mdiOne.CDlg.FileName = cmbReportHead.Text & " " & Format(Now, "DD-MMM-YY HHMMSS")
        mdiOne.CDlg.Filter = CompanyName & " Excel Report |*.xls"
        mdiOne.CDlg.ShowSave
        flxReport.FocusRect = flexFocusNone
        If mdiOne.CDlg.CancelError = False Then flxReport.SaveGrid mdiOne.CDlg.FileName, flexFileExcel, flexXLSaveFixedCells
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub LedgerBalancesParties()
    Dim LID, LNAME, LCity, LPhones As String
    Dim BAL As Long
    
    flxReport.COLS = 5: flxReport.ROWS = 1: flxReport.FixedRows = 1
    flxReport.ColAlignment(4) = flexAlignRightCenter
    flxReport.Clear
    
    CN(0).dbOpen "SELECT ID, NAME, City, Phones FROM appview_ALLACCOUNTS WHERE ID LIKE " & QT("C%") & " ORDER BY ID"
    
    flxReport.AddItem "ID" & vbTab & "NAME" & vbTab & "City" & vbTab & "Phones" & vbTab & "BALANCE", 0
    flxReport.ROWS = 1
    Do While Not CN(0).recs.EOF
        LID = CN(0).recs!id: LNAME = CN(0).recs!Name: LCity = CN(0).recs!City: LPhones = CN(0).recs!Phones
        CN(1).dbOpen "APPPROC_LEDGERBAL " & QT(LID)
        BAL = CN(1).recs!BAL
        flxReport.AddItem LID & vbTab & LNAME & vbTab & LCity & vbTab & LPhones & vbTab & Format(BAL, "#0.00 Dr; #0.00 Cr; 0.00")
        CN(0).recs.MoveNext
    Loop
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub LedgerBalancesPD()
    Dim LID, LNAME, LCity, LPhones As String
    Dim BAL As Long
    
    flxReport.COLS = 5: flxReport.ROWS = 1: flxReport.FixedRows = 1
    flxReport.ColAlignment(4) = flexAlignRightCenter
    flxReport.Clear
    
    CN(0).dbOpen "SELECT ID, NAME, City, Phones FROM appview_ALLACCOUNTS WHERE ID LIKE " & QT("[PD]%") & " ORDER BY ID"
    
    flxReport.AddItem "ID" & vbTab & "NAME" & vbTab & "City" & vbTab & "Phones" & vbTab & "BALANCE", 0
    flxReport.ROWS = 1
    Do While Not CN(0).recs.EOF
        LID = CN(0).recs!id: LNAME = CN(0).recs!Name: LCity = CN(0).recs!City: LPhones = CN(0).recs!Phones
        CN(1).dbOpen "APPPROC_LEDGERBAL " & QT(LID)
        BAL = CN(1).recs!BAL
        flxReport.AddItem LID & vbTab & LNAME & vbTab & LCity & vbTab & LPhones & vbTab & Format(BAL, "#0.00 Dr; #0.00 Cr; 0.00")
        CN(0).recs.MoveNext
    Loop
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub LedgerBalancesParties_NZ()
    Dim LID, LNAME, LCity, LPhones As String
    Dim BAL As Long
    
    flxReport.COLS = 5: flxReport.ROWS = 1: flxReport.FixedRows = 1
    flxReport.ColAlignment(4) = flexAlignRightCenter
    flxReport.Clear
    
    CN(0).dbOpen "SELECT ID, NAME, City, Phones FROM appview_ALLACCOUNTS WHERE ID LIKE " & QT("C%") & " ORDER BY ID"
    
    flxReport.AddItem "ID" & vbTab & "NAME" & vbTab & "City" & vbTab & "Phones" & vbTab & "BALANCE", 0
    flxReport.ROWS = 1
    Do While Not CN(0).recs.EOF
        LID = CN(0).recs!id: LNAME = CN(0).recs!Name: LCity = CN(0).recs!City: LPhones = CN(0).recs!Phones
        CN(1).dbOpen "APPPROC_LEDGERBAL " & QT(LID)
        BAL = CN(1).recs!BAL
        If BAL <> 0 Then flxReport.AddItem LID & vbTab & LNAME & vbTab & LCity & vbTab & LPhones & vbTab & Format(BAL, "#0.00 Dr; #0.00 Cr; 0.00")
        CN(0).recs.MoveNext
    Loop
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub LedgerBalancesPD_NZ()
    Dim LID, LNAME, LCity, LPhones As String
    Dim BAL As Long
    
    flxReport.COLS = 5: flxReport.ROWS = 1: flxReport.FixedRows = 1
    flxReport.ColAlignment(4) = flexAlignRightCenter
    flxReport.Clear
    
    CN(0).dbOpen "SELECT ID, NAME, City, Phones FROM appview_ALLACCOUNTS WHERE ID LIKE " & QT("[PD]%") & " ORDER BY ID"
    
    flxReport.AddItem "ID" & vbTab & "NAME" & vbTab & "City" & vbTab & "Phones" & vbTab & "BALANCE", 0
    flxReport.ROWS = 1
    Do While Not CN(0).recs.EOF
        LID = CN(0).recs!id: LNAME = CN(0).recs!Name: LCity = CN(0).recs!City: LPhones = CN(0).recs!Phones
        CN(1).dbOpen "APPPROC_LEDGERBAL " & QT(LID)
        BAL = CN(1).recs!BAL
        If BAL <> 0 Then flxReport.AddItem LID & vbTab & LNAME & vbTab & LCity & vbTab & LPhones & vbTab & Format(BAL, "#0.00 Dr; #0.00 Cr; 0.00")
        CN(0).recs.MoveNext
    Loop
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub LedgerBalancesAll()
    On Error Resume Next
    Dim LID, LNAME, LCity, LPhones As String
    Dim BAL As Long
    
    flxReport.COLS = 6: flxReport.ROWS = 1: flxReport.FixedRows = 1
    flxReport.ColAlignment(4) = flexAlignRightCenter
    flxReport.Clear
    
    CN(0).dbOpen "SELECT ID, NAME, City, Phones FROM appview_ALLACCOUNTS ORDER BY ID"
    
    flxReport.AddItem "ID" & vbTab & "NAME" & vbTab & "City" & vbTab & "Phones" & vbTab & "Dr" & vbTab & "Cr", 0
    flxReport.ROWS = 1
    
    CN(2).dbOpen "TRUNCATE TABLE FYCHANGE_PERSONAL", 1
    Do While Not CN(0).recs.EOF
        LID = CN(0).recs!id: LNAME = CN(0).recs!Name: LCity = CN(0).recs!City: LPhones = CN(0).recs!Phones
        CN(1).dbOpen "APPPROC_LEDGERBAL " & QT(LID)
        BAL = CN(1).recs!BAL
        CN(2).dbOpen "INSERT INTO FYCHANGE_PERSONAL VALUES (" & QT(LID) & ", " & BAL & ")", 1
        If BAL > 0 Then
            flxReport.AddItem LID & vbTab & LNAME & vbTab & LCity & vbTab & LPhones & vbTab & Format(BAL, "#0.00; #0.00; _") & vbTab & "_"
        Else
            flxReport.AddItem LID & vbTab & LNAME & vbTab & LCity & vbTab & LPhones & vbTab & "_" & vbTab & Format(BAL, "#0.00; #0.00; _")
        End If
        CN(0).recs.MoveNext
    Loop
    SubTotal 2
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub StockStatement()
    Dim LID, LNAME, LCity, LPhones As String
    Dim BAL As Long
    
    flxReport.COLS = 5: flxReport.ROWS = 1: flxReport.FixedRows = 1
    flxReport.ColAlignment(4) = flexAlignRightCenter
    flxReport.Clear
    
    CN(0).dbOpen "SELECT itemID, ISBN, itemNAME, PUBLISHERID, PUBLISHERNAME, INRPRICE, AVLBL AS STOCK, PDISCOUNT, 0 AS STOCKVAL_P, SDISCOUNT, 0 AS STOCKVAL_S FROM appview_StockExtended ORDER BY PUBLISHERID, ISBN"
    Set flxReport.DataSource = CN(0).recs
    For i = 1 To flxReport.ROWS - 1
        flxReport.TextMatrix(i, 8) = Val(flxReport.TextMatrix(i, 5)) * Val(flxReport.TextMatrix(i, 6)) * (1 - (CircularDiscount(flxReport.TextMatrix(i, 7)) / 100))
        flxReport.TextMatrix(i, 10) = Val(flxReport.TextMatrix(i, 5)) * Val(flxReport.TextMatrix(i, 6)) * (1 - (CircularDiscount(flxReport.TextMatrix(i, 9)) / 100))
    Next
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub StockReport()
    CN(0).dbOpen "EXEC APPPROC_STOCKREPORT"
    Set flxReport.DataSource = CN(0).recs
    flxReport.MergeCells = flexMergeRestrictAll
    flxReport.MergeCol(1) = True: flxReport.MergeCol(6) = True
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub SubTotal(ByVal TotalOnCols As Integer)
    flxReport.SubtotalPosition = flexSTBelow
    If TotalOnCols >= 2 Then flxReport.SubTotal flexSTSum, -1, flxReport.COLS - 1, "#,##0.00", , , True, "Total"
    If TotalOnCols >= 1 Then flxReport.SubTotal flexSTSum, -1, flxReport.COLS - 2, "#,##0.00", , , True, "Total"
End Sub

