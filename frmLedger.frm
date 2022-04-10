VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Ledger..."
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12885
   Icon            =   "frmLedger.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   12885
   Begin MSComCtl2.DTPicker DT1 
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   330
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "ddd, dd-MMM-yy"
      Format          =   80019459
      CurrentDate     =   38675
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   45
      TabIndex        =   9
      Top             =   660
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ledger..."
      TabPicture(0)   =   "frmLedger.frx":114DA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Print..."
      TabPicture(1)   =   "frmLedger.frx":114F6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   8025
         Left            =   -74910
         TabIndex        =   11
         Top             =   330
         Width           =   12645
         Begin VSPrinter8LibCtl.VSPrinter vp 
            Height          =   7860
            Left            =   75
            TabIndex        =   12
            Top             =   120
            Width           =   12525
            _cx             =   22093
            _cy             =   13864
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            MousePointer    =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoRTF         =   -1  'True
            Preview         =   -1  'True
            DefaultDevice   =   0   'False
            PhysicalPage    =   -1  'True
            AbortWindow     =   -1  'True
            AbortWindowPos  =   0
            AbortCaption    =   "Printing..."
            AbortTextButton =   "Cancel"
            AbortTextDevice =   "on the %s on %s"
            AbortTextPage   =   "Now printing Page %d of"
            FileName        =   ""
            MarginLeft      =   1440
            MarginTop       =   1440
            MarginRight     =   1440
            MarginBottom    =   1440
            MarginHeader    =   0
            MarginFooter    =   0
            IndentLeft      =   0
            IndentRight     =   0
            IndentFirst     =   0
            IndentTab       =   720
            SpaceBefore     =   0
            SpaceAfter      =   0
            LineSpacing     =   100
            Columns         =   1
            ColumnSpacing   =   180
            ShowGuides      =   2
            LargeChangeHorz =   300
            LargeChangeVert =   300
            SmallChangeHorz =   30
            SmallChangeVert =   30
            Track           =   0   'False
            ProportionalBars=   -1  'True
            Zoom            =   44.5075757575758
            ZoomMode        =   3
            ZoomMax         =   400
            ZoomMin         =   10
            ZoomStep        =   25
            EmptyColor      =   -2147483636
            TextColor       =   0
            HdrColor        =   0
            BrushColor      =   0
            BrushStyle      =   0
            PenColor        =   0
            PenStyle        =   0
            PenWidth        =   0
            PageBorder      =   0
            Header          =   ""
            Footer          =   ""
            TableSep        =   "|;"
            TableBorder     =   7
            TablePen        =   0
            TablePenLR      =   0
            TablePenTB      =   0
            NavBar          =   3
            NavBarColor     =   -2147483633
            ExportFormat    =   0
            URL             =   ""
            Navigation      =   3
            NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
            AutoLinkNavigate=   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7950
         Left            =   75
         TabIndex        =   10
         Top             =   330
         Width           =   12585
         Begin VB.TextBox txtBalance 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
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
            Left            =   960
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   7560
            Width           =   2325
         End
         Begin VB.TextBox txtBalanceAfter10Years 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
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
            Left            =   3660
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   7560
            Width           =   2325
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
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
            Left            =   9345
            TabIndex        =   4
            ToolTipText     =   "Adds New Record"
            Top             =   7545
            Width           =   1530
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   150
            Width           =   7605
         End
         Begin VB.TextBox txtID 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   510
            TabIndex        =   0
            Top             =   150
            Width           =   1275
         End
         Begin VB.CommandButton cmdSelectID 
            DownPicture     =   "frmLedger.frx":11512
            Height          =   315
            Left            =   9405
            Picture         =   "frmLedger.frx":11855
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
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
            Left            =   10935
            TabIndex        =   5
            Top             =   7545
            Width           =   1530
         End
         Begin VSFlex8UCtl.VSFlexGrid flxLedger 
            Height          =   6990
            Left            =   105
            TabIndex        =   3
            Top             =   510
            Width           =   12375
            _cx             =   21828
            _cy             =   12330
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
            MousePointer    =   1
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   12632319
            ForeColorSel    =   64
            BackColorBkg    =   12632256
            BackColorAlternate=   16777215
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
            Rows            =   1
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
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   0
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
            WallPaperAlignment=   10
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3390
            TabIndex        =   17
            Top             =   7620
            Width           =   135
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Balance:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   14
            Top             =   7620
            Width           =   840
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "ID:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   13
            Top             =   180
            Width           =   270
         End
      End
   End
   Begin MSComCtl2.DTPicker DT2 
      Height          =   315
      Left            =   10740
      TabIndex        =   7
      Top             =   330
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "ddd, dd-MMM-yy"
      Format          =   80019459
      CurrentDate     =   38675
   End
   Begin VB.Label Label2 
      Caption         =   "TO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10170
      TabIndex        =   16
      Top             =   330
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "FROM:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   75
      TabIndex        =   15
      Top             =   330
      Width           =   915
   End
   Begin VB.Label lblLedgerHead 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6390
      TabIndex        =   8
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(5) As New clsData
Private MyCAPTION As String
Dim CustId As String
Dim MaxValidRow As Integer

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    Me.Move 0, 0
    MyCAPTION = "LEDGER"
    Me.Caption = MyCAPTION
    DT1.Value = CDate("1-Apr-" & Left(FinancialYear(Now), 2))
    DT2.Value = Now
    txtID.Text = XID
    SSTab1.Tab = 0
End Sub

Private Sub DT1_Change()
    txtID_Change
End Sub

Private Sub DT2_Change()
    txtID_Change
End Sub

Private Sub cmdSelectID_Click()
    frmShow.Init "SELECT * FROM appview_ALLACCOUNTS ORDER BY 2"
    If sArray(0) <> "" Then
        txtID.Text = sArray(0)
        txtName.Text = sArray(1)
    End If
    CustId = txtID.Text
End Sub

Private Sub txtID_Change()
    On Error Resume Next
    CustId = txtID.Text: MaxValidRow = -1
    flxLedger.BackColorFixed = vbWhite
    SetflxLedgerFont ("[LEDGER_Render_Section_C]")
    
    lblLedgerHead.Caption = "": txtName.Text = "": txtBalance.Text = 0
    CN(0).dbOpen "SELECT * FROM appview_ALLACCOUNTS WHERE ID=" & QT(CustId), 1
    If CN(0).recs.RecordCount = 1 Then
        txtName.Text = Trim(CN(0).recs!Name) & ", " & Trim(CN(0).recs!City) & ", " & Trim(CN(0).recs!State)
        lblLedgerHead.Caption = "Account of " & txtName.Text
    End If
    
    If CustId <> "" Then
        CN(0).dbOpen "appproc_Ledger " & QT(CustId) & "," & QT(Format(DT1.Value, "dd-MMM-yy")) & "," & QT(Format(DT2.Value, "dd-MMM-yy")), 1
        Set flxLedger.DataSource = CN(0)
    End If
    
    txtBalance.Text = Format(WhatIsLedgerBalance(CustId, DT2.Value), "##,##0.00 Dr; ##,##0.00 Cr; NIL")
    txtBalanceAfter10Years.Text = Format(WhatIsLedgerBalance(CustId, DateAdd("Y", 10, DT2.Value)), "##,##0.00 Dr; ##,##0.00 Cr; NIL")
    flxLedger.SubtotalPosition = flexSTBelow
    flxLedger.SubTotal flexSTSum, -1, 2, "#,##0.00", , , True, "Total"
    flxLedger.SubTotal flexSTSum, -1, 3, "#,##0.00", , , True, "Total"
    flxLedger.ColWidth(0) = 1200: flxLedger.ColFormat(0) = "dd-MMM-yy"
    flxLedger.ColWidth(1) = 2800
    flxLedger.ColWidth(2) = 1450: flxLedger.ColFormat(2) = "#,##0.00"
    flxLedger.ColWidth(3) = 1450: flxLedger.ColFormat(3) = "#,##0.00"
    flxLedger.ColWidth(4) = 2050
    flxLedger.ColWidth(5) = 1000
    flxLedger.ColWidth(6) = 1000
    flxLedger.ColWidth(7) = 1000
    flxLedger.ColWidth(8) = 1000
    flxLedger.ColWidth(9) = 3400
    
    flxLedger.RowHeightMin = 400
    For Each hidecol In Array(5, 6, 7, 8)
        flxLedger.ColHidden(hidecol) = True
    Next
        
    'CHECK MAXVALIDROW
    For i = 1 To flxLedger.ROWS - 2
        If DateDiff("n", Now, CDate(Format(flxLedger.TextMatrix(i, 0), "dd-MM-yy"))) > 0 Then
            MaxValidRow = i
        End If
    Next

    X = Now
    RunningBalance
End Sub

Private Sub flxLedger_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = vbCtrlMask Then SaveGrid Me.flxLedger
    Me.Caption = FlxSum(Me.flxLedger)
End Sub

Private Sub flxLedger_DblClick()
    ShowRelatedMemo
End Sub

Private Sub flxLedger_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ShowRelatedMemo
End Sub

Private Sub cmdPrint_Click()
    With vp
        .AutoRTF = True
        PprDim = Split(mdiOne.sckGo.GReadINI("[LEDGER_PaperHeight/Width/Left/Right/Top/Bottom]"), ":")
        vp.PaperSize = pprFanfoldStdGerman
        vp.PhysicalPage = False
        vp.PaperHeight = PprDim(0): vp.PaperWidth = PprDim(1)
        vp.MarginLeft = PprDim(2): vp.MarginRight = PprDim(3)
        vp.MarginTop = PprDim(4): vp.MarginBottom = PprDim(5)
        vp.Orientation = orPortrait
        .PenStyle = psSolid: .TrueType = ttBitmap
        
        .StartDoc
            SetFont ("[LEDGER_Render_Section_A]")
            .TextAlign = taCenterTop
            .Text = "{ \b\ul " & CompanyName & ", " & CompanyCity & ", " & CompanyState & " \ulnone\par }"
            
            SetFont ("[LEDGER_Render_Section_B]")
            .Text = "{ \b Account of " & txtName.Text & " \par for the period of " & Format(DT1.Value, "dd-MMM-yy") & " to " & Format(DT2.Value, "dd-MMM-yy") & " \par }"
            
            SetflxLedgerFont ("[LEDGER_Render_Section_C]")
            
            flxLedger.ColWidth(0) = 1400
            flxLedger.ColWidth(1) = 3200
            flxLedger.ColWidth(2) = 1850
            flxLedger.ColWidth(3) = 1850
            flxLedger.ColWidth(4) = 1950
            flxLedger.ColHidden(9) = True
            flxLedger.GridLines = flexGridNone
            flxLedger.GridLinesFixed = flexGridNone
                .RenderControl = flxLedger.hWnd
            flxLedger.GridLinesFixed = flexGridFlat
            flxLedger.GridLines = flexGridFlat
            flxLedger.ColWidth(0) = 1200
            flxLedger.ColWidth(1) = 2800
            flxLedger.ColWidth(2) = 1350
            flxLedger.ColWidth(3) = 1350
            flxLedger.ColWidth(4) = 1450
            flxLedger.ColHidden(9) = False
            
            .Text = ""
            SetFont ("[LEDGER_Render_Section_D]")
            .TextAlign = taLeftTop
            .Text = vbCrLf & "Current balance = Rs." & txtBalance.Text & "  /  Overall balance = Rs." & txtBalanceAfter10Years.Text
            .Text = vbCrLf & "Note: Cheque payments are subject to realisation."
            .TextAlign = taRightTop
            .Text = vbCrLf & "Authorised signatory"
        .EndDoc
        
        For i = 1 To vp.PageCount
            vp.TextAlign = taRightTop
            vp.StartOverlay i
            vp.CurrentY = 150
            vp.Text = "Page " & i & " of " & vp.PageCount
            vp.EndOverlay
        Next
    End With
    SSTab1.Tab = 1
End Sub

Private Sub cmdClose_Click()
    Me.ValidateControls: Unload Me
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If KeyCode = vbKeyUp Then txtID.Text = Left(txtID.Text, 1) + Format(Val(Right((txtID.Text), 4)) + 1, "0000")
        If KeyCode = vbKeyDown Then txtID.Text = Left(txtID.Text, 1) + Format(Val(Right((txtID.Text), 4)) - 1, "0000")
    End If
End Sub

Private Function flxLedgerSum(ByVal StartRow As Integer, ByVal StopRow As Integer, ByVal sumCol As Integer) As Double
    flxLedgerSum = Val(flxLedger.TextMatrix(StopRow, sumCol + 5))
End Function

Private Sub RunningBalance()
    'DATE COLORING, RUNNING BALANCE
    Dim BALANCE As Double
    BALANCE = 0
    For i = 2 To MaxValidRow
        For Each CumulativeCol In Array(4, 7, 8)
            flxLedger.TextMatrix(i, CumulativeCol) = Val(flxLedger.TextMatrix(i - 1, CumulativeCol)) + Val(flxLedger.TextMatrix(i, CumulativeCol))
        Next
        DoEvents
    Next
    For i = 1 To flxLedger.ROWS - 2
        flxLedger.TextMatrix(i, 4) = Format(Val(flxLedger.TextMatrix(i, 4)), "##,##0.00 Dr; ##,##0.00 Cr; NIL")
        DoEvents
    Next
End Sub

Private Sub ShowRelatedMemo()
    MemoRef = flxLedger.TextMatrix(flxLedger.Row, 5)
    MEMOTYPE = Left(MemoRef, 2)
    XREF = Val(Replace(MemoRef, MEMOTYPE, ""))
    Select Case UCase(MEMOTYPE)
        Case "SA": Sale.frm.LinkOpen
        Case "SR": SaleReturn.frm.LinkOpen
        Case "PU": Purchase.frm.LinkOpen
        Case "PR": PurchaseReturn.frm.LinkOpen
        Case "TI": StockTransferIN.frm.LinkOpen
        Case "TO": StockTransferOUT.frm.LinkOpen
        
        Case "RT": RCPT.frm.LinkOpen
        Case "PT": PYMT.frm.LinkOpen
        Case "VR": frmVouchers.LinkOpen
    End Select
End Sub

Public Sub LinkOpen()
    Me.SetFocus
    txtID.Text = XID
End Sub

Private Sub SetFont(S As String)
    vp.FontName = ReadFont(S, 0)
    vp.FontSize = ReadFont(S, 1)
    vp.FontBold = ReadFont(S, 2)
    vp.FontItalic = ReadFont(S, 3)
End Sub

Private Sub SetflxLedgerFont(S As String)
    flxLedger.FontName = ReadFont(S, 0)
    flxLedger.FontSize = ReadFont(S, 1)
    flxLedger.FontBold = ReadFont(S, 2)
    flxLedger.FontItalic = ReadFont(S, 3)
End Sub

