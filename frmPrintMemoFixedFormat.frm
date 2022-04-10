VERSION 5.00
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmPrintMemoFixedFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BILL PRINTING..."
   ClientHeight    =   11280
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmPrintMemoFixedFormat.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   9855
   Begin VB.PictureBox Picture2 
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9810
      TabIndex        =   17
      Top             =   0
      Width           =   9870
      Begin VB.CommandButton cmdMail 
         Height          =   330
         Left            =   7785
         Picture         =   "frmPrintMemoFixedFormat.frx":4E0E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   15
         Width           =   465
      End
      Begin VB.TextBox txtEmail 
         Height          =   360
         Left            =   1740
         TabIndex        =   18
         Top             =   15
         Width           =   6030
      End
      Begin VB.Label Label1 
         Caption         =   "Mail this bill to >>> "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   30
         TabIndex        =   20
         Top             =   75
         Width           =   1770
      End
      Begin VSPDF8LibCtl.VSPDF8 vspdf 
         Left            =   9360
         Top             =   -30
         Author          =   ""
         Creator         =   ""
         Title           =   ""
         Subject         =   ""
         Keywords        =   ""
         Compress        =   3
      End
   End
   Begin VB.Frame Frame1 
      Height          =   10830
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   9855
      Begin VB.PictureBox Picture1 
         Height          =   10605
         Left            =   8340
         ScaleHeight     =   10545
         ScaleWidth      =   1365
         TabIndex        =   1
         Top             =   150
         Width           =   1425
         Begin VB.ComboBox cmbPrintFormat 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "frmPrintMemoFixedFormat.frx":4FE6
            Left            =   0
            List            =   "frmPrintMemoFixedFormat.frx":4FE8
            TabIndex        =   12
            Text            =   "cmbMemoFormat"
            Top             =   0
            Width           =   1350
         End
         Begin VB.CommandButton cmdTransportForwarding 
            Caption         =   "&Forwarding"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2055
            Width           =   1260
         End
         Begin VB.TextBox txtDBRef 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   135
            TabIndex        =   10
            Top             =   10185
            Width           =   1425
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   45
            TabIndex        =   9
            Top             =   2520
            Width           =   1260
         End
         Begin VB.CommandButton cmdQuit 
            Caption         =   "&Quit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   45
            TabIndex        =   8
            Top             =   2985
            Width           =   1260
         End
         Begin VB.ComboBox cmbPaperSizes 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "frmPrintMemoFixedFormat.frx":4FEA
            Left            =   0
            List            =   "frmPrintMemoFixedFormat.frx":4FEC
            TabIndex        =   7
            Text            =   "cmbPaperSizes"
            Top             =   330
            Width           =   1350
         End
         Begin VB.TextBox txtPaperSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   15
            TabIndex        =   6
            Text            =   "0"
            Top             =   675
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.TextBox txtPaperSize 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   690
            TabIndex        =   5
            Text            =   "0"
            Top             =   675
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.CheckBox chkPNotes 
            Caption         =   "PNotes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   375
            TabIndex        =   4
            Top             =   9795
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.CommandButton cmdFooterPad 
            Caption         =   "<<"
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
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   1500
            Width           =   435
         End
         Begin VB.CommandButton cmdFooterPad 
            Caption         =   ">>"
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
            Index           =   1
            Left            =   930
            TabIndex        =   2
            Top             =   1500
            Width           =   435
         End
         Begin VB.Label lblPaperSize 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PaperSize"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   195
            TabIndex        =   14
            Top             =   1125
            Width           =   885
         End
         Begin VB.Label lblFooterPad 
            Alignment       =   2  'Center
            Caption         =   "FooterPad"
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
            Left            =   420
            TabIndex        =   13
            Top             =   1560
            Width           =   555
         End
      End
      Begin VSPrinter8LibCtl.VSPrinter vp 
         Height          =   10605
         Left            =   60
         TabIndex        =   15
         Top             =   150
         Width           =   8250
         _cx             =   14552
         _cy             =   18706
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         PhysicalPage    =   0   'False
         PalettePicture  =   "frmPrintMemoFixedFormat.frx":4FEE
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   1440
         MarginTop       =   720
         MarginRight     =   1440
         MarginBottom    =   360
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
         Zoom            =   64
         ZoomMode        =   0
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
         NavBar          =   1
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
         Begin VSFlex8UCtl.VSFlexGrid flxItems 
            Height          =   1530
            Left            =   1230
            TabIndex        =   16
            Top             =   5310
            Visible         =   0   'False
            Width           =   5520
            _cx             =   9737
            _cy             =   2699
            Appearance      =   0
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
            BackColorFixed  =   16777215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   12632319
            ForeColorSel    =   64
            BackColorBkg    =   -2147483624
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   8454143
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   4
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
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
            AutoResize      =   0   'False
            AutoSizeMode    =   1
            AutoSearch      =   1
            AutoSearchDelay =   2
            MultiTotals     =   0   'False
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   2
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   0   'False
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
   End
End
Attribute VB_Name = "frmPrintMemoFixedFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'  PRINT BILL * FULL SIZE PAGE * WITH FIXED HEADER/FOOTER POSITIONS
'**********************************************************************
Private CN(5) As New clsData

Dim Ref, CurrY As Long
Dim TMAIN, Support, MemoHeading, Salutation, MemoNo, ToPayMode As String
Dim Terms, pdfFileName As String
Dim ItemCodeCol, itemNameCol, ProducerIDCol, ProducerNameCol As Integer
Dim LedgerBalancePrinting As Boolean
Dim FooterPad As Integer

Private Sub Form_Load()
    On Error Resume Next
    FooterPad = Val(mdiOne.sckGo.GReadINI("[FooterPad]")): lblFooterPad.Caption = FooterPad

    ItemCodeCol = 2: itemNameCol = 3: ProducerIDCol = 5: ProducerNameCol = 6
    
    PprDim = Split(mdiOne.sckGo.GReadINI("[PaperHeight/Width/Left/Right/Top/Bottom]"), ":")
    'vp.PaperHeight = PprDim(0): vp.PaperWidth = PprDim(1)
    vp.MarginLeft = PprDim(2): vp.MarginRight = PprDim(3)
    'vp.MarginTop = PprDim(4): vp.MarginBottom = PprDim(5)
    
    vp.PenStyle = psSolid: vp.PenColor = vbBlack: vp.BrushColor = vbBlack: vp.PenWidth = 5
    vp.AbortWindow = True: vp.AutoRTF = True
    vp.Copies = 1: vp.Collate = colTrue
    
    DebugPageSizes
    RunningNews = mdiOne.sckGo.GReadINI("[RunningNews]", "[END]")
    LedgerBalancePrinting = mdiOne.sckGo.GReadINI("[LedgerBalancePrinting]")
    cmbPaperSizes.ListIndex = Val(mdiOne.sckGo.GReadINI("[PaperSizeNumber]"))
    vp.PaperSize = Val(cmbPaperSizes.Text)
    vp.PaperSize = pprUser
    vp.PhysicalPage = False
    'vp.PageBorder = pbBottom
    vp.Orientation = orPortrait
    
    cmbPrintFormat.Clear
    cmbPrintFormat.AddItem "Format A: GENERAL"       ' 0
    cmbPrintFormat.AddItem "Format B: CHALLAN"       ' 1
    cmbPrintFormat.AddItem "Format C: STOK TRFR"     ' 2
    cmbPrintFormat.AddItem "Format D: LIBR SUPP"     ' 3
    cmbPrintFormat.AddItem "Format E: PURC ORDR"     ' 4
    cmbPrintFormat.AddItem "Format F: PURC RETN"     ' 5
End Sub

Public Sub PrintIT(ByVal DBRef As Long, ByVal MType As String, ByVal MemoPrintFormat As Integer)
    Ref = DBRef
    Select Case UCase(MType)
        Case "SALE": TMAIN = "SMAIN": Support = "SALE"
        Case "SALERETURN": TMAIN = "SRETURNMAIN": Support = "SALERETURN"
        Case "PURCHASE": TMAIN = "PMAIN": Support = "PURCHASE"
        Case "PURCHASERETURN": TMAIN = "PRETURNMAIN": Support = "PURCHASERETURN"
        Case "TOUT": TMAIN = "TOUTMAIN": Support = "TOUT"
        Case "TIN": TMAIN = "TINMAIN": Support = "TIN"
    End Select
    If MemoPrintFormat <= cmbPrintFormat.ListCount - 1 Then
        cmbPrintFormat.ListIndex = MemoPrintFormat   'THIS RENDERS THE BILL
    End If
    Me.Show vbModal
End Sub

Private Sub cmbPrintFormat_Click()
    cmbPaperSizes_Click
End Sub

Private Sub cmbPaperSizes_Click()
    vp.PaperSize = Val(cmbPaperSizes.Text)
    lblPaperSize.Caption = " W:" & Format(vp.PageWidth / 1440, "#00.00") & " H:" & Format(vp.PageHeight / 1440, "#00.00")
    msgUITS cmbPaperSizes.ListIndex
    RenderMemo cmbPrintFormat.ListIndex
End Sub

Private Sub lblPaperSize_DblClick()
    vp.PhysicalPage = Not vp.PhysicalPage
    If vp.PhysicalPage = True Then
        lblPaperSize.FontBold = True
    Else
        lblPaperSize.FontBold = False
    End If
    lblPaperSize.Caption = " W:" & Format(vp.PageWidth / 1440, "#00.00") & " H:" & Format(vp.PageHeight / 1440, "#00.00")
End Sub

Private Sub chkDouble_Click()
    RenderMemo cmbPrintFormat.ListIndex
End Sub

Private Sub cmdFooterPad_Click(Index As Integer)
    If Index = 0 Then
        FooterPad = FooterPad - 200
    Else
        FooterPad = FooterPad + 200
    End If
    lblFooterPad.Caption = FooterPad
    RenderMemo cmbPrintFormat.ListIndex
    vp.PreviewPage = vp.PageCount
End Sub

Private Sub cmdMail_Click()
On Error Resume Next
    vspdf.ConvertDocument vp, App.Path & "\" & pdfFileName
    If InStr(1, txtEmail.Text, "@", vbTextCompare) > 0 Then
        mdiOne.MAPISess.SignOn
        mdiOne.MAPIMsg.SessionID = mdiOne.MAPISess.SessionID
        mdiOne.MAPIMsg.Compose
        mdiOne.MAPIMsg.RecipAddress = Trim(txtEmail.Text)
        mdiOne.MAPIMsg.ResolveName
        mdiOne.MAPIMsg.MsgSubject = CompanyName & ": " & sArray(0)
        mdiOne.MAPIMsg.MsgNoteText = "Sir, " & vbCrLf & "Kindly acknowledge the recepit of this message at my email address." & vbCrLf & "Thanking you." & vbCrLf & vbCrLf & CompanyName & vbCrLf
        mdiOne.MAPIMsg.AttachmentPathName = App.Path & "\" & pdfFileName
        mdiOne.MAPIMsg.Send False
        mdiOne.MAPISess.SignOff
    Else
        msgUITS "Invalid address"
    End If
End Sub

Public Sub RenderMemo(ByVal printFormatOption As Integer)
    CN(0).dbOpen "Select * from " & TMAIN & " where DBRef=" & Ref, 0
    If CN(0).recs.RecordCount = 1 Then
        CustNo = CN(0).recs!id: txtEmail.Text = CN(0).recs!Email: pdfFileName = Support & "_" & Format(Ref, "0000") & ".pdf"
        Select Case CN(0).recs!ToPayMode
            Case 0: ToPayMode = "PAID-FULL"
            Case 1: ToPayMode = "PAID-HALF"
            Case 2: ToPayMode = "PAID-ZERO"
            Case 3: ToPayMode = "TOPAY-FULL"
            Case 4: ToPayMode = "TOPAY-HALF"
            Case 5: ToPayMode = "TOPAY-ZERO"
        End Select
        SupportSQL = "SELECT Serial as Sl, ItemID, ItemCode as CODE, ItemName, ProducerName as Mfg, Packing, Qty, MRP, SRP, Gross, aDiscAmt as 'Disc', Amount FROM " & Support & " WHERE DBRefX=" & Ref & "  ORDER BY SERIAL"
        CN(1).dbOpen SupportSQL, 1: Set flxItems.DataSource = CN(1).recs
        CN(2).dbOpen "SELECT DISTINCT * FROM CURRENCY WHERE CURRENCY<>" & QT("X"), 1
        CN(3).dbOpen "SELECT * FROM PERSONAL WHERE ID = " & QT(CN(0).recs!id), 1
        
        Select Case Support
            Case "SALE":
                Select Case UCase(CN(0).recs!Status)
                    Case "ORDER": MemoHeading = "ORDER FORM/REQUISITION FORM": MemoNo = "Memo# " & CompanyBillInitial & "/PF/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                    Case "CHALLAN": MemoHeading = "     APPROVAL CHALLAN": MemoNo = "Inv# " & CompanyBillInitial & "/CH/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                    Case "CASH": MemoHeading = "    CASH INVOICE": MemoNo = "Inv# " & CompanyBillInitial & "/CS/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                    Case "CREDIT": MemoHeading = "  CREDIT INVOICE": MemoNo = "Inv# " & CompanyBillInitial & "/CR/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                End Select
                Salutation = "To, "
                Terms = SaleInvoiceTerms
            Case "SALERETURN":
                MemoHeading = "SALES RETURN/ CREDIT NOTE": MemoNo = "Memo# " & CompanyBillInitial & "/SR/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                Salutation = "To, "
                Terms = "K.A.:" & vbCrLf & "Validity of sales return is subject to acceptance of items by the concerned Producers. If the Producer does not accept the return, the same will be debited to your account."
            Case "PURCHASE":
                MemoHeading = "PURCHASE ORDER": MemoNo = "Memo# " & CompanyBillInitial & "/PO/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                Salutation = "To, "
                Terms = ""
            Case "PURCHASERETURN":
                MemoHeading = "PURCHASE RETURN": MemoNo = "Memo# " & CompanyBillInitial & "/PR/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                Salutation = "To, "
                Terms = ""
                RunningNews = ""
            Case "TOUT":
                MemoHeading = "        STK TRANSFER": MemoNo = "Memo# " & CompanyBillInitial & "/TO/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                Salutation = "To, "
                Terms = ""
            Case "TIN":
                Select Case UCase(CN(0).recs!Status)
                    Case "ORDER":   MemoHeading = "ORDER": MemoNo = "Memo# " & CompanyBillInitial & "TI" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                                    Salutation = "To, "
                    Case "CHALLAN": MemoHeading = "        STK TRANSFER IN-CH": MemoNo = "Memo# " & CompanyBillInitial & "CH" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                                    Salutation = "From, "
                    Case "CASH":    MemoHeading = "        STK TRANSFER IN-CA": MemoNo = "Inv# " & CompanyBillInitial & "CS" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                                    Salutation = "From, "
                    Case "CREDIT":  MemoHeading = "        STK TRANSFER IN-AC": MemoNo = "Inv# " & CompanyBillInitial & "CR" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                                    Salutation = "From, "
                End Select
                Terms = ""
        End Select
        
        ' ****************************************
        ' MEASURE BLOCKS
        ' ****************************************
            vp.StartDoc: vp.MarginTop = 200: vp.MarginBottom = 200
                    i = 1
                    blockATop = 200
                    vp.TextAlign = taLeftTop: vp.CurrentY = blockATop: vp.Text = "{\b " & "CIN:U80302JH2007PTC012887 & PAN: AADCB4036F" & " \b}"
                    vp.TextAlign = taRightTop: vp.CurrentY = blockATop: vp.Text = "{ " & "Page " & i & " of " & vp.PageCount & " \par\b [" & Ref & "] }"
                    
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_A]": Render_Section_A printFormatOption, i         ' Memo Heading
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_B]": Render_Section_B printFormatOption, i         ' Company name and heading
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_C]": Render_Section_C printFormatOption, i         ' Company address
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_D]": Render_Section_D printFormatOption, i         ' Invoice, transportation details
                    'vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_E]": Render_Section_E printFormatOption, i        ' Conversion rates
                    
                    MarginTop = vp.CurrentY
                    blockCTop = vp.CurrentY
                    
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_G]": Render_Section_G printFormatOption, i         ' discounts, frieght, postage, net amount
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_H]": Render_Section_H printFormatOption, i         ' amount in words, tax exemption statement, declaration
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_I]": Render_Section_I printFormatOption, i         ' Comments
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_J]": Render_Section_J printFormatOption, i         ' RunningNews
                    vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_K]": Render_Section_K printFormatOption, i         ' Terms, Prepared by, Checked by, Signatory
    
                    MarginBottom = vp.CurrentY - blockCTop + FooterPad
            vp.EndDoc
    
        ' ****************************************
        ' END OF MEASURE BLOCKS
        ' ****************************************
        
        '**********************************************************************
        '                   START OF PAGE RENDERING
        '**********************************************************************
        vp.Clear
        vp.MarginTop = MarginTop: vp.MarginBottom = MarginBottom
        
        vp.StartDoc
            SetFlxFont "[Render_Section_F]": Render_Section_F printFormatOption, 0      ' Books details
        vp.EndDoc
        
        i = vp.PageCount
        For i = 1 To vp.PageCount
            vp.StartOverlay i
                vp.TextAlign = taLeftTop: vp.CurrentY = blockATop: vp.Text = "{\b " & "CIN:U80302JH2007PTC012887 & PAN: AADCB4036F" & " \b}"
                vp.TextAlign = taRightTop: vp.CurrentY = blockATop: vp.Text = "{ " & "Page " & i & " of " & vp.PageCount & " \par\b [" & Ref & "] }"
                
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_A]": Render_Section_A printFormatOption, i         ' Memo Heading
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_B]": Render_Section_B printFormatOption, i         ' Company name and heading
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_C]": Render_Section_C printFormatOption, i         ' Company address
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_D]": Render_Section_D printFormatOption, i         ' Invoice, transportation details
                'vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_E]": Render_Section_E printFormatOption, i        ' Conversion rates
                
                '<<< SetFlxFont "[Render_Section_F]": Render_Section_F   >>> Rendered first before this loop
           
                vp.CurrentY = vp.PageHeight - MarginBottom: SetFont "[Render_Section_G]": Render_Section_G printFormatOption, i        ' discounts, frieght, postage, net amount
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_H]": Render_Section_H printFormatOption, i         ' amount in words, tax exemption statement, declaration
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_I]": Render_Section_I printFormatOption, i         ' Comments
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_J]": Render_Section_J printFormatOption, i         ' RunningNews
                vp.CurrentY = vp.CurrentY + 150: SetFont "[Render_Section_K]": Render_Section_K printFormatOption, i         ' Terms, Prepared by, Checked by, Signatory
            vp.EndOverlay
        Next
        '**********************************************************************
        '                   END OF PAGE RENDERING
        '**********************************************************************
    End If
End Sub

Public Sub Render_Section_A(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    With vp
        .TextAlign = taCenterTop
        .Text = MemoHeading
        .TextAlign = taLeftTop
    End With
End Sub

Public Sub Render_Section_B(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    Select Case printFormatOption
        Case 2:
            With vp
                .StartTable
                    .TableBorder = tbNone
                    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "6.00in"
                    .TableCell(tcColAlign, , 1) = taLeftTop
                    .FontSize = 14
                    .TableCell(tcText, 1, 1) = CompanyName & ", " & CompanyDivision
                .EndTable
            End With
        Case Else
            With vp
                Y = .CurrentY
             '   .DrawPicture mdiOne.ImgList.ListImages(1).Picture, .PageWidth / 1.7 - 415, .CurrentY + 270, 837, 1000
                vp.DrawPicture mdiOne.ImgList.ListImages(1).Picture, .MarginLeft, .CurrentY, 837, 1000
                .CurrentY = Y + 200
                .StartTable
                    .TableBorder = tbNone
                    .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "0.60in"
                    .TableCell(tcColWidth, , 2) = "6.00in"
                    .TableCell(tcColAlign, , 2) = taCenterLeft
                    .TableCell(tcText, 1, 2) = CompanyName
                .EndTable
            End With
    End Select
End Sub

Public Sub Render_Section_C(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    Select Case printFormatOption
        Case 2:
        Case 4:
            With vp
                .StartTable
                    .TableBorder = tbNone
                    .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "0.60in"
                    .TableCell(tcColWidth, , 2) = "6.00in"
                    .TableCell(tcColWidth, , 3) = "1.10in"
                    .TableCell(tcColAlign, , 2) = taLeftTop
                    .TableCell(tcColAlign, , 3) = taRightTop
                    .TableCell(tcText, 1, 2) = AboutCompany & vbCrLf & CompanyAddr0 & vbCrLf & CompanyPhone & ", " & CompanyFax
                    .TableCell(tcText, 1, 3) = "Div.:" & CompanyDivision
                .EndTable
            End With
        Case Else:
            With vp
                .StartTable
                    .TableBorder = tbNone
                    .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "0.60in"
                    .TableCell(tcColWidth, , 2) = "5.8.00in"
                    .TableCell(tcColAlign, , 2) = taCenterLeft
                    .TableCell(tcText, 1, 2) = "{" & AboutCompany & vbCrLf & "\par \b (" & CompanyDivision & " DIVISION)    CIN No.: U80302JH2007PTC012887, PAN: AADCB4036F, GSTIN: .............." & "}"
                .EndTable
                .StartTable
                    .TableBorder = tbNone
                    .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "3.85in": .TableCell(tcColWidth, , 2) = "3.85in"
                    .TableCell(tcColAlign, , 1) = taLeftTop: .TableCell(tcColAlign, , 2) = taRightTop
                    .TableCell(tcText, 1, 1) = CompanyAddr1 & vbCrLf & CompanyAddr2 & vbCrLf & CompanyPhone & "," & CompanyFax & vbCrLf & CompanyEmail & vbCrLf & "PAN: " & CompanyPAN & " | GSTIN: " & CompanyGSTIN
                .EndTable
            End With
    End Select
End Sub

Public Sub Render_Section_D(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    With vp
    Select Case printFormatOption
        Case 1:
                .StartTable
                    .TableBorder = tbAll: .TableCell(tcRowSpaceAfter) = 20
                    .TableCell(tcCols) = 3: .TableCell(tcRows) = 4: .TableCell(tcColBorder, , 2, , 2) = 0
                    .TableCell(tcRowSpan, 1, 1) = 4
                    .TableCell(tcColWidth, , 1) = "4.0in": .TableCell(tcColWidth, , 2) = "2.3in": .TableCell(tcColWidth, , 3) = "1.4in"
                    .TableCell(tcText, 1, 1) = "{" & Salutation & "\b " & CN(0).recs!Name & "\b0(" & CN(0).recs!id & ")" & "\par " & CN(0).recs!Address & "\par " & CN(0).recs!City & ". Ph: " & CN(0).recs!Phones & "\par " & "PAN: " & CN(3).recs!PAN & "| GSTIN: " & CN(3).recs!GSTIN & "}"
                    
                    .TableCell(tcText, 1, 2) = "{\b " & MemoNo & " }"
                    .TableCell(tcText, 1, 3) = "Date: " & Format(CN(0).recs!DBDate, "dd-MM-yy")
                    .TableCell(tcText, 2, 2) = "GR/RR No.: " & CN(0).recs!GRNo
                    If Trim(CN(0).recs!GRNo) <> "_" Then
                        '.TableCell(tcText, 2, 3) = "Date: " & Format(CN(0).recs!GRDate, "dd-MM-yy")
                        .TableCell(tcText, 2, 3) = ToPayMode
                    End If
                    .TableCell(tcText, 3, 2) = "Delivery By: " & CN(0).recs!TName
                    .TableCell(tcText, 3, 3) = "Bundles: " & CN(0).recs!BundleCount
                .EndTable
        Case 2:
                .StartTable
                    .TableBorder = tbAll: .TableCell(tcRowSpaceAfter) = 20
                    .TableCell(tcCols) = 3: .TableCell(tcRows) = 4: .TableCell(tcColBorder, , 2, , 2) = 0
                    .TableCell(tcRowSpan, 1, 1) = 2
                    .TableCell(tcColWidth, , 1) = "4.0in": .TableCell(tcColWidth, , 2) = "2.3in": .TableCell(tcColWidth, , 3) = "1.4in"
                    .TableCell(tcText, 1, 1) = "{" & Salutation & "\b " & CN(0).recs!Name & "\b0(" & CN(0).recs!id & ")" & "\par " & CN(0).recs!Address & "\par " & CN(0).recs!City & ". Ph: " & CN(0).recs!Phones & "\par " & "PAN: " & CN(3).recs!PAN & "| GSTIN: " & CN(3).recs!GSTIN & "}"
                    
                    .TableCell(tcText, 1, 2) = "{\b " & MemoNo & " }"
                    .TableCell(tcText, 1, 3) = "Date: " & Format(CN(0).recs!DBDate, "dd-MM-yy")
                    .TableCell(tcText, 2, 2) = "GR/RR No.: " & CN(0).recs!GRNo
                    If Trim(CN(0).recs!GRNo) <> "_" Then
                        '.TableCell(tcText, 2, 3) = "Date: " & Format(CN(0).recs!GRDate, "dd-MM-yy")
                        .TableCell(tcText, 2, 3) = ToPayMode
                    End If
                    .TableCell(tcText, 3, 1) = CN(0).recs!InvRef
                    .TableCell(tcText, 3, 2) = "Delivery By: " & CN(0).recs!TName
                    .TableCell(tcText, 3, 3) = "Bundles: " & CN(0).recs!BundleCount
                .EndTable
        Case 4:
                .StartTable
                    .TableBorder = tbAll: .TableCell(tcRowSpaceAfter) = 20
                    .TableCell(tcCols) = 3: .TableCell(tcRows) = 4: .TableCell(tcColBorder, , 2, , 2) = 0
                    .TableCell(tcRowSpan, 1, 1) = 4
                    .TableCell(tcColWidth, , 1) = "4.0in": .TableCell(tcColWidth, , 2) = "2.3in": .TableCell(tcColWidth, , 3) = "1.4in"
                    .TableCell(tcText, 1, 1) = "{" & Salutation & "\b " & CN(0).recs!Name & "\b0(" & CN(0).recs!id & ")" & "\par " & CN(0).recs!Address & "\par " & CN(0).recs!City & ". Ph: " & CN(0).recs!Phones & "\par " & "PAN: " & CN(3).recs!PAN & "| GSTIN: " & CN(3).recs!GSTIN & "}"
                    
                    .TableCell(tcText, 1, 2) = "{\b " & MemoNo & " }"
                    .TableCell(tcText, 1, 3) = "Date: " & Format(CN(0).recs!DBDate, "dd-MM-yy")
                .EndTable
                .StartTable
                    .TableBorder = tbAll: .TableCell(tcRowSpaceAfter) = 20
                    .TableCell(tcCols) = 3: .TableCell(tcRows) = 2: .TableCell(tcColBorder, , 2, , 2) = 0
                    .TableCell(tcColWidth, , 1) = "4.0in": .TableCell(tcColWidth, , 2) = "2.3in": .TableCell(tcColWidth, , 3) = "1.4in"
                    .TableCell(tcRowHeight) = ".2in"
                    .TableCell(tcColSpan, 2, 1) = 3
                    .TableCell(tcText, 2, 1) = "Ship To: " & CN(0).recs!ShipAddress
                .EndTable
        Case Else:
                .StartTable
                    .TableBorder = tbAll: .TableCell(tcRowSpaceAfter) = 20
                    .TableCell(tcCols) = 3: .TableCell(tcRows) = 4: .TableCell(tcColBorder, , 2, , 2) = 0
                    .TableCell(tcRowSpan, 1, 1) = 4
                    .TableCell(tcColWidth, , 1) = "4.0in": .TableCell(tcColWidth, , 2) = "2.3in": .TableCell(tcColWidth, , 3) = "1.4in"
                    .TableCell(tcText, 1, 1) = "{" & Salutation & "\b " & CN(0).recs!Name & "\b0(" & CN(0).recs!id & ")" & "\par " & CN(0).recs!Address & "\par " & CN(0).recs!City & ". Ph: " & CN(0).recs!Phones & "\par " & "PAN: " & CN(3).recs!PAN & "| GSTIN: " & CN(3).recs!GSTIN & "}"
                    
                    .TableCell(tcText, 1, 2) = "{\b " & MemoNo & " }"
                    .TableCell(tcText, 1, 3) = "Date: " & Format(CN(0).recs!DBDate, "dd-MM-yy")
                    .TableCell(tcText, 2, 2) = "GR/RR No.: " & CN(0).recs!GRNo
                    If Trim(CN(0).recs!GRNo) <> "_" Then
                        '.TableCell(tcText, 2, 3) = "Date: " & Format(CN(0).recs!GRDate, "dd-MM-yy")
                        .TableCell(tcText, 2, 3) = ToPayMode
                    End If
                    .TableCell(tcText, 3, 2) = "Delivery By: " & CN(0).recs!TName
                    .TableCell(tcText, 3, 3) = "Bundles: " & CN(0).recs!BundleCount
                .EndTable
                
                .StartTable
                    .TableBorder = tbAll: .TableCell(tcRowSpaceAfter) = 20
                    .TableCell(tcCols) = 3: .TableCell(tcRows) = 2: .TableCell(tcColBorder, , 2, , 2) = 0
                    .TableCell(tcColWidth, , 1) = "4.0in": .TableCell(tcColWidth, , 2) = "2.3in": .TableCell(tcColWidth, , 3) = "1.4in"
                    
                    .TableCell(tcRowHeight) = ".2in"
                    
                    .TableCell(tcText, 1, 1) = "Order Ref.: " & CN(0).recs!OrderRef
                    .TableCell(tcText, 1, 2) = "Docs thru: " & CN(0).recs!PName
                    '.TableCell(tcText, 1, 3) = "Credit Days: "
                                        
                    .TableCell(tcColSpan, 2, 1) = 3
                    .TableCell(tcText, 2, 1) = "Ship To: " & CN(0).recs!ShipAddress
                .EndTable
    End Select
    End With
End Sub

Public Sub Render_Section_E(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    Select Case printFormatOption
        Case 10:
        Case Else:
            With vp
                .StartTable
                    cur = "Conversion rates: "
                    CN(2).recs.MoveFirst
                    Do Until CN(2).recs.EOF
                        cur = cur & "  " & CN(2).recs!Currency & "=" & Format(CN(2).recs!CurrPrice, "#0.00")
                        CN(2).recs.MoveNext
                    Loop
                    .TableBorder = tbNone
                    .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "7.7in"
                    .TableCell(tcColAlign, , 1) = taLeftTop
                    .TableCell(tcText, 1, 1) = cur
                .EndTable
            End With
    End Select
End Sub

Public Sub Render_Section_F(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    flxItems.BackColorFixed = vbWhite
    flxItems.GridLines = flexGridFlat: flxItems.GridColor = vbWhite
    flxItems.GridLinesFixed = flexGridFlatHorz: flxItems.GridColorFixed = vbYellow
    flxItems.BorderStyle = flexBorderFlat
    
    flxItems.ColAlignment(0) = flexAlignCenterTop       'SERIAL
    flxItems.ColAlignment(1) = flexAlignRightTop        'itemID
    flxItems.ColAlignment(2) = flexAlignLeftTop         'ItemCode
    flxItems.ColAlignment(3) = flexAlignLeftTop         'itemNAME
    flxItems.ColAlignment(4) = flexAlignLeftTop         'Mfg
    flxItems.ColAlignment(5) = flexAlignLeftTop         'Packing
    flxItems.ColAlignment(6) = flexAlignLeftTop         'Qty
    flxItems.ColAlignment(7) = flexAlignRightTop        'MRP
    flxItems.ColAlignment(8) = flexAlignRightTop        'SRP
    flxItems.ColAlignment(9) = flexAlignRightTop        'Gross
    flxItems.ColAlignment(10) = flexAlignRightTop       'Disc
    flxItems.ColAlignment(11) = flexAlignRightTop       'Amount
    
    For c = 0 To 11
        flxItems.ColWidth(c) = "0"
    Next
    
    flxItems.ColFormat(7) = "##,##0.00"      'MRP
    flxItems.ColFormat(8) = "##,##0.00"      'SRP
    flxItems.ColFormat(9) = "##,##0.00"      'Gross
    flxItems.ColFormat(10) = "##,##0.00"      'Disc
    flxItems.ColFormat(11) = "##,##0.00"      'Amount
   
    flxItems.AddItem ""
    flxItems.TextMatrix(flxItems.ROWS - 1, 6) = CN(0).recs!itemCount
    flxItems.TextMatrix(flxItems.ROWS - 1, 11) = Format(CN(0).recs!itemAmount, "##0.00")
    flxItems.Select flxItems.ROWS - 1, 0, flxItems.ROWS - 1, flxItems.COLS - 1
    flxItems.CellBorder RGB(0, 0, 125), 0, 1, 0, 1, 0, 0
    
    'Sl, ItemID, ItemCode as CODE, ItemName, ProducerName as Mfg, Packing, Qty, MRP, SRP, Gross, aDiscAmt as 'Disc', Amount
    Select Case printFormatOption
        Case 0: 'GENERAL
                flxItems.ColWidth(0) = "500"    'SERIAL
                flxItems.ColWidth(3) = "4500"   'itemNAME
                flxItems.ColWidth(6) = "800"    'Qty
                flxItems.ColWidth(7) = "900"    'MRP
                flxItems.ColWidth(8) = "900"    'SRP
                flxItems.ColWidth(9) = "1100"   'Gross
                flxItems.ColWidth(10) = "900"   'Disc
                flxItems.ColWidth(11) = "1100"  'Amount
        Case Else:
                flxItems.ColWidth(0) = "500"    'SERIAL
                flxItems.ColWidth(3) = "4500"   'itemNAME
                flxItems.ColWidth(6) = "800"    'Qty
                flxItems.ColWidth(7) = "900"    'MRP
                flxItems.ColWidth(8) = "900"    'SRP
                flxItems.ColWidth(9) = "1100"   'Gross
                flxItems.ColWidth(10) = "900"   'Disc
                flxItems.ColWidth(11) = "1100"  'Amount
    End Select
    vp.RenderControl = flxItems.hWnd
    Set flxItems.DataSource = Nothing
End Sub

Public Sub Render_Section_G(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    If PageNum = vp.PageCount Then
        With vp
            If printFormatOption <> 1 And printFormatOption <> 4 Then
                .CurrentY = .CurrentY + 50
                .StartTable
                    .TableBorder = tbBottom: .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "6.00in"
                    .TableCell(tcColAlign) = taLeftTop
                    .TableCell(tcText, 1, 1) = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
                .EndTable
                .StartTable
                    .TableBorder = tbTopBottom: .TableCell(tcCols) = 8: .TableCell(tcRows) = 2
                    '.TableBorder = tbAll: .TableCell(tcCols) = 8: .TableCell(tcRows) = 2
                    .TableCell(tcColWidth, , 1) = "0.90in": .TableCell(tcColWidth, , 2) = "0.85in"
                    .TableCell(tcColWidth, , 3) = "0.85in": .TableCell(tcColWidth, , 4) = "1.25in"
                    .TableCell(tcColWidth, , 5) = "0.85in": .TableCell(tcColWidth, , 6) = "0.75in"
                    .TableCell(tcColWidth, , 7) = "1.05in": .TableCell(tcColWidth, , 8) = "1.20in"
                    .TableCell(tcColAlign) = taRightTop
                    
                    If Val(CN(0).recs!SplDisc) <> 0 Then
                        .TableCell(tcText, 1, 1) = "SplDisc:"
                        .TableCell(tcText, 1, 2) = Format(CN(0).recs!SplDisc, "##.00\%")
                    End If
                    
                    If Val(CN(0).recs!NetCGSTAmt) <> 0 Then
                        .TableCell(tcText, 1, 3) = "CGST:"
                        .TableCell(tcText, 1, 4) = Format(CN(0).recs!NetCGSTAmt, "##0.00")
                    End If
                    
                    If Val(CN(0).recs!NetSGSTAmt) <> 0 Then
                        .TableCell(tcText, 1, 5) = "SGST:"
                        .TableCell(tcText, 1, 6) = Format(CN(0).recs!NetSGSTAmt, "##0.00")
                    End If
                    
                    .TableCell(tcText, 1, 7) = "Round Off:"
                    .TableCell(tcText, 1, 8) = Format(CN(0).recs!RoundOff, "(+) #0.00; (-) #0.00; NIL")
                    
                    
                    
                    
                    If Val(CN(0).recs!BulkDisc) <> 0 Then
                        .TableCell(tcText, 2, 1) = "BulkDisc:"
                        .TableCell(tcText, 2, 2) = Format(CN(0).recs!BulkDisc, "##.00\%")
                    End If
                    
                    If Val(CN(0).recs!NetIGSTAmt) <> 0 Then
                        .TableCell(tcText, 2, 3) = "IGST:"
                        .TableCell(tcText, 2, 4) = Format(CN(0).recs!NetIGSTAmt, "##0.00")
                    End If
                    
                    If Val(CN(0).recs!NetCessAmt) <> 0 Then
                        .TableCell(tcText, 2, 5) = "Cess:"
                        .TableCell(tcText, 2, 6) = Format(CN(0).recs!NetCessAmt, "##0.00")
                    End If

                    If Val(CN(0).recs!NetAmount) <> 0 Then
                        .TableCell(tcText, 2, 7) = "Net Amount:"
                        .TableCell(tcText, 2, 8) = "{\b\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Rupee Foradian;}{\f1\fswiss\fcharset0 Verdana;}}\viewkind4\uc1\pard\qr\f0\fs20 `\f1" & Format(CN(0).recs!NetAmount, "   ##,##0.00") & "}"
                    End If
                .EndTable
            End If
        End With
    End If
End Sub

Public Sub Render_Section_H(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    If PageNum = vp.PageCount Then
        With vp
            If printFormatOption <> 1 And printFormatOption <> 4 Then
                .FontBold = True
                .StartTable
                    .CurrentY = .CurrentY + 50
                    .TableBorder = tbBottom: .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "7.70in": .TableCell(tcColAlign, , 1) = taLeftMiddle
                    .TableCell(tcText, 1, 1) = "(" & ConvertCurrencyToEnglish(Val(CN(0).recs!NetAmount)) & ")"
                .EndTable
                .FontBold = False
            End If
        End With
    End If
End Sub

Public Sub Render_Section_I(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    If PageNum = vp.PageCount Then
        With vp
        .FontBold = True
        Select Case printFormatOption
            Case 0, 1, 2, 3, 4:
                .StartTable
                    .CurrentY = .CurrentY + 50
                    .TableBorder = tbBottom: .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "7.70in": .TableCell(tcColAlign, , 1) = taLeftMiddle
                    .TableCell(tcText, 1, 1) = "BillComments: " & CN(0).recs!Comments
                .EndTable
        End Select
        .FontBold = False
        End With
    End If
End Sub

Public Sub Render_Section_J(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    With vp
        If LedgerBalancePrinting = True Then
            If PageNum = vp.PageCount Then
                With vp
                    .FontBold = True
                        .StartTable
                            .CurrentY = .CurrentY + 50
                            .TableBorder = tbBottom: .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                            .TableCell(tcColWidth, , 1) = "7.70in": .TableCell(tcColAlign, , 1) = taLeftMiddle
                            .TableCell(tcText, 1, 1) = "A/c balance of " & CN(0).recs!id & " after this transactions on" & CN(0).recs!INVDate & " is : " & Format(WhatIsLedgerBalance(CN(0).recs!id, CN(0).recs!INVDate), "##,##0.00 Dr; ##,##0.00 Cr; NIL")
                            .TableCell(tcText, 1, 1) = RunningNews
                        .EndTable
                    .FontBold = False
                End With
            End If
        End If
    End With
End Sub

Public Sub Render_Section_K(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    On Error Resume Next
    PREPARER = Split(CN(0).recs!UserNo, "|")
    Select Case printFormatOption
        Case 2:
            With vp
                .CurrentY = .CurrentY + 50
                .StartTable
                    .TableBorder = tbNone: .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "4.90in": .TableCell(tcColWidth, , 2) = "1.10in": .TableCell(tcColWidth, , 3) = "1.70in"
                    .TableCell(tcColAlign, , 1) = taLeftMiddle: .TableCell(tcColAlign, , 2) = taCenterTop: .TableCell(tcColAlign, , 3) = taCenterTop
                    .TableCell(tcText, 1, 1) = ""
                    .TableCell(tcText, 1, 2) = "{\par\i prepared by \par\b " & Right(PREPARER(0), 10) & " \b0 \par checked by }"
                    .TableCell(tcText, 1, 3) = "{\b\par For " & CompanyName & " \par\par\par\b0 Auth. Signatory }"
                .EndTable
            End With
        Case Else:
            With vp
                .CurrentY = .CurrentY + 50
                .StartTable
                    .TableBorder = tbNone: .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "4.90in": .TableCell(tcColWidth, , 2) = "1.10in": .TableCell(tcColWidth, , 3) = "1.70in"
                    .TableCell(tcColAlign, , 1) = taJustMiddle: .TableCell(tcColAlign, , 2) = taCenterTop: .TableCell(tcColAlign, , 3) = taCenterTop
                    .TableCell(tcText, 1, 1) = Terms
                    .TableCell(tcText, 1, 2) = "{\par\i prepared by \par\b " & Right(PREPARER(0), 10) & " \b0 \par checked by }"
                    .TableCell(tcText, 1, 3) = "{\b\par For " & CompanyName & " \par\par\par\b0 Auth. Signatory }"
                .EndTable
            End With
    End Select
End Sub



'==============================================================================================
Private Sub SetFont(S As String)
    On Error Resume Next
    vp.FontName = ReadFont(S, 0)
    vp.FontSize = ReadFont(S, 1)
    vp.FontBold = ReadFont(S, 2)
    vp.FontItalic = ReadFont(S, 3)
End Sub

Private Sub SetFlxFont(S As String)
    flxItems.FontName = ReadFont(S, 0)
    flxItems.FontSize = ReadFont(S, 1)
    flxItems.FontBold = ReadFont(S, 2)
    flxItems.FontItalic = ReadFont(S, 3)
End Sub

Private Sub cmdPrint_Click()
    vp.Visible = False: vp.PrintDoc True: vp.Visible = True
    cmdQuit.SetFocus
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdTransportForwarding_Click()
    If Val(sArray(0)) <> 0 Then
    With vp
        .StartDoc
        .CurrentY = 200
        .TextAlign = taCenterTop
        .Text = "FORWARDING NOTE"
        .TextAlign = taLeftTop
        Y = .CurrentY
        .CurrentY = Y + 250
        .StartTable
            .TableBorder = tbBottom
            .TableCell(tcCols) = 3: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "3.00in"
            .TableCell(tcColWidth, , 2) = "1.50in"
            .TableCell(tcColWidth, , 3) = "3.00in"
            
            .TableCell(tcColAlign, , 1) = taLeftTop
            .TableCell(tcColAlign, , 2) = taCenterTop
            .TableCell(tcColAlign, , 3) = taLeftTop
            
            .TableCell(tcText, 1, 1) = "{\b\fs38 " & CompanyName & " \par\fs24\b0 " & CompanyAddr0 & " \par " & CompanyPhone & " \par " & CompanyFax & " }"
            .TableCell(tcText, 1, 2) = ""
            .DrawPicture mdiOne.ImgList.ListImages(1).Picture, .PageWidth / 2 - 415, .CurrentY + 300, 837, 1100
            .TableCell(tcText, 1, 3) = "{" & Salutation & " \b " & CN(0).recs!Name & "(" & CN(0).recs!id & ")" & " } \par " & CN(0).recs!Address & " \par " & CN(0).recs!City & ". Ph: " & CN(0).recs!Phones & " \par\par Through, " & CN(0).recs!TName & "\par " & CN(0).recs!Taddress & " \par " & CN(0).recs!TCity & " }"
        .EndTable
        .StartTable
            .TableBorder = tbBottom
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "7.50in"
            .TableCell(tcColAlign, , 1) = taJustTop
            .TableCell(tcText, 1, 1) = "Dear Sir," & vbCrLf & "      Please accept the following parcels which contains printed items in good condition vide your confirm order by .................... dated ..................." & vbCrLf
        .EndTable
        .StartTable
            .TableBorder = tbBottom
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "3.75in"
            .TableCell(tcColWidth, , 2) = "3.75in"
            .TableCell(tcColAlign, , 1) = taJustTop
            .TableCell(tcColAlign, , 2) = taJustTop
            .TableCell(tcText, 1, 1) = vbCrLf & "Bill#" & MemoNo & vbCrLf & "Date:" & Format(CN(0).recs!DBDate, "dd-mmm-yyyy") & vbCrLf & "Destination:" & CN(0).recs!City & vbCrLf
            .TableCell(tcText, 1, 2) = vbCrLf & "VALUE Rs." & Format(CN(0).recs!NetAmount, "##,#0.00") & vbCrLf & "Bundles:" & CN(0).recs!BundleCount & vbCrLf & "ToPayMode:" & ToPayMode
        .EndTable
        .StartTable
            .TableBorder = tbBottom
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "3.75in"
            .TableCell(tcColWidth, , 2) = "3.75in"
            .TableCell(tcColAlign, , 1) = taJustTop
            .TableCell(tcColAlign, , 2) = taCenterTop
            .TableCell(tcText, 1, 1) = "NOTE: If any discrepancy kindly inform alongwith this forwarding note."
            .TableCell(tcText, 1, 2) = "For " & CompanyName & vbCrLf & vbCrLf & "Despatcher"
        .EndTable
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "7.50in"
            .TableCell(tcColAlign, , 1) = taJustTop
            .TableCell(tcText, 1, 1) = "Enclosures: Copy of bill/ challan attached."
        .EndTable
        .EndDoc
    End With
    End If
End Sub

Private Sub txtDBRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
         Ref = Val(txtDBRef.Text)
       '  RenderMemo cmbPrintFormat.ListIndexa
    End If
End Sub

Private Sub chkPNotes_Click()
    RenderMemo cmbPrintFormat.ListIndex
End Sub

Private Sub DebugPageSizes()
    'Debug.Print "Paper sizes available on the "; vp.Device; ":"
    cmbPaperSizes.Clear
    
    For i = 1 To 256
        If vp.PaperSizes(i) Then
            cmbPaperSizes.AddItem i
            'vp.PaperSize = i
            'Debug.Print " Paper size "; Format(i, "000"); " is available & PprWd=" & Format(vp.PaperWidth / 1440, "#00.00") & " PprHt=" & Format(vp.PaperHeight / 1440, "#00.00") & " Pgwd=" & Format(vp.PageWidth / 1440, "#00.00") & " Pght=" & Format(vp.PageHeight / 1440, "#00.00")
        End If
    Next
End Sub

Private Sub txtPaperSize_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        vp.PaperSize = pprUser
        vp.PaperWidth = Val(txtPaperSize(0).Text)
        vp.PaperHeight = Val(txtPaperSize(1).Text)
    End If
End Sub

