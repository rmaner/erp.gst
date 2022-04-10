VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmPrintMemoLongFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PrintMemoLongFormat"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   9075
      Left            =   5355
      ScaleHeight     =   9015
      ScaleMode       =   0  'User
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   0
      Width           =   1635
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
         Height          =   360
         Index           =   1
         Left            =   810
         TabIndex        =   7
         Text            =   "0"
         Top             =   2985
         Visible         =   0   'False
         Width           =   780
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
         Height          =   360
         Index           =   0
         Left            =   15
         TabIndex        =   6
         Text            =   "0"
         Top             =   2985
         Visible         =   0   'False
         Width           =   780
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
         ItemData        =   "frmPrintMemoLongFormat.frx":0000
         Left            =   0
         List            =   "frmPrintMemoLongFormat.frx":0002
         TabIndex        =   5
         Text            =   "cmbPaperSizes"
         Top             =   330
         Width           =   1590
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
         Left            =   15
         TabIndex        =   4
         Top             =   2325
         Width           =   1560
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
         Left            =   15
         TabIndex        =   3
         Top             =   1860
         Width           =   1560
      End
      Begin VB.TextBox txtDBRef 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   660
         TabIndex        =   2
         Top             =   1380
         Width           =   915
      End
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
         ItemData        =   "frmPrintMemoLongFormat.frx":0004
         Left            =   0
         List            =   "frmPrintMemoLongFormat.frx":0006
         TabIndex        =   1
         Text            =   "cmbMemoFormat"
         Top             =   0
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "DBRef:"
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
         Left            =   30
         TabIndex        =   11
         Top             =   1410
         Width           =   915
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
         Left            =   353
         TabIndex        =   8
         Top             =   675
         Width           =   885
      End
   End
   Begin VSPrinter8LibCtl.VSPrinter vp 
      Align           =   3  'Align Left
      Height          =   9075
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5340
      _cx             =   9419
      _cy             =   16007
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
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
      PalettePicture  =   "frmPrintMemoLongFormat.frx":0008
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
         Height          =   960
         Left            =   900
         TabIndex        =   10
         Top             =   5850
         Visible         =   0   'False
         Width           =   3480
         _cx             =   6138
         _cy             =   1693
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
Attribute VB_Name = "frmPrintMemoLongFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'  PRINT BILL * LONG PAGES LIKE THERMAL PRINTERS *
'**********************************************************************


Private CN(5) As New clsData

Dim Ref, CurrY As Long
Dim TMAIN, Support, MemoHeading, Salutation, MemoNo, ToPayMode As String
Dim Terms, pdfFileName As String
Dim ItemCodeCol, itemNameCol, ProducerIDCol, ProducerNameCol As Integer
Dim LedgerBalancePrinting As Boolean

Private Sub Form_Load()
    On Error Resume Next
    Me.Move Screen.Width - Me.Width, 0
    ItemCodeCol = 2: itemNameCol = 3: ProducerIDCol = 5: ProducerNameCol = 6
    
    DebugPageSizes
    RunningNews = mdiOne.sckGo.GReadINI("[RunningNews]", "[END]")
    LedgerBalancePrinting = mdiOne.sckGo.GReadINI("[LedgerBalancePrinting]")
    cmbPaperSizes.ListIndex = Val(mdiOne.sckGo.GReadINI("[PaperSizeNumber]"))
    vp.PaperSize = Val(cmbPaperSizes.Text)
    vp.PaperSize = pprUser
    vp.PhysicalPage = False
    'vp.PageBorder = pbBottom
    vp.Orientation = orPortrait
    
    PprDim = Split(mdiOne.sckGo.GReadINI("[PaperHeight/Width/Left/Right/Top/Bottom]"), ":")
    'vp.PaperHeight = PprDim(0): vp.PaperWidth = PprDim(1)
    vp.MarginLeft = PprDim(2): vp.MarginRight = PprDim(3)
    vp.MarginTop = PprDim(4): vp.MarginBottom = PprDim(5)
    
    vp.PenStyle = psSolid: vp.PenColor = vbBlack: vp.BrushColor = vbBlack
    vp.PenWidth = 5
    'vp.TrueType = ttSubDevice
    vp.AbortWindow = True
    vp.AutoRTF = True
    vp.Copies = 1: vp.Collate = colTrue
    
    cmbPrintFormat.Clear
    cmbPrintFormat.AddItem "Format A: GENERAL"       ' 0
    cmbPrintFormat.AddItem "Format B: CHALLAN"       ' 1
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
        cmbPrintFormat.ListIndex = MemoPrintFormat                  'THIS RENDERS THE BILL <<<
    End If
    Me.Show vbModal
End Sub

Private Sub cmbPrintFormat_Click()
    RenderMemo cmbPrintFormat.ListIndex
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

Public Sub RenderMemo(ByVal printFormatOption As Integer)
    CN(0).dbOpen "Select * from " & TMAIN & " where DBRef=" & Ref, 0
    If CN(0).recs.RecordCount = 1 Then
        Select Case CN(0).recs!ToPayMode
            Case 0: ToPayMode = "PAID-FULL"
            Case 1: ToPayMode = "PAID-HALF"
            Case 2: ToPayMode = "PAID-ZERO"
            Case 3: ToPayMode = "TOPAY-FULL"
            Case 4: ToPayMode = "TOPAY-HALF"
            Case 5: ToPayMode = "TOPAY-ZERO"
        End Select
        SupportSQL = "SELECT Serial as Sl, ItemID, ItemCode as CODE, ItemName, ProducerName as Mfg, Packing, Qty, MRP, SRP, GST, Gross, aDiscAmt as 'Disc', Amount FROM " & Support & " WHERE DBRefX=" & Ref & "  ORDER BY SERIAL"
        CN(1).dbOpen SupportSQL, 1: Set flxItems.DataSource = CN(1).recs
        CN(2).dbOpen "SELECT DISTINCT * FROM CURRENCY WHERE CURRENCY<>" & QT("X"), 1
        CN(3).dbOpen "SELECT * FROM PERSONAL WHERE ID = " & QT(CN(0).recs!id), 1
        
        Select Case Support
            Case "SALE":
                Select Case UCase(CN(0).recs!Status)
                    Case "ORDER": MemoHeading = "ORDER FORM/REQUISITION FORM": MemoNo = CompanyBillInitial & "/PF/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                    Case "CHALLAN": MemoHeading = " APPROVAL CHALLAN": MemoNo = CompanyBillInitial & "/CH/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                    Case "CASH": MemoHeading = " CASH INVOICE": MemoNo = CompanyBillInitial & "/CS/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                    Case "CREDIT": MemoHeading = " CREDIT INVOICE": MemoNo = CompanyBillInitial & "/CR/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                End Select
                Salutation = "To, "
                Terms = SaleInvoiceTerms
            Case "SALERETURN":
                MemoHeading = "SALES RETURN/ CREDIT NOTE": MemoNo = "Memo# " & CompanyBillInitial & "/SR/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                Salutation = "To, "
                Terms = "K.A.:" & vbCrLf & "Validity of sales return is subject to acceptance of items by the concerned Producers. If the Producer does not accept the return, the same will be debited to your account."
            Case "PURCHASE":
                MemoHeading = "PURCHASE ORDER": MemoNo = "Memo# " & CompanyBillInitial & "/PO/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                Salutation = "To, "
                Terms = ""
            Case "PURCHASERETURN":
                MemoHeading = "PURCHASE RETURN": MemoNo = "Memo# " & CompanyBillInitial & "/PR/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                Salutation = "To, "
                Terms = ""
                RunningNews = ""
            Case "TOUT":
                MemoHeading = "        STK TRANSFER": MemoNo = "Memo# " & CompanyBillInitial & "/TO/" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                Salutation = "To, "
                Terms = ""
            Case "TIN":
                Select Case UCase(CN(0).recs!Status)
                    Case "ORDER":   MemoHeading = "ORDER": MemoNo = "Memo# " & CompanyBillInitial & "TI" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                                    Salutation = "To, "
                    Case "CHALLAN": MemoHeading = "        STK TRANSFER IN-CH": MemoNo = "Memo# " & CompanyBillInitial & "CH" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "00000")
                                    Salutation = "From, "
                    Case "CASH":    MemoHeading = "        STK TRANSFER IN-CA": MemoNo = "Inv# " & CompanyBillInitial & "CS" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                                    Salutation = "From, "
                    Case "CREDIT":  MemoHeading = "        STK TRANSFER IN-AC": MemoNo = "Inv# " & CompanyBillInitial & "CR" & FinancialYear(CN(0).recs!DBDate) & "/" & Format(Val(CN(0).recs!DBRef), "0000")
                                    Salutation = "From, "
                End Select
                Terms = ""
        End Select
        
        '**********************************************************************
        '**********************************************************************
        '                   START OF PAGE RENDERING
        '**********************************************************************
        '**********************************************************************
        Select Case printFormatOption
            Case 0, 1:
                vp.MarginTop = 200: vp.MarginBottom = 200
                vp.MarginLeft = 20: vp.MarginRight = 10
        End Select
        vp.StartDoc
            SetFont "[Render_Section_Thermal_A]": Render_Section_A printFormatOption, i     ' Print memo headings like Sale, Purchase etc.
            SetFont "[Render_Section_Thermal_B]": Render_Section_B printFormatOption, i     ' Print Company name, branch etc.
            SetFont "[Render_Section_Thermal_C]": Render_Section_C printFormatOption, i     ' Print company address, phone etc.
            SetFont "[Render_Section_Thermal_D]": Render_Section_D printFormatOption, i     ' Print Customer name, address, shipping details etc.
            'SetFont "[Render_Section_Thermal_E]": Render_Section_E printFormatOption , i   ' Print currency conversion rates
            SetFlxFont "[Render_Section_Thermal_F]": Render_Section_F printFormatOption, 0
            SetFont "[Render_Section_Thermal_G]": Render_Section_G printFormatOption, i     ' Print bulk discount, gst, netamount etc.
            'SetFont "[Render_Section_Thermal_H]": Render_Section_H printFormatOption, i     ' Print netamount in Words
            'SetFont "[Render_Section_Thermal_I]": Render_Section_I printFormatOption, i     ' Print comments if any on bill
            'SetFont "[Render_Section_Thermal_J]": Render_Section_J printFormatOption, i     ' Print customer account balance
            'SetFont "[Render_Section_Thermal_K]": Render_Section_K printFormatOption, i     ' Print Bill preparer, place for signature
        vp.EndDoc
        
        '**********************************************************************
        '**********************************************************************
        '                   END OF PAGE RENDERING
        '**********************************************************************
        '**********************************************************************
    End If
End Sub

Public Sub Render_Section_A(ByVal printFormatOption As Integer, ByVal PageNum As Integer)        ' Print memo headings like Sale, Purchase etc.
    vp.CurrentY = 50
    vp.TextAlign = taCenterTop
    vp.Text = MemoHeading
    vp.TextAlign = taLeftTop
End Sub

Public Sub Render_Section_B(ByVal printFormatOption As Integer, ByVal PageNum As Integer)        ' Print Company name, branch etc.
    vp.CurrentY = vp.CurrentY + 150
    vp.TextAlign = taCenterTop
    vp.Text = CompanyName
    vp.TextAlign = taLeftTop
End Sub

Public Sub Render_Section_C(ByVal printFormatOption As Integer, ByVal PageNum As Integer)        ' Print company address, phone etc.
    Select Case printFormatOption
        Case 2:     ' Do nothing
        Case Else:
    End Select
End Sub

Public Sub Render_Section_D(ByVal printFormatOption As Integer, ByVal PageNum As Integer)            ' Print Customer name, address, shipping details etc.
    With vp
        .CurrentY = .CurrentY + 300
        .StartTable
            .TableBorder = tbAll: .TableCell(tcRowSpaceAfter) = 10
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, 1, 1) = "2.8in"
            .TableCell(tcText, 1, 1) = "{" & MemoNo & "  Dt: " & Format(CN(0).recs!DBDate, "dd-MM-yy") & "\par " & Salutation & "\b " & CN(0).recs!Name & "\b0(" & CN(0).recs!id & ")" & "\par " & CN(0).recs!Address & "\par " & CN(0).recs!City & "\par " & "PAN: " & CN(3).recs!PAN & "\par " & "GSTIN: " & CN(3).recs!GSTIN & "}"
        .EndTable
    End With
End Sub

Public Sub Render_Section_E(ByVal printFormatOption As Integer, ByVal PageNum As Integer)            ' Print currency conversion rates
    With vp
        .CurrentY = .CurrentY + 50
        .StartTable
            cur = "Conversion rates: "
            CN(2).recs.MoveFirst
            Do Until CN(2).recs.EOF
                cur = cur & "  " & CN(2).recs!Currency & "=" & Format(CN(2).recs!CurrPrice, "#0.00")
                CN(2).recs.MoveNext
            Loop
            .TableBorder = tbNone
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "3.0in"
            .TableCell(tcColAlign, , 1) = taLeftTop
            .TableCell(tcText, 1, 1) = cur
        .EndTable
    End With
End Sub

Public Sub Render_Section_F(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    flxItems.BackColorFixed = vbWhite
    flxItems.GridLines = flexGridFlat: flxItems.GridColor = vbWhite
    flxItems.GridLinesFixed = flexGridFlatHorz: flxItems.GridColorFixed = vbYellow
    flxItems.BorderStyle = flexBorderFlat
    flxItems.WordWrap = True
    
    flxItems.ColAlignment(0) = flexAlignCenterTop       'SERIAL
    flxItems.ColAlignment(1) = flexAlignRightTop        'itemID
    flxItems.ColAlignment(2) = flexAlignLeftTop         'ItemCode
    flxItems.ColAlignment(3) = flexAlignLeftTop         'itemNAME
    flxItems.ColAlignment(4) = flexAlignLeftTop         'Mfg
    flxItems.ColAlignment(5) = flexAlignLeftTop         'Packing
    flxItems.ColAlignment(6) = flexAlignRightTop         'Qty
    flxItems.ColAlignment(7) = flexAlignRightTop        'MRP
    flxItems.ColAlignment(8) = flexAlignRightTop        'SRP
    flxItems.ColAlignment(9) = flexAlignRightTop        'GST
    flxItems.ColAlignment(10) = flexAlignRightTop        'Gross
    flxItems.ColAlignment(11) = flexAlignRightTop       'Disc
    flxItems.ColAlignment(12) = flexAlignRightTop       'Amount
    
    flxItems.AddItem ""
    flxItems.TextMatrix(flxItems.ROWS - 1, 6) = CN(0).recs!itemCount
    flxItems.TextMatrix(flxItems.ROWS - 1, 12) = Format(CN(0).recs!itemAmount, "##0.00")
    flxItems.Select flxItems.ROWS - 1, 0, flxItems.ROWS - 1, flxItems.COLS - 1
    flxItems.CellBorder RGB(0, 0, 125), 0, 1, 0, 1, 0, 0
    
    'Sl, ItemID, ItemCode as CODE, ItemName, ProducerName as Mfg, Packing, Qty, MRP, SRP, Gross, aDiscAmt as 'Disc', Amount
    flxItems.ColWidth(0) = "400"    'SERIAL
    flxItems.ColWidth(1) = "0"      'itemID
    flxItems.ColWidth(2) = "0"   'ItemCode
    flxItems.ColWidth(3) = "1550"   'itemNAME
    flxItems.ColWidth(4) = "0"      'Mfg
    flxItems.ColWidth(5) = "0"      'Packing
    flxItems.ColWidth(6) = "400"    'Qty
    flxItems.ColWidth(7) = "600": flxItems.ColFormat(7) = "##,##0.00"      'MRP
    flxItems.ColWidth(8) = "0": flxItems.ColFormat(8) = "##,##0.00"      'SRP
    flxItems.ColWidth(9) = "0": flxItems.ColFormat(9) = "##,##0.00"      'GST
    flxItems.ColWidth(10) = "0": flxItems.ColFormat(9) = "##,##0"      'Gross
    flxItems.ColWidth(11) = "0": flxItems.ColFormat(10) = "##,##0.00"      'Disc
    flxItems.ColWidth(12) = "1050": flxItems.ColFormat(11) = "##,##0.00"      'Amount
    flxItems.AutoSize 3, 3, False
    
    vp.CurrentY = vp.CurrentY + 150
    vp.RenderControl = flxItems.hWnd
    vp.CurrentY = vp.CurrentY + 150
    Set flxItems.DataSource = Nothing
End Sub

Public Sub Render_Section_G(ByVal printFormatOption As Integer, ByVal PageNum As Integer)           ' Print bulk discount, gst, netamount etc.
    With vp
        .StartTable
            .TableBorder = tbTopBottom: .TableCell(tcCols) = 2: .TableCell(tcRows) = 4
            .TableCell(tcColWidth, , 1) = "1.80in": .TableCell(tcColWidth, , 2) = "1.00in"
            .TableCell(tcColAlign) = taRightTop
            .TableCell(tcText, 1, 1) = "GST:": .TableCell(tcText, 1, 2) = Format(CN(0).recs!NetGSTAmt, "##,##0.00")
            .TableCell(tcText, 2, 1) = "Cess:": .TableCell(tcText, 2, 2) = Format(CN(0).recs!NetCessAmt, "##,##0.00")
            .TableCell(tcText, 3, 1) = "Round Off:": .TableCell(tcText, 3, 2) = Format(CN(0).recs!RoundOff, "(+) #0.00; (-) #0.00; NIL")
            .TableCell(tcText, 4, 1) = "Net Amount:": .TableCell(tcText, 4, 2) = Format(CN(0).recs!NetAmount, "##,##0.00")
        .EndTable
    End With
End Sub

Public Sub Render_Section_H(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    With vp
        .StartTable
            .CurrentY = .CurrentY + 50
            .TableBorder = tbBottom: .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "7.70in": .TableCell(tcColAlign, , 1) = taLeftMiddle
            .TableCell(tcText, 1, 1) = "(" & ConvertCurrencyToEnglish(Val(CN(0).recs!NetAmount)) & ")"
        .EndTable
    End With
End Sub

Public Sub Render_Section_I(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    With vp
    .FontBold = True
    .StartTable
        .CurrentY = .CurrentY + 50
        .TableBorder = tbBottom: .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
        .TableCell(tcColWidth, , 1) = "3.0in": .TableCell(tcColAlign, , 1) = taLeftMiddle
        .TableCell(tcText, 1, 1) = "BillComments: " & CN(0).recs!Comments
    .EndTable
    .FontBold = False
    End With
End Sub

Public Sub Render_Section_J(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    If LedgerBalancePrinting = True Then
        With vp
            .FontBold = True
                .StartTable
                    .CurrentY = .CurrentY + 50
                    .TableBorder = tbBottom: .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
                    .TableCell(tcColWidth, , 1) = "3.0in": .TableCell(tcColAlign, , 1) = taLeftMiddle
                    .TableCell(tcText, 1, 1) = "A/c balance of " & CN(0).recs!id & " after this transactions on" & CN(0).recs!INVDate & " is : " & Format(WhatIsLedgerBalance(CN(0).recs!id, CN(0).recs!INVDate), "##,##0.00 Dr; ##,##0.00 Cr; NIL")
                    .TableCell(tcText, 1, 1) = RunningNews
                .EndTable
            .FontBold = False
        End With
    End If
End Sub

Public Sub Render_Section_K(ByVal printFormatOption As Integer, ByVal PageNum As Integer)
    On Error Resume Next
    PREPARER = Split(CN(0).recs!UserNo, "|")
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

Private Sub txtDBRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
         Ref = Val(txtDBRef.Text)
         'RenderMemo cmbPrintFormat.ListIndex
    End If
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

Private Sub cmdQuit_Click()
    Unload Me
End Sub
