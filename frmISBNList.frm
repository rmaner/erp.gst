VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmISBNList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ISBN BARCODE PRINTING..."
   ClientHeight    =   9045
   ClientLeft      =   -45
   ClientTop       =   420
   ClientWidth     =   11910
   Icon            =   "frmISBNList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   360
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3555
      TabIndex        =   13
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtFrom 
         BorderStyle     =   0  'None
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
         Left            =   1095
         TabIndex        =   0
         Text            =   "0"
         Top             =   60
         Width           =   885
      End
      Begin VB.TextBox txtTo 
         BorderStyle     =   0  'None
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
         Left            =   2475
         TabIndex        =   1
         Text            =   "28"
         Top             =   45
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "ISBN From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   15
         Top             =   30
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "to "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2175
         TabIndex        =   14
         Top             =   30
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   3615
      ScaleHeight     =   315
      ScaleWidth      =   8235
      TabIndex        =   10
      Top             =   -15
      Width           =   8295
      Begin VB.TextBox txtFontSize 
         BorderStyle     =   0  'None
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
         Left            =   2955
         TabIndex        =   4
         Text            =   "28"
         Top             =   60
         Width           =   720
      End
      Begin VB.TextBox txtBlankCells 
         BorderStyle     =   0  'None
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
         Left            =   1590
         TabIndex        =   3
         Text            =   "0"
         Top             =   60
         Width           =   690
      End
      Begin VB.CommandButton cmdRender 
         Caption         =   "&Render"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4275
         TabIndex        =   5
         Top             =   0
         Width           =   1320
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
         Height          =   315
         Left            =   5595
         TabIndex        =   6
         Top             =   0
         Width           =   1320
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
         Height          =   315
         Left            =   6915
         TabIndex        =   7
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2460
         TabIndex        =   12
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Start after cell:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   11
         Top             =   30
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8760
      Left            =   3615
      TabIndex        =   8
      Top             =   300
      Width           =   8295
      Begin VSPrinter8LibCtl.VSPrinter vp 
         Height          =   8550
         Left            =   75
         TabIndex        =   9
         Top             =   135
         Width           =   8160
         _cx             =   14393
         _cy             =   15081
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
   Begin VSFlex8UCtl.VSFlexGrid flxItems 
      Height          =   8655
      Left            =   15
      TabIndex        =   2
      Top             =   375
      Width           =   3600
      _cx             =   6350
      _cy             =   15266
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorBkg    =   -2147483636
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
      SelectionMode   =   0
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
      AutoSearch      =   1
      AutoSearchDelay =   5
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
      Editable        =   2
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
End
Attribute VB_Name = "frmISBNList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New clsData

Dim SQ, ISBN As String
Dim BarcodeList() As String
Private Const ROWS = 13
Private Const COLS = 5

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    txtBlankCells.BackColor = Me.BackColor
    txtFontSize.BackColor = Me.BackColor
    txtFrom.BackColor = Me.BackColor
    txtTo.BackColor = Me.BackColor
    txtFontSize.Text = 32
    
    vp.PaperSize = pprA4
    vp.MarginLeft = 200: vp.MarginRight = 50
    vp.MarginTop = 750: vp.MarginBottom = 150
    vp.FontName = "Code EAN13"
    
    SQ = "SELECT ISBN, 0 As BARCODE, 1 AS Repeat, itemID, itemNAME FROM itemS ORDER BY 1"
    DataLoad
End Sub

Private Sub flxitems_DblClick()
    SQ = InputBox("Input SQL to load desired data in grid...")
    DataLoad
End Sub

Private Sub flxitems_EnterCell()
    If flxitems.Col = 2 Then
        flxitems.AutoSearch = flexSearchNone
        flxitems.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub flxitems_LeaveCell()
    flxitems.AutoSearch = flexSearchFromTop
    flxitems.Editable = flexEDNone
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then GenerateISBNList
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then GenerateISBNList
End Sub

Private Sub txtFontSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdRender_Click
End Sub

Private Sub cmdRender_Click()
    Dim Exhausted As Boolean
    Exhausted = False
    
    CreateBarcodeList
    With vp
        .PenColor = vbYellow
        .PenStyle = psDot
        .FontName = "Code EAN13"
        .FontSize = Val(txtFontSize.Text)
    
        .StartDoc
            itemCount = UBound(BarcodeList) + 1
            i = 0
            While Not Exhausted
                .StartTable
                .TableBorder = tbAll
                .TableCell(tcCols) = COLS: .TableCell(tcRows) = ROWS
                .TableCell(tcColWidth, 1, 1, ROWS, COLS) = "1.6in"
                .TableCell(tcRowHeight, 1, 1, ROWS, COLS) = "0.83in"
                .TableCell(tcRowKeepTogether) = True
                
                .TableCell(tcColAlign) = taCenterMiddle
                For j = 0 To (ROWS * COLS) - 1
                    R = Int(j / COLS) + 1
                    c = (j Mod COLS) + 1
                    If (i < itemCount) Then .TableCell(tcText, R, c) = BarcodeList(i)
                    i = i + 1
                    If i >= itemCount Then Exhausted = True
                Next
                .EndTable
                If Not Exhausted Then .NewPage
            Wend
        .EndDoc
    End With
    cmdPrint.SetFocus
End Sub

Private Sub cmdPrint_Click()
    vp.PrintDoc
    cmdQuit.SetFocus
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub DataLoad()
    On Error Resume Next
    CN.dbOpen SQ
    Set flxitems.DataSource = CN.recs
    For i = 1 To flxitems.ROWS - 1
        ISBN = flxitems.TextMatrix(i, 0)
        flxitems.TextMatrix(i, 1) = EAN13(ISBN)
    Next
    flxitems.SelectionMode = flexSelectionFree
    flxitems.Cell(flexcpFontName, 1, 1, flxitems.ROWS - 1, 1) = "Code EAN13"
    flxitems.Cell(flexcpFontSize, 1, 1, flxitems.ROWS - 1, 1) = "14"
    flxitems.SelectionMode = flexSelectionListBox
End Sub

Private Sub GenerateISBNList()
    Dim ISBNfrom, ISBNto, i As Long
    ISBNfrom = Val(txtFrom.Text)
    ISBNto = Val(txtTo.Text)
    
    If ISBNfrom > ISBNto Then
        txtFrom.Text = ISBNto: txtTo.Text = ISBNfrom
        ISBNfrom = Val(txtFrom.Text): ISBNto = Val(txtTo.Text)
    End If
    
    Set flxitems.DataSource = Nothing
    flxitems.Delete
    flxitems.ROWS = ISBNto - ISBNfrom + 2
    flxitems.COLS = 3
    For i = 1 To (ISBNto - ISBNfrom + 1)
        flxitems.TextMatrix(i, 0) = Format((ISBNfrom + i - 1), "00-000-0000")
    Next
    For i = 1 To flxitems.ROWS - 1
        ISBN = flxitems.TextMatrix(i, 0)
        flxitems.TextMatrix(i, 1) = EAN13(ISBN)
        flxitems.TextMatrix(i, 2) = 1
    Next
    flxitems.SelectionMode = flexSelectionFree
    flxitems.Cell(flexcpFontName, 1, 1, flxitems.ROWS - 1, 1) = "Code EAN13"
    flxitems.Cell(flexcpFontSize, 1, 1, flxitems.ROWS - 1, 1) = "14"
    flxitems.SelectionMode = flexSelectionListBox
End Sub

Private Sub CreateBarcodeList()
    Dim Count, i, j As Long
    j = 0
    For i = 0 To flxitems.SelectedRows - 1
        Count = Count + Val(flxitems.TextMatrix(flxitems.SelectedRow(i), 2))
    Next
    Count = Count + 0

    If Count > 0 Then
        ReDim BarcodeList(Count - 1 + Val(txtBlankCells.Text))
    Else
        ReDim BarcodeList(0)
    End If
    For j = 0 To Val(txtBlankCells.Text) - 1
        BarcodeList(j) = ""
    Next
    For i = 0 To flxitems.SelectedRows - 1
        For k = 1 To Val(flxitems.TextMatrix(flxitems.SelectedRow(i), 2))
            BarcodeList(j) = flxitems.TextMatrix(flxitems.SelectedRow(i), 1)
            j = j + 1
        Next
    Next
    j = j + 0
End Sub
