VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReportsDayBook 
   Caption         =   "Day Item..."
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11318.11
   ScaleMode       =   0  'User
   ScaleWidth      =   10862.95
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   10560
      TabIndex        =   0
      Top             =   9360
      Width           =   10620
      Begin VB.PictureBox Picture2 
         Height          =   405
         Left            =   7050
         ScaleHeight     =   345
         ScaleWidth      =   3405
         TabIndex        =   1
         Top             =   0
         Width           =   3465
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&CreateReport"
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
            Left            =   30
            TabIndex        =   3
            Top             =   0
            Width           =   1695
         End
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
            Left            =   1710
            TabIndex        =   2
            Top             =   0
            Width           =   1695
         End
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   360
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   635
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
         CustomFormat    =   "ddd, dd-MMM-yyyy"
         Format          =   81068035
         CurrentDate     =   38023
      End
   End
   Begin VSPrinter8LibCtl.VSPrinter vp 
      Height          =   10466
      Left            =   15
      TabIndex        =   5
      Top             =   300
      Width           =   10620
      _cx             =   18732
      _cy             =   18461
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
      MarginLeft      =   360
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1140
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
      Zoom            =   60.9848484848485
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
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10620
      _cx             =   18732
      _cy             =   529
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
      AutoSizeMode    =   1
      AutoSearch      =   1
      AutoSearchDelay =   3
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
      WordWrap        =   -1  'True
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
Attribute VB_Name = "frmReportsDayBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MyCAPTION = "Date Reports..."
Private MyData(5) As New clsData
Private sqlStr As String

Private Sub Form_Load()
    txtDate.Value = Now
    Me.Move 0, 0
    mdiOne.SetFormFont Me
End Sub

Private Sub GenrateReport(sqlStr As String)
    'sqlStr = "EXEC appproc_DayitemSale DATE"
    If Right(sqlStr, 4) = "DATE" Then sqlStr = Replace(sqlStr, "DATE", QT(Format(txtDate.Value, "dd-MMM-yy")))
    MyData(0).dbOpen sqlStr
    
    Set flxReport.DataSource = Nothing
    flxReport.Clear
    flxReport.FontSize = "8"
    Set flxReport.DataSource = MyData(0)
    flxReport.ColFormat(2) = "DD-MMM-YY"
    
    flxReport.ColWidth(0) = "600"
    flxReport.ColWidth(1) = "800"
    flxReport.ColWidth(2) = "1200"
    flxReport.ColWidth(3) = "3000"
    flxReport.ColWidth(4) = "1400"
    flxReport.ColWidth(5) = "1000"
    flxReport.ColWidth(6) = "1000"
    flxReport.ColWidth(7) = "1800"

    For i = 1 To flxReport.ROWS - 1
        flxReport.TextMatrix(i, 0) = i
    Next
    flxReport.AutoSize 1, 7
    flxReport.SubtotalPosition = flexSTBelow
    flxReport.SubTotal flexSTSum, -1, 5, "#,##0.00", , , True, "Total"
    flxReport.SubTotal flexSTSum, -1, 6, "#,##0.00", , , True, "Total"
End Sub

Private Sub txtDate_Change()
    GenrateReport "EXEC appproc_DayitemSale DATE"
    GenrateReport "EXEC appproc_DayitemSaleReturn DATE"
    GenrateReport "EXEC appproc_DayitemPurchase DATE"
    GenrateReport "EXEC appproc_DayitemPurchaseReturn DATE"
    GenrateReport "EXEC appproc_DayitemRCT DATE"
    GenrateReport "EXEC appproc_DayitemPMT DATE"
    GenrateReport "EXEC appproc_DayitemVouchers DATE"
    GenrateReport "EXEC appproc_DayitemTIN DATE"
    GenrateReport "EXEC appproc_DayitemTOUT DATE"
    End Sub

Private Sub cmdRefresh_Click()
    With vp
        headString = "{\f0\fs16 " & reportTitle & "\par \b\fs22 " & CompanyName & "\par}"
        .PaperSize = pprA4
        .MarginLeft = 260: .MarginRight = 260
        '.FontName = "Tahoma"
        .StartDoc
            .TextAlign = taCenterTop
            .PenWidth = 40
            .Text = CompanyName & " " & vbCrLf & CompanyDivision & vbCrLf & "*** DAY END REPORT ***"
            .Text = "{ \b\ul\par\par SALE \ulnone }"
            GenrateReport "EXEC appproc_DayitemSale DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par SALE RETURN \ulnone }"
            GenrateReport "EXEC appproc_DayitemSaleReturn DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par PURCHASE \ulnone }"
            GenrateReport "EXEC appproc_DayitemPurchase DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par PURCHASE RETURN \ulnone }"
            GenrateReport "EXEC appproc_DayitemPurchaseReturn DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par RECEIPTS \ulnone }"
            GenrateReport "EXEC appproc_DayitemRCT DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par PAYMENTS \ulnone }"
            GenrateReport "EXEC appproc_DayitemPMT DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par VOUCHERS \ulnone }"
            GenrateReport "EXEC appproc_DayitemVouchers DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par STOCK IN \ulnone }"
            GenrateReport "EXEC appproc_DayitemTIN DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par STOCK OUT \ulnone }"
            GenrateReport "EXEC appproc_DayitemTOUT DATE": .RenderControl = flxReport.hWnd
            .Text = "{ \b\ul\par\par END OF REPORT \ulnone\par }"
        .EndDoc
    End With
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

