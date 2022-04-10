VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSaleAndStockHolding 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SaleAndStockHolding...."
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10110
   Icon            =   "frmSaleAndStockHolding.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10050
      TabIndex        =   0
      Top             =   8715
      Width           =   10110
      Begin VB.ComboBox txtReportFor 
         Height          =   315
         Left            =   4470
         TabIndex        =   7
         Top             =   45
         Width           =   1335
      End
      Begin VB.CheckBox chkMonthly 
         Caption         =   "Monthly"
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
         Left            =   2085
         TabIndex        =   4
         Top             =   60
         Width           =   1350
      End
      Begin VB.PictureBox Picture2 
         Height          =   405
         Left            =   5835
         ScaleHeight     =   345
         ScaleWidth      =   4170
         TabIndex        =   1
         Top             =   0
         Width           =   4230
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
            Left            =   3015
            TabIndex        =   3
            Top             =   0
            Width           =   1155
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
            Left            =   1875
            TabIndex        =   2
            Top             =   0
            Width           =   1155
         End
         Begin VB.CommandButton cmdShowStockHolding 
            Caption         =   "ShowStockHolding"
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
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1890
         End
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   330
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
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
         Format          =   45481987
         CurrentDate     =   38023
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Report for: "
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
         Left            =   3495
         TabIndex        =   8
         Top             =   75
         Width           =   930
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   1  'Align Top
      Height          =   8700
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10110
      _cx             =   17833
      _cy             =   15346
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
      MultiTotals     =   0   'False
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
Attribute VB_Name = "frmSaleAndStockHolding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MyCAPTION = "Sale and stock holding report..."
Private MyData(5) As New clsData
Private sqlStr As String
Private RKsqlStr As String
Private LEVEL As Integer

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    Me.Move 0, 0
    txtReportFor.Clear: txtReportFor.AddItem "MERGED": txtReportFor.AddItem "ALL": txtReportFor.AddItem "PU": txtReportFor.AddItem "PR": txtReportFor.AddItem "SA": txtReportFor.AddItem "SR": txtReportFor.AddItem "TI": txtReportFor.AddItem "TO"
    txtReportFor.ListIndex = 4
    txtDate.Value = Now
    Me.Enabled = True
    LEVEL = 0
End Sub

Private Sub LoadReport()
    LEVEL = 0
    If chkMonthly.Value = 0 Then
        Select Case txtReportFor.Text
            Case "MERGED": sqlStr = "SELECT PublisherName, Sum(Qty) as Qty, Sum(Value) as Value from appview_PublisherWiseSalesAndStockHolding WHERE Datediff(D, TDate, " & QT(Format(txtDate.Value, "dd-mmm-yyyy")) & ")=0 Group by publishername order by 3 DESC"
            Case "ALL": sqlStr = "SELECT PublisherName, Type, Sum(Qty) as Qty, Sum(Value) as Value from appview_PublisherWiseSalesAndStockHolding WHERE Datediff(D, TDate, " & QT(Format(txtDate.Value, "dd-mmm-yyyy")) & ")=0 Group by type, publishername order by 4 DESC, type"
            Case Else: sqlStr = "SELECT PublisherName, Type, ABS(Sum(Qty)) as Qty, ABS(Sum(Value)) as Value from appview_PublisherWiseSalesAndStockHolding WHERE Datediff(D, TDate, " & QT(Format(txtDate.Value, "dd-mmm-yyyy")) & ")=0 AND TYPE=" & QT(txtReportFor.Text) & " Group by type, publishername order by 4 DESC, type"
        End Select
    Else
        Select Case txtReportFor.Text
            Case "MERGED": sqlStr = "SELECT PublisherName, Sum(Qty) as Qty, Sum(Value) as Value from appview_PublisherWiseSalesAndStockHolding WHERE Datediff(M, TDate, " & QT(Format(txtDate.Value, "dd-mmm-yyyy")) & ")=0 Group by publishername order by 3 DESC"
            Case "ALL": sqlStr = "SELECT PublisherName, Type, Sum(Qty) as Qty, Sum(Value) as Value from appview_PublisherWiseSalesAndStockHolding WHERE Datediff(M, TDate, " & QT(Format(txtDate.Value, "dd-mmm-yyyy")) & ")=0 Group by type, publishername order by 4 DESC, type"
            Case Else: sqlStr = "SELECT PublisherName, Type, ABS(Sum(Qty)) as Qty, ABS(Sum(Value)) as Value from appview_PublisherWiseSalesAndStockHolding WHERE Datediff(M, TDate, " & QT(Format(txtDate.Value, "dd-mmm-yyyy")) & ")=0 AND TYPE=" & QT(txtReportFor.Text) & " Group by type, publishername order by 4 DESC, type"
        End Select
    End If
    
    MyData(0).dbOpen sqlStr
    Set flxReport.DataSource = Nothing
    flxReport.Clear
    Set flxReport.DataSource = MyData(0)
    For i = 0 To flxReport.COLS - 1
        If flxReport.ColDataType(i) = flexDTDate Then flxReport.ColFormat(i) = "dd-mmm-yy"
    Next
    
    flxReport.COLS = flxReport.COLS + 3
    flxReport.TextMatrix(0, flxReport.COLS - 3) = "Stock"
    flxReport.TextMatrix(0, flxReport.COLS - 2) = "StockValue"
    flxReport.TextMatrix(0, flxReport.COLS - 1) = "Ratio"
    flxReport.ColFormat(flxReport.COLS - 4) = "#.00"
    flxReport.ColFormat(flxReport.COLS - 2) = "#.00"
    flxReport.ColFormat(flxReport.COLS - 1) = "#.00"
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub chkMonthly_Click()
    LoadReport
End Sub

Private Sub flxReport_DblClick()
    On Error Resume Next
    If LEVEL = 0 Then
        MyData(5).dbOpen "SELECT PUBLISHERNAME, MONTH(TDATE), DATENAME(month, TDATE), TYPE, SUM(QTY) AS QTY, SUM(VALUE) AS VALUE FROM appview_PublisherWiseSalesAndStockHolding WHERE PUBLISHERNAME=" & QT(flxReport.TextMatrix(flxReport.Row, 0)) & " GROUP BY PUBLISHERNAME, MONTH(TDATE), DATENAME(month, TDATE), TYPE ORDER BY 1,2,4"
        Set flxReport.DataSource = Nothing
        flxReport.Clear
        Set flxReport.DataSource = MyData(5)
        flxReport.AutoSize 0, flxReport.COLS - 1
        LEVEL = 1
    End If
End Sub

Private Sub flxReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim sum As Double
    sum = 0
    For i = 0 To flxReport.SelectedRows - 1
        If flxReport.SelectedRow(i) >= 1 Then
            sum = sum + Val(flxReport.TextMatrix(flxReport.SelectedRow(i), flxReport.Col))
        End If
        Me.Caption = MyCAPTION & "Sum on " & str(flxReport.Col) & " = " & Format(sum, "##,##0.00")
    Next
    If Button = 2 And Shift = vbCtrlMask Then
        SaveGrid Me.flxReport, "STOCK HOLDING REPORT OF " & CompanyDivision & " " & Format(Now, "DD-MMM-YY HHMM")
    End If
    If Button = 2 And Shift = vbCtrlMask + vbShiftMask + vbAltMask Then
        InputBox "SQL", , RKsqlStr
    End If
End Sub

Private Sub txtDate_Change()
    LoadReport
End Sub

Private Sub txtReportFor_Change()
    LoadReport
End Sub

Private Sub cmdShowStockHolding_Click()
    cmdShowStockHolding.Caption = "ShowStkHolding" & vbCrLf & getLastDateOfMonth(txtDate.Value)
    ShowStockHolding
End Sub

Private Sub cmdRefresh_Click()
    LoadReport
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub ShowStockHolding()
    On Error Resume Next
    Me.Enabled = False
    For i = 1 To flxReport.ROWS - 1
        PublisherName = flxReport.TextMatrix(i, 0)
        RKsqlStr = "SELECT SUM(QTY) as Stock, SUM(VALUE) as StockValue FROM appview_PublisherWiseSalesAndStockHolding WHERE PUBLISHERNAME=" & QT(PublisherName) & " AND DATEDIFF(D, TDATE, " & QT(Format(getLastDateOfMonth(txtDate.Value), "DD-MMM-YYYY")) & ")>=0 " & " GROUP BY PUBLISHERNAME"
        MyData(1).dbOpen RKsqlStr, 1
        If Not MyData(1).recs.EOF Then
            flxReport.TextMatrix(i, flxReport.COLS - 3) = MyData(1).recs!Stock
            flxReport.TextMatrix(i, flxReport.COLS - 2) = MyData(1).recs!StockValue
        Else
            flxReport.TextMatrix(i, flxReport.COLS - 3) = "No data"
            flxReport.TextMatrix(i, flxReport.COLS - 2) = "No data"
        End If
        flxReport.TextMatrix(i, flxReport.COLS - 1) = Val(flxReport.TextMatrix(i, flxReport.COLS - 4)) / Val(flxReport.TextMatrix(i, flxReport.COLS - 2)) * 100
        flxReport.ShowCell i, flxReport.COLS - 1
        DoEvents
    Next
    Me.Enabled = True
End Sub
