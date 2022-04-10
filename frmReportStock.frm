VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmReportStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Stock..."
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicB 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   14820
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8820
      Width           =   14880
      Begin VB.ComboBox cmbFilter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   7
         Text            =   "Filter Zero Or Negative Stock"
         Top             =   15
         Width           =   3060
      End
      Begin VB.CommandButton cmdDeleteRow 
         Caption         =   "&DeleteRow(s)"
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
         Left            =   9270
         TabIndex        =   6
         Top             =   0
         Width           =   1860
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "&Calculate"
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
         Left            =   11115
         TabIndex        =   5
         Top             =   0
         Width           =   1845
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
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
         Left            =   12960
         TabIndex        =   3
         Top             =   0
         Width           =   1860
      End
   End
   Begin VB.PictureBox PicA 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14820
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14880
      Begin VB.ComboBox cmbPublisher 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmReportStock.frx":0000
         Left            =   0
         List            =   "frmReportStock.frx":0002
         TabIndex        =   1
         Text            =   "cmbPublisher"
         Top             =   0
         Width           =   14820
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxItem 
      Align           =   1  'Align Top
      Height          =   8430
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   14880
      _cx             =   26247
      _cy             =   14870
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
      AllowUserResizing=   1
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
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   1
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
Attribute VB_Name = "frmReportStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ORDERING = " ORDER BY Price Desc"
Private CN(15) As New clsData
Private Const PUBS = " PubName "

Private PriceCol, DiscCol, StockCol, StockValCol, DiscStockValCol As Integer
Dim tflx As Control
Dim tsql As String

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    Me.Move 0, 0
    flxItem.DataMode = flexDMFree
    PriceCol = 5: DiscCol = 6: StockCol = 7: StockValCol = 8: DiscStockValCol = 9
    tsql = "Select itemID, ISBN, itemNAME, AUTHORS, CURRENCY, PRICE, PDISCOUNT as DISC, STOCK, PRICE*STOCK AS StkVal, 0 as DiscStkVal from appview_ReportStock "
    
    CN(5).dbOpen "SELECT DISTINCT " & PUBS & " FROM APPVIEW_SELECTitem ORDER BY 1", 1
    If Not CN(5).recs.EOF Then CN(5).recs.MoveFirst: cmbPublisher.Clear
    cmbPublisher.AddItem "ALL PUBLISHERS"
    Do Until CN(5).recs.EOF
        cmbPublisher.AddItem CN(5).recs.FIELDS(0)
        CN(5).recs.MoveNext
    Loop
    cmbPublisher.ListIndex = 0
    cmbFilter.AddItem "All": cmbFilter.AddItem "NonZero": cmbFilter.AddItem "OnlyPositive": cmbFilter.AddItem "OnlyZeroStock": cmbFilter.AddItem "OnlyNegative"
    cmbFilter.ListIndex = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW
        For i = 0 To flxItem.COLS - 1
            flxItem.ColHidden(i) = False
        Next
    End If
    If KeyCode = 191 And Shift = vbCtrlMask Then  '/ hide oneb
        flxItem.ColHidden(flxItem.Col) = True
        If flxItem.Col < flxItem.COLS - 2 Then flxItem.Col = flxItem.Col + 1
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmbpublisher_Click()
    'cmbFilter.AddItem "All": cmbFilter.AddItem "NonZero": cmbFilter.AddItem "OnlyPositive": cmbFilter.AddItem "OnlyZeroStock": cmbFilter.AddItem "OnlyNegative"
    Select Case cmbFilter.ListIndex
        Case 0:
            If Trim(cmbPublisher.Text) = "ALL PUBLISHERS" Then
                CN(0).dbOpen tsql & ORDERING
            Else
                CN(0).dbOpen tsql & " WHERE " & PUBS & "=" & QT(Trim(cmbPublisher.Text)) & " " & ORDERING
            End If
        
        Case 1:
            If Trim(cmbPublisher.Text) = "ALL PUBLISHERS" Then
                CN(0).dbOpen tsql & " WHERE Stock <> 0 " & ORDERING
            Else
                CN(0).dbOpen tsql & " WHERE " & PUBS & "=" & QT(Trim(cmbPublisher.Text)) & " AND Stock <> 0 " & ORDERING
            End If
        
        Case 2:
            If Trim(cmbPublisher.Text) = "ALL PUBLISHERS" Then
                CN(0).dbOpen tsql & " WHERE Stock > 0 " & ORDERING
            Else
                CN(0).dbOpen tsql & " WHERE " & PUBS & "=" & QT(Trim(cmbPublisher.Text)) & " AND Stock > 0 " & ORDERING
            End If
        
        Case 3:
            If Trim(cmbPublisher.Text) = "ALL PUBLISHERS" Then
                CN(0).dbOpen tsql & " WHERE Stock = 0 " & ORDERING
            Else
                CN(0).dbOpen tsql & " WHERE " & PUBS & "=" & QT(Trim(cmbPublisher.Text)) & " AND Stock = 0 " & ORDERING
            End If
        
        Case 4:
            If Trim(cmbPublisher.Text) = "ALL PUBLISHERS" Then
                CN(0).dbOpen tsql & " WHERE Stock < 0 " & ORDERING
            Else
                CN(0).dbOpen tsql & " WHERE " & PUBS & "=" & QT(Trim(cmbPublisher.Text)) & " AND Stock < 0 " & ORDERING
            End If
    End Select
    
    
    Set flxItem.DataSource = CN(0).recs
    
    flxItem.ColFormat(PriceCol) = "##0.00"
    flxItem.ColFormat(StockValCol) = "##0.00"
    flxItem.ColFormat(DiscStockValCol) = "##0.00"
    flxItem.AutoSize 0, flxItem.COLS - 1
End Sub

Private Sub flxItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then
        SaveGrid Me.flxItem, "Stock Report " & cmbPublisher.Text & " at " & CompanyDivision
    End If
End Sub

Private Sub chkStockFilter_Click()
    cmbpublisher_Click
End Sub

Private Sub cmbFilter_Click()
    cmbpublisher_Click
End Sub


Private Sub cmdCalculate_Click()
    On Error Resume Next
    If flxItem.ROWS >= 1 Then
        For i = 1 To flxItem.ROWS - 1
            disc = CircularDiscount(flxItem.TextMatrix(i, DiscCol))
            flxItem.TextMatrix(i, DiscStockValCol) = Val(flxItem.TextMatrix(i, StockValCol)) * (1 - disc / 100)
            flxItem.Row = i: flxItem.ShowCell i, DiscStkValCol
        Next
    End If
    flxItem.SubtotalPosition = flexSTBelow
    flxItem.SubTotal flexSTSum, -1, StockCol, "#,##0", , , True, "Total"
    flxItem.SubTotal flexSTSum, -1, StockValCol, "#,##0.00", , , True, "Total"
    flxItem.SubTotal flexSTSum, -1, DiscStockValCol, "#,##0.00", , , True, "Total"
End Sub

Private Sub cmdDeleteRow_Click()
    For deleterow = 0 To flxItem.SelectedRows - 1
        If flxItem.SelectedRow(i) >= 1 Then flxItem.RemoveItem flxItem.SelectedRow(i)
    Next
End Sub

Private Sub cmdQuit_Click()
    Me.ValidateControls: Unload Me
End Sub
