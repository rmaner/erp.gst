VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.4#0"; "UniBox10.ocx"
Begin VB.Form frmSelectItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ITEM SELECTOR"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   9675
   DrawMode        =   1  'Blackness
   DrawWidth       =   5
   Icon            =   "frmSelectItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicA 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9615
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   9675
      Begin VB.CheckBox chkUpdateMode 
         Alignment       =   1  'Right Justify
         Caption         =   "UpdateGrid"
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
         Left            =   5550
         TabIndex        =   18
         ToolTipText     =   "Update the Grid"
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtFound 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4125
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   -30
         Width           =   660
      End
      Begin VB.ComboBox cmbProducer 
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
         ItemData        =   "frmSelectItem.frx":4E0E
         Left            =   6510
         List            =   "frmSelectItem.frx":4E10
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "cmbProducer"
         Top             =   0
         Width           =   3105
      End
      Begin UniToolbox.UniText txtSearch 
         Height          =   315
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   3225
         _Version        =   65540
         _ExtentX        =   5689
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtQty 
         Height          =   315
         Left            =   3240
         TabIndex        =   1
         Top             =   0
         Width           =   750
         _Version        =   65540
         _ExtentX        =   1323
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label lblSearchCol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   4935
         TabIndex        =   3
         Top             =   -45
         Width           =   555
      End
   End
   Begin VB.PictureBox PicB 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   9615
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8025
      Width           =   9675
      Begin VB.CommandButton cmdFlyReport 
         Caption         =   "&FlyReport"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3090
         TabIndex        =   19
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuickItem 
         Caption         =   "Q&uickItem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4185
         TabIndex        =   15
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton cmdDuplicate 
         Caption         =   "&Duplicate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   14
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6375
         TabIndex        =   13
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "&Transfer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7470
         TabIndex        =   12
         Top             =   45
         Width           =   1095
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
         Height          =   300
         Left            =   8565
         TabIndex        =   11
         Top             =   45
         Width           =   1065
      End
      Begin VB.PictureBox Picture1 
         Height          =   435
         Left            =   30
         ScaleHeight     =   375
         ScaleWidth      =   1245
         TabIndex        =   8
         Top             =   0
         Width           =   1305
         Begin VB.TextBox txtAutoSelect 
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
            Left            =   525
            TabIndex        =   10
            Text            =   "SR-1"
            ToolTipText     =   "SA SR PU PR TO TI (Copy items from a bill to current bill)"
            Top             =   15
            Width           =   720
         End
         Begin VB.TextBox txtAutoSelectProducerID 
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
            Left            =   -15
            TabIndex        =   9
            Text            =   "ALL"
            ToolTipText     =   "Copy items from a bill - filtered on ProducerID"
            Top             =   15
            Width           =   540
         End
      End
      Begin MSForms.ToggleButton cmdFilter 
         Height          =   405
         Left            =   1410
         TabIndex        =   17
         Top             =   0
         Width           =   705
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1244;714"
         Value           =   "1"
         Picture         =   "frmSelectItem.frx":4E12
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ToggleButton cmdShowReportsDiscount 
         Height          =   405
         Left            =   2070
         TabIndex        =   16
         Top             =   0
         Width           =   705
         VariousPropertyBits=   738199579
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1244;714"
         Value           =   "0"
         Caption         =   "Trace"
         PicturePosition =   393224
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxItem 
      Align           =   1  'Align Top
      Height          =   7635
      Left            =   0
      TabIndex        =   5
      Top             =   375
      Width           =   9675
      _cx             =   17066
      _cy             =   13467
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
      AutoSearch      =   2
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
Attribute VB_Name = "frmSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ORDERING = " ORDER BY ProducerName ASC, ItemName ASC, MRP Desc"
Private CN(15) As New clsData
Private Const PRODUCERS = " ProducerName "
Private Const MyCAPTION = "Select Item here ... "

Dim rw As Long
Dim tflx As Control
Dim AddMode As Boolean
Dim SelectFor As String
Dim tsql As String
Dim ItemID_Col As Long, ItemCode_Col As Long, HSNCode_Col As Long, ItemName_Col As Long, ProducerName_Col As Long, Qty_Col As Long, MRP_Col As Long, SRP_Col, Stock_Col As Long
Public DestinationForm As frmPPSS

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    ItemID_Col = 2: ItemCode_Col = 3: HSNCode_Col = 4: ItemName_Col = 5: ProducerName_Col = 8: Qty_Col = 14: MRP_Col = 16: SRP_Col = 17: Stock_Col = 30
    Me.Move Screen.Width - Me.Width, 0
End Sub

Public Sub GetLink(ByRef DestForm As frmPPSS)
    Set DestinationForm = DestForm
    tsql = "Select * from appview_SelectItemSale "
    
    CN(5).dbOpen "SELECT DISTINCT " & PRODUCERS & " FROM appview_SelectItemSale ORDER BY 1", 1
    If Not CN(5).recs.EOF Then CN(5).recs.MoveFirst: cmbProducer.Clear
    cmbProducer.AddItem "*"
    Do Until CN(5).recs.EOF
        cmbProducer.AddItem CN(5).recs.FIELDS(0)
        CN(5).recs.MoveNext
    Loop
    
    CN(0).dbOpen tsql & " WHERE ItemID=100 " & ORDERING
    Set flxItem.DataSource = CN(0).recs
    
    rw = -1
    flxItem.Row = 0
    flxItem.Col = ItemCode: flxItem.AutoSearch = flexSearchFromCursor
    lblSearchCol.Caption = ItemCode
    cmbProducer.ListIndex = 0
    ColDisplay False
    Me.Caption = "SELECTOR > " & DestinationForm.MyCAPTION
    Me.Show
    Me.ZOrder 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 187 And Shift = vbCtrlMask Then '= SORT THE COLUMN
        flxItem.Col = flxItem.Col
        flxItem.Sort = flexSortGenericAscending
    End If
    If KeyCode = 188 And Shift = vbCtrlMask + vbShiftMask Then '< HIDE FOR DISCOUNT FILL
        ColDisplayDiscountFill False
    End If
    If KeyCode = 188 And Shift = vbCtrlMask Then  '< HIDE
        ColDisplay False
    End If
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW
        ColDisplay True
    End If
    If KeyCode = 191 And Shift = vbCtrlMask Then  '/ hide one
        flxItem.ColHidden(flxItem.Col) = True
        If flxItem.Col < flxItem.COLS - 2 Then flxItem.Col = flxItem.Col + 1
    End If
    If KeyCode = vbKeyF1 Then cmbProducer.SetFocus
    If KeyCode = vbKeyF12 Then DestinationForm.cmdSaveOrder_Click
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyT And Shift = vbCtrlMask + vbShiftMask Then  'COL=3 AND GOTO TXTSEARCH
        frmReportsFly.LoadReport "appproc_TraceItems " & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col))
    End If
    If KeyCode = vbKeyT And Shift = vbCtrlMask Then
        frmReportsFly.LoadReport "appproc_TraceItemsForPDiscount " & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col)), 5925
    End If
    If KeyCode = vbKeyM And Shift = vbCtrlMask + vbShiftMask Then
        B = Val(InputBox("ENTER DESTINATION ItemID TO MERGE SELECTED Items INTO OR 0 TO CANCEL THIS OPERATION"))
        If B <> 0 Then
            For M = 0 To flxItem.SelectedRows - 1
                'MsgBox flxItem.SelectedRow(M) & " : " & flxItem.TextMatrix(flxItem.SelectedRow(M), ItemID_Col) & " : " & str(B), vbOKOnly
                CN(7).dbOpen "UPDATE Items SET ItemID1=" & B & " WHERE ItemID=" & flxItem.TextMatrix(flxItem.SelectedRow(M), ItemID_Col), 1
            Next
        End If
        cmdRefresh_Click
    End If
    If KeyCode = vbKeyM And Shift = vbCtrlMask Then
        frmReportsFly.LoadReport "SELECT ItemID, ItemCode, ItemName, ProducerName, MRP, SRP from Items WHERE ItemID1=" & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col))
    End If
End Sub

Private Sub lblSearchCol_Change()
    lblSearchCol.FontSize = 14
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Static SearchedRow As Long
    
    If txtSearch.SelLength = 0 Then
        If (KeyCode = 38 Or KeyCode = 40) Then
            Select Case KeyCode                     'UpArrow=38 /DnArrow=40
                Case 38: If flxItem.Row > 1 Then flxItem.Row = flxItem.Row - 1
                Case 40: If flxItem.Row < flxItem.ROWS - 1 Then flxItem.Row = flxItem.Row + 1
            End Select
            SearchedRow = flxItem.Row
        Else
            SearchedRow = flxItem.FindRow(txtSearch.Text, SearchedRow + 1, Val(lblSearchCol.Caption), False, False)
        End If
    Else
        If KeyCode = 13 Then
            SearchText = txtSearch.SelText
            For i = 1 To Len(SearchText)
                searchpattern = searchpattern & "[" & LCase(Mid(SearchText, i, 1)) & UCase(Mid(SearchText, i, 1)) & "]"
            Next
            SearchText = searchpattern
            searchpattern = ""
            SearchedRow = flxItem.FindRowRegex(SearchText, SearchedRow + 1, flxItem.Col)
        End If
    End If
    
    If SearchedRow >= 1 And SearchedRow <= flxItem.ROWS - 1 Then  'Show MRP & Color
        flxItem.Row = SearchedRow: flxItem.ShowCell SearchedRow, ItemName_Col
        txtFound.Text = flxItem.TextMatrix(flxItem.Row, MRP_Col): txtFound.BackColor = vbRed
    Else
        flxItem.Row = 0: flxItem.ShowCell 0, ItemName_Col
        txtFound.Text = "0": txtFound.BackColor = vbYellow
    End If
End Sub

Private Sub txtQty_GotFocus()
    txtQty.SelLength = Len(txtQty.Text)
End Sub

Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If flxItem.Row >= 1 Then
            R = flxItem.Row
            flxItem.TextMatrix(R, Qty_Col) = Val(txtQty.Text)
            transfertext R, Val(txtQty.Text)
        End If
        txtQty.Text = "0": txtSearch.Text = "": txtSearch.SetFocus
    End If
End Sub

Private Sub cmbProducer_Click()
    If cmbProducer.Text = "*" Then
        If cmdFilter.Value = 0 Then
            CN(0).dbOpen tsql & " WHERE ItemID IN (SELECT ItemID FROM appview_INItems) " & ORDERING
        Else
            CN(0).dbOpen tsql & ORDERING
        End If
    Else
        If cmdFilter.Value = 0 Then
            CN(0).dbOpen tsql & " WHERE " & PRODUCERS & "=" & QT(Trim(cmbProducer.Text)) & " AND ItemID IN (SELECT ItemID FROM appview_INItems) " & ORDERING
        Else
            CN(0).dbOpen tsql & " WHERE " & PRODUCERS & "=" & QT(Trim(cmbProducer.Text)) & ORDERING
        End If
    End If
    Set flxItem.DataSource = CN(0).recs
End Sub
Private Sub cmdFlyReport_Click()
    frmReportsFly.LoadReport "appproc_TraceItemsForPDiscount " & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col)), 5925
End Sub

Private Sub cmdFilter_Click()
    cmbProducer_Click
End Sub

Private Sub cmbProducer_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxItem.SetFocus
        flxItem.Col = ItemName_Col:  flxItem.Sort = flexSortGenericAscending
        flxItem.Col = MRP_Col:     flxItem.Sort = flexSortGenericAscending
    End If
End Sub

Private Sub flxItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    Me.Caption = "SELECTOR > " & DestinationForm.MyCAPTION & " > " & FlxSum(Me.flxItem)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then
        SaveGrid Me.flxItem
    End If
End Sub

Private Sub flxItem_EnterCell()
    On Error Resume Next
    lblSearchCol.Caption = flxItem.Col
    If chkUpdateMode.Value = 1 Then
        If flxItem.Col = ItemName_Col Or flxItem.Col = ProducerID_Col Or flxItem.Col = MakerAuthor_Col Then
            flxItem.Editable = flexEDKbd
        End If
        If flxItem.Col = PDisc_Col Or flxItem.Col = SDisc_Col Or flxItem.Col = Stock_Col Then
            flxItem.AutoSearch = flexSearchNone
            flxItem.Editable = flexEDKbdMouse
        End If
    End If
    If flxItem.Col = Qty_Col Then
        flxItem.AutoSearch = flexSearchNone
        flxItem.Editable = flexEDKbdMouse
    End If
    If cmdShowReportsDiscount.Value = True Then frmReportsDiscount.LoadReport "appproc_TraceItemsForPDiscount " & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col)), "appproc_TraceItemsForCDiscount " & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col)) & ", " & QT(DestinationForm.txtID.Text), "appproc_TraceItemsForCDiscountGeneral " & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col))
End Sub

Private Sub flxITEM_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = Stock_Col Then
        CN(6).dbOpen "appproc_UpdateItems_InitStock " & flxItem.TextMatrix(Row, ItemID_Col) & ", " & QT(flxItem.TextMatrix(Row, Col))
    End If
    If Col = Qty_Col Then
        flxItem.TextMatrix(Row, Qty_Col) = Val(flxItem.TextMatrix(Row, Qty_Col))
        flxITEM_DblClick
        flxItem.AutoSearch = flexSearchFromCursor
    End If
End Sub

Private Sub flxITEM_LeaveCell()
    flxItem.AutoSearch = flexSearchFromCursor
    flxItem.Editable = flexEDNone
End Sub

Private Sub flxITEM_DblClick()
    rw = -1
    If flxItem.Row >= 1 Then
        R = flxItem.Row
        transfertext R, Val(flxItem.TextMatrix(R, Qty_Col))
    End If
End Sub

Private Sub txtAutoSelect_KeyPress(KeyAscii As Integer)
    Dim SELECT_TABLE As String
    SELECT_TABLE = "NONE"
    If KeyAscii = vbKeyReturn Then
        a = Split(txtAutoSelect.Text, "-")
        If UBound(a) = 1 Then
            Select Case UCase(a(0))
                Case "SA": SELECT_TABLE = "SALE"
                Case "SR": SELECT_TABLE = "SALERETURN"
                Case "PU": SELECT_TABLE = "PURCHASE"
                Case "PR": SELECT_TABLE = "PURCHASERETURN"
                Case "TO": SELECT_TABLE = "TOUT"
                Case "TI": SELECT_TABLE = "TIN"
                Case Else: SELECT_TABLE = "NONE"
            End Select
            If SELECT_TABLE <> "NONE" Then
                If UCase(Trim(txtAutoSelectProducerID.Text)) <> "ALL" Then
                    SQ = "SELECT ItemID, QTY FROM " & SELECT_TABLE & " WHERE DBREFX=" & Val(a(1)) & " AND ProducerID=" & QT(txtAutoSelectProducerID.Text)
                Else
                    SQ = "SELECT ItemID, QTY FROM " & SELECT_TABLE & " WHERE DBREFX=" & Val(a(1))
                End If
                CN(9).dbOpen SQ, 1
                If CN(9).recs.RecordCount <> 0 Then CN(9).recs.MoveFirst
                Do Until CN(9).recs.EOF
                    rw = flxItem.FindRow(CN(9).recs!ItemID, 1, ItemID_Col, False, True)
                    flxItem.Row = rw
                    If flxItem.Row >= 1 Then
                        R = flxItem.Row
                        flxItem.TextMatrix(R, Qty_Col) = Val(CN(9).recs!Qty)
                        transfertext R, Val(flxItem.TextMatrix(R, Qty_Col))
                    End If
                    CN(9).recs.MoveNext
                Loop
            End If
        End If
    End If
End Sub

Private Sub cmdQuickItem_Click()
    If flxItem.Row <> -1 Then
        frmQuickItem.GetLink Me.DestinationForm, Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col))
    Else
        frmQuickItem.GetLink Me.DestinationForm
    End If
End Sub

Private Sub cmdDuplicate_Click()
    Dim NewItemID As Long
    CN(8).dbOpen "appproc_DuplicateItem " & Val(flxItem.TextMatrix(flxItem.Row, ItemID_Col)) & ", " & Val(InputBox("Enter MRP for the duplicated Item", , "0")), 1
    Set CN(8).recs = CN(8).recs.NextRecordset
    NewItemID = CN(8).recs!NewItemID
    cmbProducer_Click
    SearchedRow = flxItem.FindRow(NewItemID, 1, ItemID_Col, True, True)
    If SearchedRow <> -1 Then
        flxItem.ShowCell SearchedRow, ItemID_Col
        flxItem.Row = SearchedRow: flxItem.Col = Qty_Col
    End If
End Sub

Private Sub cmdRefresh_Click()
    cmbProducer_Click
End Sub

Private Sub cmdTransfer_Click()
    For i = 0 To flxItem.SelectedRows - 1
        R = flxItem.SelectedRow(i)
        transfertext R, Val(flxItem.TextMatrix(R, Qty_Col))
    Next
End Sub

Private Sub transfertext(ByVal SelectedRow As Long, ByVal SelectedQty As Long)
    Dim ItemID, SearchedRow As Long
    ItemID = flxItem.TextMatrix(SelectedRow, ItemID_Col)
    
    If chkUpdateMode.Value = 0 Then
        SearchedRow = DestinationForm.flxOrder.FindRow(flxItem.TextMatrix(SelectedRow, 0), , 2, False, True)
        If SearchedRow = -1 Then
            DestinationForm.SelectItem ItemID, SelectedQty
            DestinationForm.flxOrder.ShowCell DestinationForm.flxOrder.ROWS - 1, 1
        Else
            msgUITS "Dup"
            DestinationForm.SelectItem ItemID, SelectedQty
            DestinationForm.flxOrder.ShowCell DestinationForm.flxOrder.ROWS - 1, 1
        End If
        msgUITS (SelectedQty)
    End If
End Sub

Private Sub chkItemListType_Click()
    If chkItemListType.Value = 0 Then
        tsql = "Select * from " & DestinationForm.ItemSelectView & " "
    Else
        tsql = "Select * from " & DestinationForm.ItemSelectView & " "
    End If
    cmbProducer_Click
    ColDisplay False
End Sub

Private Sub cmdQuit_Click()
    Me.ValidateControls: Unload Me
End Sub

Private Sub ColDisplay(ByVal Op As Boolean)
    If Op = True Then   'SHOW COLUMNS
        For i = 0 To flxItem.COLS - 1
            flxItem.ColHidden(i) = False
        Next
    Else                'HIDE COLUMNS
        For Each i In Split(mdiOne.sckGo.GReadINI("[SELECT-ITEM-HiddenCols]"), ",")
            If i >= 0 And i <= flxItem.COLS - 1 Then flxItem.ColHidden(i) = True
        Next
    End If
    If flxItem.COLS >= 5 Then
        flxItem.ColWidth(ItemCode_Col) = 1000
        flxItem.ColWidth(HSNCode_Col) = 1000
        flxItem.ColWidth(ItemName_Col) = 3600
        flxItem.ColWidth(ProducerName_Col) = 900
    End If
End Sub

Private Sub ColDisplayDiscountFill(ByVal Op As Boolean)
    If Op = True Then   'SHOW COLUMNS
        For i = 0 To flxItem.COLS - 1
            flxItem.ColHidden(i) = False
        Next
    Else                'HIDE COLUMNS
        For Each i In Split(mdiOne.sckGo.GReadINI("[SELECT-ITEM-DiscountFill-HiddenCols]"), ",")
            If i >= 0 And i <= flxItem.COLS - 1 Then flxItem.ColHidden(i) = True
        Next
    End If
    If flxItem.COLS >= 5 Then
        flxItem.ColWidth(ItemCode_Col) = 1000
        flxItem.ColWidth(HSNCode_Col) = 1000
        flxItem.ColWidth(ItemName_Col) = 3600
        flxItem.ColWidth(ProducerName_Col) = 900
    End If
End Sub

Private Sub flxItem_KeyDown(KeyCode%, Shift%)
        Dim Cpy As Boolean, Pst As Boolean
    ' copy: ctrl-C, ctrl-X, ctrl-ins
    If KeyCode = vbKeyC And Shift = 2 Then Cpy = True
    ' paste: ctrl-V, shift-ins
    If KeyCode = vbKeyV And Shift = 2 Then Pst = True
    
    ' do it
    If Cpy Then
        Clipboard.Clear
        Clipboard.SetText flxItem.TextMatrix(flxItem.Row, flxItem.Col)
    ElseIf Pst Then
        For i = 0 To flxItem.SelectedRows - 1
            flxItem.TextMatrix(flxItem.SelectedRow(i), flxItem.Col) = Clipboard.GetText
            flxITEM_AfterEdit flxItem.SelectedRow(i), flxItem.Col
        Next
    End If
End Sub
