VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.4#0"; "UniBox10.ocx"
Begin VB.Form frmQuickItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QuickItem..."
   ClientHeight    =   5175
   ClientLeft      =   465
   ClientTop       =   420
   ClientWidth     =   7005
   Icon            =   "frmQuickItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   2250
      Left            =   0
      ScaleHeight     =   2190
      ScaleWidth      =   6945
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2475
      Width           =   7005
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   0
         Top             =   75
         Width           =   990
         _Version        =   65540
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Locked          =   -1  'True
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   1
         Left            =   3870
         TabIndex        =   2
         Top             =   75
         Width           =   990
         _Version        =   65540
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   3
         Left            =   1230
         TabIndex        =   4
         Top             =   417
         Width           =   5625
         _Version        =   65540
         _ExtentX        =   9922
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   10
         Left            =   1230
         TabIndex        =   11
         Top             =   1800
         Width           =   1185
         _Version        =   65540
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   4
         Left            =   1230
         TabIndex        =   5
         Top             =   765
         Width           =   1185
         _Version        =   65540
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   5
         Left            =   2460
         TabIndex        =   6
         Top             =   765
         Width           =   4425
         _Version        =   65540
         _ExtentX        =   7805
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Text            =   "0058"
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   6
         Left            =   1230
         TabIndex        =   7
         Top             =   1110
         Width           =   1185
         _Version        =   65540
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   7
         Left            =   2460
         TabIndex        =   8
         Top             =   1110
         Width           =   1185
         _Version        =   65540
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   8
         Left            =   1230
         TabIndex        =   9
         Top             =   1470
         Width           =   1185
         _Version        =   65540
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   9
         Left            =   2460
         TabIndex        =   10
         Top             =   1470
         Width           =   1185
         _Version        =   65540
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtQty 
         Height          =   315
         Left            =   6150
         TabIndex        =   14
         Top             =   1815
         Width           =   705
         _Version        =   65540
         _ExtentX        =   1244
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         BackColor       =   12648384
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   12
         Left            =   4560
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   2
         Left            =   4890
         TabIndex        =   3
         Top             =   75
         Width           =   990
         _Version        =   65540
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   11
         Left            =   2460
         TabIndex        =   12
         Top             =   1800
         Width           =   1185
         _Version        =   65540
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin VB.Label lblLabels 
         Caption         =   "ItemID1 :"
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
         Index           =   5
         Left            =   3690
         TabIndex        =   32
         Top             =   1830
         Width           =   840
      End
      Begin MSForms.ToggleButton cmdFilter 
         Height          =   360
         Left            =   6030
         TabIndex        =   21
         Top             =   15
         Width           =   900
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "1587;635"
         Value           =   "0"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Qty:"
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
         Index           =   4
         Left            =   5535
         TabIndex        =   31
         Top             =   1830
         Width           =   585
      End
      Begin VB.Label lblLabels 
         Caption         =   "PDisc/SDisc:"
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
         Index           =   12
         Left            =   15
         TabIndex        =   30
         Top             =   1500
         Width           =   1260
      End
      Begin VB.Label lblLabels 
         Caption         =   "Title:"
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
         Index           =   2
         Left            =   15
         TabIndex        =   29
         Top             =   435
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         Caption         =   "GST/Cess: "
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
         Index           =   3
         Left            =   30
         TabIndex        =   28
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         Caption         =   "ItemID:"
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
         Index           =   0
         Left            =   15
         TabIndex        =   27
         Top             =   105
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Item/HSN Code:"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   26
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label lblLabels 
         Caption         =   "Publisher:"
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
         Index           =   7
         Left            =   15
         TabIndex        =   25
         Top             =   795
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         Caption         =   "MRP/SRP:"
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
         Index           =   10
         Left            =   15
         TabIndex        =   24
         Top             =   1155
         Width           =   1125
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6945
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4725
      Width           =   7005
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   45
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   15
         Width           =   1140
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   3465
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   1140
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
         Left            =   4605
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   15
         Width           =   1140
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
         Left            =   5745
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   15
         Width           =   1140
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
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
         Left            =   1185
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   15
         Width           =   1140
      End
      Begin VB.CommandButton cmdDuplicate 
         Caption         =   "&Duplicate"
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
         Left            =   2325
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   15
         Width           =   1140
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxTable 
      Align           =   3  'Align Left
      Height          =   2475
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10110
      _cx             =   17833
      _cy             =   4366
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      AutoSearchDelay =   30
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
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
Attribute VB_Name = "frmQuickItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const FIELDS = 10
Private flds As String
Private bLoadLists As Boolean
Private CN(10) As New clsData
Public DestinationForm As frmPPSS

Public Sub GetLink(ByRef DestForm As frmPPSS, Optional ItemID As Long)
    On Error Resume Next
    Set DestinationForm = DestForm
    Me.Caption = "QuickItems... > " & DestinationForm.Caption
    CN(8).dbOpen "SELECT ProducerName FROM Items WHERE ItemID=" & ItemID, 1
    If Not CN(8).recs.EOF Then txtFields(5).Text = CN(8).recs!ProducerName
    LoadData
    Me.Show: Me.ZOrder 0
    rw = flxTable.FindRow(ItemID, , 0, False, True)
    If rw <> -1 Then
        flxTable.Row = rw
        flxTable.ShowCell flxTable.Row, 0
    End If
End Sub

Private Sub SelectItem(ByVal ItemID As Long, ByVal SelectedQty As Long)
    On Error Resume Next
    DestinationForm.SelectItem ItemID, SelectedQty
    DestinationForm.flxOrder.ShowCell DestinationForm.flxOrder.ROWS - 1, 1
    msgUITS (SelectedQty)
End Sub

Private Sub cmdFilter_Click()
    LoadData
End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    bLoadLists = True
    mdiOne.SetFormFont Me
    cmdFilter.Value = 1
    LoadData
End Sub

Private Sub LoadData()
    If cmdFilter.Value = 0 Then
        CN(0).dbOpen "SELECT ItemID, ItemCode, HSNCode, ItemName, ProducerID, ProducerName, CurrMRP, CurrSRP, PDisc, SDisc, GST, Cess, ItemID1 FROM Items ORDER BY 1 DESC"
    Else
        CN(0).dbOpen "SELECT ItemID, ItemCode, HSNCode, ItemName, ProducerID, ProducerName, CurrMRP, CurrSRP, PDisc, SDisc, GST, Cess, ItemID1 FROM Items WHERE ProducerName LIKE " & QT(txtFields(5).Text & "%") & " ORDER BY 1 DESC"
    End If
    
    Set flxTable.DataSource = CN(0)
    flxTable.DataMode = flexDMBoundImmediate
    If flxTable.ROWS > 1 Then flxTable.Row = 1
    flxTable_EnterCell
    cmdFilter.Caption = flxTable.ROWS - 1
End Sub

Private Sub Form_Activate()
    If UserRights = 0 Or UserRights = 1 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    flxTable.Width = Me.ScaleWidth
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyQ And Shift = vbCtrlMask Then
        txtQty.SetFocus
    End If
    If KeyCode = vbKeyN And Shift = vbCtrlMask Then
        cmdUpdate_Click
        cmdAdd_Click
    End If
    If KeyCode = vbKeyF12 Then cmdUpdate_Click
End Sub

Private Sub flxTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = vbCtrlMask Then
        SaveGrid Me.flxTable
    End If
End Sub

Private Sub flxTable_DblClick()
    CL = flxTable.Col
    If txtFields(CL).Enabled = True And txtFields(CL).Visible = True Then txtFields(CL).SetFocus
End Sub

Private Sub flxTable_EnterCell()
    On Error Resume Next
    For i = 0 To flxTable.COLS - 1
        txtFields(i).Text = flxTable.TextMatrix(flxTable.Row, i)
    Next
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    CN(1).dbOpen "INSERT INTO ItemS (ProducerID, ProducerName) VALUES (" & QT(txtFields(4).Text) & ", " & QT(txtFields(5).Text) & ")"
    cmdRefresh_Click
    flxTable_EnterCell
    txtFields(1).SetFocus
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    For i = 1 To flxTable.COLS - 1
        flxTable.TextMatrix(flxTable.Row, i) = txtFields(i).Text
    Next
    msgUITS "Updated"
End Sub

Private Sub cmdDuplicate_Click()
    CN(2).dbOpen "appproc_DuplicateItem " & Val(txtFields(0).Text)
    cmdRefresh_Click
End Sub

Private Sub cmdDelete_Click()
    Dim rw As Integer
    rw = flxTable.Row
    If MsgBox("Confirm deletion?", vbYesNo + vbQuestion) = vbYes Then
        CN(3).dbOpen "DELETE ItemS WHERE " & flxTable.TextMatrix(0, 0) & "=" & Val(txtFields(0).Text)
        cmdRefresh_Click
        If rw - 1 > 0 Then flxTable.Row = rw - 1
    End If
End Sub

Private Sub cmdRefresh_Click()
    CN(0).Requery
    flxTable.DataRefresh
    If flxTable.ROWS > 1 Then
        flxTable.Row = 1
        flxTable.ShowCell 1, 0
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set flxTable.DataSource = Nothing
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    txtFields(Index).SelStart = 0: txtFields(Index).SelLength = Len(txtFields(Index).Text)
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
    If Index = 4 Then
        CN(9).dbOpen "SELECT NAME FROM PERSONAL WHERE ID=" & QT(txtFields(4).Text), 1
        If Not CN(9).recs.EOF Then txtFields(5).Text = CN(9).recs!Name
    End If
End Sub

Private Sub txtFields_DblClick(Index As Integer)
    If Index = 4 Or Index = 5 Then
        frmShow.Init "SELECT ID, NAME FROM PERSONAL WHERE ID LIKE " & QT("P%") & " ORDER BY 2"
        If sArray(0) <> "" Then
            txtFields(4).Text = sArray(0)
            txtFields(5).Text = sArray(1)
        End If
    End If
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And Shift = vbCtrlMask Then
        If flxTable.Row > 1 Then flxTable.Row = flxTable.Row - 1
        flxTable.ShowCell flxTable.Row, 0
    End If
    If KeyCode = vbKeyDown And Shift = vbCtrlMask Then
        If flxTable.Row < flxTable.ROWS - 1 Then flxTable.Row = flxTable.Row + 1
        flxTable.ShowCell flxTable.Row, 0
    End If
    If KeyCode = vbKeyPageUp Then LoadData
    If KeyCode = vbKeyPageDown Then txtFields_DblClick 4
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    Static SearchedRow As Long
    If KeyAscii = vbKeyReturn And txtFields(Index).SelLength <> 0 Then
        SearchText = LTrim(txtFields(Index).SelText)
        If SearchedRow = -1 Then
            SearchedRow = flxTable.FindRow(SearchText, , Index, False, False)
        Else
            SearchedRow = flxTable.FindRow(SearchText, SearchedRow + 1, Index, False, False)
        End If
        If SearchedRow <> -1 Then
            flxTable.Row = SearchedRow
            txtFields(Index).SelStart = 0
            txtFields(Index).SelLength = Len(SearchText)
            flxTable.ShowCell SearchedRow, Index
        End If
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SelectItem Val(txtFields(0).Text), Val(txtQty.Text)
End Sub
