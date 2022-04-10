VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmEditTable 
   Caption         =   "Table Editing..."
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "frmEditTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   10695
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   10635
      TabIndex        =   1
      Top             =   6960
      Width           =   10695
      Begin VB.TextBox txtCol 
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
         Height          =   285
         Left            =   9630
         TabIndex        =   6
         Text            =   "0"
         Top             =   0
         Width           =   1020
      End
      Begin VB.TextBox txtSql 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5325
         TabIndex        =   4
         Top             =   0
         Width           =   4050
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   630
         TabIndex        =   2
         Top             =   0
         Width           =   4050
      End
      Begin VB.Label Label2 
         Caption         =   "Sql:"
         Height          =   240
         Left            =   5040
         TabIndex        =   5
         Top             =   30
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Search:"
         Height          =   240
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   840
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxItem 
      Align           =   3  'Align Left
      Height          =   6960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _cx             =   18865
      _cy             =   12277
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
      BackColorFixed  =   8454143
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
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
      Begin VB.Timer tmrAuto 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
   End
End
Attribute VB_Name = "frmEditTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(2) As New clsData

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    txtSql.Text = "SELECT itemID, itemNAME, PUBLISHERNAME, PDISCOUNT, SDISCOUNT, PRICE, DISCGROUP FROM itemS ORDER BY PUBLISHERNAME, itemNAME, PDISCOUNT"
    LoadSql txtSql.Text
End Sub

Private Sub Form_Resize()
    flxItem.Width = Me.ScaleWidth
End Sub

Private Sub flxItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Me.Caption = "Table Editor (Row: " & NewRow & "  Col: " & NewCol & ")"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'enable/ disable timer
    If KeyCode = vbKeyF12 And Shift = vbCtrlMask + vbAltMask + vbShiftMask Then
        tmrAuto.Enabled = Not tmrAuto.Enabled
        If tmrAuto.Enabled Then
            flxItem.BackColorAlternate = vbYellow
        Else
            flxItem.BackColorAlternate = vbWhite
        End If
    End If
End Sub

Private Sub tmrAuto_Timer()
    LoadSql txtSql.Text
End Sub

Private Sub flxItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then SaveGrid Me.flxItem
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Static SearchedRow As Long
    If KeyCode = 13 Then
        If SearchedRow = -1 Then
            SearchedRow = flxItem.FindRow(Trim(txtSearch.Text), , flxItem.Col, False, False)
        Else
            SearchedRow = flxItem.FindRow(Trim(txtSearch.Text), SearchedRow + 1, Val(flxItem.Col), False, False)
        End If
        If SearchedRow <> -1 Then
            flxItem.Row = SearchedRow
            flxItem.ShowCell SearchedRow, Val(flxItem.Col)
            txtSearch.BackColor = vbCyan
        End If
    Else
        txtSearch.BackColor = vbWhite
    End If
    txtSearch.SetFocus
End Sub

Private Sub txtSql_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LoadSql txtSql.Text
    flxItem.AutoSearch = flexSearchNone
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set flxItem.DataSource = Nothing
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
            If flxItem.Col = Val(txtCol.Text) Then flxItem.TextMatrix(flxItem.SelectedRow(i), flxItem.Col) = Clipboard.GetText
        Next
    End If
End Sub

Private Sub LoadSql(ByVal SQ As String)
    On Error Resume Next
    CN(0).dbOpen SQ, 1
    Set flxItem.DataSource = CN(0).recs
    flxItem.DataMode = flexDMBoundImmediate
    flxItem.Editable = flexEDKbdMouse
End Sub

