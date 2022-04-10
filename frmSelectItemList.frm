VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSelectItemList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Item for ItemList..."
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmSelectItemList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicB 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   8970
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7200
      Width           =   9030
      Begin VB.CommandButton cmdHide 
         Caption         =   "&Hide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Top             =   15
         Width           =   1410
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6105
         TabIndex        =   7
         Top             =   15
         Width           =   1440
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&AddNew"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1440
      End
   End
   Begin VB.PictureBox PicA 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8970
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9030
      Begin VB.TextBox txtQty 
         Height          =   315
         Left            =   3105
         TabIndex        =   1
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   15
         Width           =   2235
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
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   4275
         TabIndex        =   3
         Top             =   -45
         Width           =   615
      End
      Begin MSForms.TextBox txtPBID 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxItem 
      Align           =   1  'Align Top
      Height          =   6735
      Left            =   0
      TabIndex        =   9
      Top             =   375
      Width           =   9030
      _cx             =   15928
      _cy             =   11880
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
      AutoSearch      =   0
      AutoSearchDelay =   2
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
Attribute VB_Name = "frmSelectItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rw As Long
Dim tflx As Control
Dim tsql As String

Private Sub Form_Load()
    Me.Move Screen.Width - Me.Width, 0
    mdiOne.SetFormFont Me
    
    lblSearchCol.Caption = 3
    flxItem.FontName = ReadFont("[frmSelectitem-flxItem-Font]", 0): flxItem.FontSize = ReadFont("[frmSelectitem-flxItem-Font]", 1): flxItem.FontBold = ReadFont("[frmSelectitem-flxItem-Font]", 2)
    sSQL(6) = tsql: dbOpen (6): Set flxItem.DataSource = recs(6)
End Sub

Public Sub GetLink(ByRef F As Control, ByVal paramSQL As String)
    AddMode = pAddMode: Set tflx = F
    tsql = paramSQL
    frmSelectItemList.Show vbModal
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 188 And Shift = vbCtrlMask Then  '< HIDE
        For Each i In Split(mdiOne.sckGo.GReadINI("[frmSelectitem-flxItem-HiddenCols]"), ",")
            flxItem.ColHidden(i) = True
        Next
    End If
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW
        For i = 0 To flxItem.COLS - 1
            flxItem.ColHidden(i) = False
        Next
    End If
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Private Sub txtPBID_Change()
    rw = flxItem.FindRow(txtPBID.Text, 1, Val(lblSearchCol.Caption), False, False)
    If rw >= 1 Then
        flxItem.Row = rw
        flxItem.ShowCell rw, 0
    End If
End Sub

Private Sub txtPBID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyReturn Then
        flxItem.Row = rw
        If flxItem.Row >= 1 Then
            R = flxItem.Row: flxItem.TextMatrix(R, 9) = Val(txtQty.Text)
            S = flxItem.Cell(flexcpText, R, 0, R, flxItem.COLS - 1)
            Call transfertext(tflx, S, R)
        End If
        txtPBID.Text = "": txtQty.Text = "0"
    End If
End Sub

Private Sub txtQty_GotFocus()
    txtQty.SelLength = Len(txtQty.Text)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxItem.Row = rw
        If flxItem.Row >= 1 Then
            R = flxItem.Row: flxItem.TextMatrix(R, 10) = Val(txtQty.Text)
            S = flxItem.Cell(flexcpText, R, 0, R, flxItem.COLS - 1)
            Call transfertext(tflx, S, R)
        End If
        txtPBID.Text = "": txtQty.Text = "0": txtPBID.SetFocus
    End If
End Sub

Private Sub flxITEM_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    flxITEM_DblClick
End Sub

Private Sub flxItem_EnterCell()
    lblSearchCol.Caption = flxItem.Col: lblSearch.Caption = "Search on " & flxItem.ColKey(flxItem.Col)
    If flxItem.Col = 9 Then flxItem.Editable = flexEDKbdMouse
End Sub

Private Sub flxITEM_LeaveCell()
    flxItem.Editable = flexEDNone
End Sub

Private Sub flxITEM_DblClick()
    If flxItem.Row >= 1 Then
        R = flxItem.Row
        S = flxItem.Cell(flexcpText, R, 0, R, flxItem.COLS - 1)
        Call transfertext(tflx, S, R)
    End If
End Sub

Private Sub flxITEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If flxItem.Row >= 1 Then
            R = flxItem.Row
            S = flxItem.Cell(flexcpText, R, 0, R, flxItem.COLS - 1)
            Call transfertext(tflx, S, R)
        End If
    End If
End Sub

Private Sub cmdAddNew_Click()
    sSQL(0) = "Insert into itemS (itemName) Values (" & QT("New item") & ")"
    dbOpen (0): dbClose (0)
    
End Sub

Private Sub cmdSelect_Click()
    For i = 0 To flxItem.SelectedRows - 1
        R = flxItem.SelectedRow(i)
        S = flxItem.Cell(flexcpText, R, 0, R, flxItem.COLS - 1)
        Call transfertext(tflx, S, R)
    Next
End Sub

Public Sub transfertext(ByRef anyControl As Control, ByVal a As String, ByVal R As Long)
    Dim irow As Long
    irow = anyControl.FindRow(flxItem.TextMatrix(R, 1), , 1, False, True)
    'X = flxItem.FindRow("ABC", R, C, CASESENSITIVE, FULLMATCH)
    
    If irow = -1 Then
        anyControl.AddItem a
        anyControl.ShowCell anyControl.ROWS - 1, 1
    Else
        msgUITS "Dup"
        anyControl.ShowCell irow, 3: anyControl.Row = irow
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dbClose (6)
End Sub

Private Sub cmdHide_Click()
    Me.ValidateControls: Me.Hide
End Sub


