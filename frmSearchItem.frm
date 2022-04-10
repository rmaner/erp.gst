VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.4#0"; "UniBox10.ocx"
Begin VB.Form frmSearchItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Item..."
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmSearchItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicA 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7755
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   7815
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
         IntegralHeight  =   0   'False
         ItemData        =   "frmSearchItem.frx":114DA
         Left            =   5610
         List            =   "frmSearchItem.frx":114DC
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cmbPublisher"
         Top             =   0
         Width           =   2175
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
         Left            =   4155
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   -15
         Width           =   825
      End
      Begin UniToolbox.UniText txtSearch 
         Height          =   330
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   4125
         _Version        =   65540
         _ExtentX        =   7276
         _ExtentY        =   582
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
         AutoSize        =   -1  'True
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
         Height          =   405
         Left            =   5235
         TabIndex        =   2
         Top             =   -45
         Width           =   195
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxItem 
      Align           =   2  'Align Bottom
      Height          =   5130
      Left            =   0
      TabIndex        =   4
      Top             =   405
      Width           =   7815
      _cx             =   13785
      _cy             =   9049
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
Attribute VB_Name = "frmSearchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private CN(5) As New clsData

Dim rw As Long
Dim tflx As Control
Dim AddMode As Boolean
Dim SelectFor As String
Dim tsql As String
Dim PriceCol As Long
Public DestinationForm As frmPPSS

Public Sub GetLink(ByRef DestForm As frmPPSS)
    Set DestinationForm = DestForm
    tsql = "Select itemID, ISBN, itemName, PublisherName, Price from items "
    
    CN(0).dbOpen "SELECT DISTINCT PUBLISHERNAME FROM itemS WHERE PUBLISHERNAME IS NOT NULL ORDER BY 1", 1
    If Not CN(0).recs.EOF Then CN(0).recs.MoveFirst: cmbPublisher.Clear
    cmbPublisher.AddItem "*"
    Do Until CN(0).recs.EOF
        cmbPublisher.AddItem CN(0).recs!PublisherName
        CN(0).recs.MoveNext
    Loop
    cmbPublisher.ListIndex = 0
    lblSearchCol.Caption = 2
    flxItem.Col = 2
    flxItem.AutoSearch = flexSearchFromCursor
    flxItem.ColWidth(0) = 700: flxItem.ColWidth(1) = 1100: flxItem.ColWidth(2) = 3400: flxItem.ColWidth(3) = 1200: flxItem.ColWidth(4) = 1000
    Me.Caption = "SELECTOR > " & DestinationForm.Caption
    Me.Show vbModal
End Sub

Private Sub Form_Load()
    rw = -1: PriceCol = 4
    Me.Move Screen.Width - Me.Width, 585
    mdiOne.SetFormFont Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdQuit_Click
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Unload Me
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 38 Or KeyCode = 40) Then
        Select Case KeyCode ' UpArrow=38         DnArrow=40
            Case 38: If flxItem.Row > 1 Then flxItem.Row = flxItem.Row - 1
            Case 40: If flxItem.Row < flxItem.ROWS - 1 Then flxItem.Row = flxItem.Row + 1
        End Select
        rw = flxItem.Row
        txtSearch.Text = flxItem.TextMatrix(flxItem.Row, flxItem.Col)
    Else
        rw = flxItem.FindRow(txtSearch.Text, 1, Val(lblSearchCol.Caption), False, False)
        flxItem.Row = rw
    End If
    
    If rw >= 1 And rw <= flxItem.ROWS - 1 Then  'Show Price & Color
        txtFound.Text = flxItem.TextMatrix(flxItem.Row, PriceCol)
        txtFound.BackColor = vbRed
        flxItem.ShowCell rw, 0
    Else
        txtFound.BackColor = vbYellow: txtFound.Text = "0"
    End If
End Sub

Private Sub cmbpublisher_Click()
    If cmbPublisher.Text <> "*" Then
        CN(1).dbOpen tsql & " WHERE PublisherName=" & QT(cmbPublisher.Text) & "  ORDER BY 3 ASC", 1
    Else
        CN(1).dbOpen tsql & "  ORDER BY 3 ASC"
    End If
    Set flxItem.DataSource = CN(1).recs
End Sub

Private Sub flxItem_EnterCell()
    lblSearchCol.Caption = flxItem.Col
End Sub

Private Sub flxITEM_DblClick()
    Unload Me
End Sub

Private Sub flxITEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Unload Me
End Sub

Private Sub cmdQuit_Click()
    Me.ValidateControls: Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReturnSearch
End Sub

Private Sub ReturnSearch()
    If flxItem.Row >= 1 Then
        DestinationForm.flxOrder.Text = flxItem.TextMatrix(flxItem.Row, 0)
    Else
        DestinationForm.flxOrder.Text = 0
    End If
End Sub
