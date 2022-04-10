VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.4#0"; "UniBox10.ocx"
Begin VB.Form frmHydSearchBills 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicA 
      Align           =   1  'Align Top
      Height          =   1110
      Left            =   0
      ScaleHeight     =   1050
      ScaleWidth      =   7755
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7815
      Begin VB.CheckBox chkExactLike 
         Caption         =   "Exact"
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
         Left            =   5775
         TabIndex        =   8
         Top             =   720
         Width           =   1890
      End
      Begin VB.ComboBox cmbFields 
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
         ItemData        =   "frmSearchBills.frx":0000
         Left            =   885
         List            =   "frmSearchBills.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cmbPublisher"
         Top             =   360
         Width           =   3825
      End
      Begin VB.ComboBox cmbTables 
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
         ItemData        =   "frmSearchBills.frx":0004
         Left            =   885
         List            =   "frmSearchBills.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cmbPublisher"
         Top             =   15
         Width           =   3825
      End
      Begin UniToolbox.UniText txtSearch 
         Height          =   330
         Left            =   885
         TabIndex        =   2
         Top             =   705
         Width           =   3825
         _Version        =   65540
         _ExtentX        =   6747
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
      Begin VB.Label Label3 
         Caption         =   "Search:"
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
         Left            =   15
         TabIndex        =   6
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Fields:"
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
         Left            =   15
         TabIndex        =   5
         Top             =   367
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Tables:"
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
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   870
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxItem 
      Align           =   2  'Align Bottom
      Height          =   5130
      Left            =   0
      TabIndex        =   7
      Top             =   1140
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
Attribute VB_Name = "frmHydSearchBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(5) As New clsData
Dim opt As Integer

Private Sub chkExactLike_Click()
    If chkExactLike.Value = 0 Then
        chkExactLike.Caption = "EXACT"
    Else
        chkExactLike.Caption = "LIKE"
    End If
    txtSearch_KeyPress vbKeyReturn
End Sub

Private Sub Form_Load()
    cmbTables.Clear
    opt = 0
    cmbTables.AddItem "PMAIN"
    cmbTables.AddItem "PRETURNMAIN"
    cmbTables.AddItem "SMAIN"
    cmbTables.AddItem "SRETURNMAIN"
    cmbTables.AddItem "TINMAIN"
    cmbTables.AddItem "TOUTMAIN"
    opt = 1
End Sub

Private Sub cmbTables_Click()
    LoadFields
End Sub

Private Sub LoadFields()
    If opt = 1 Then
        CN(0).dbOpen "SELECT TOP 1 * FROM SMAIN", 1
        cmbFields.Clear
        For i = 0 To CN(0).recs.FIELDS.Count - 1
            cmbFields.AddItem CN(0).recs.FIELDS.Item(i).Name
        Next
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Set flxItem.DataSource = Nothing
        flxItem.Clear
        If chkExactLike.Value = 1 Then
            CN(1).dbOpen "SELECT * from " & cmbTables.Text & " WHERE " & cmbFields.Text & " like " & QT("%" & Trim(txtSearch.Text) & "%"), 1
        Else
            CN(1).dbOpen "SELECT * from " & cmbTables.Text & " WHERE " & cmbFields.Text & "=" & txtSearch.Text, 1
        End If
        Set flxItem.DataSource = CN(1)
    End If
End Sub
