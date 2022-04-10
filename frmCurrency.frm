VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCurrency 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Currency Rates..."
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3060
   Icon            =   "frmCurrency.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   4290
      Width           =   3060
      Begin MSForms.CommandButton cmdDelete 
         Height          =   360
         Left            =   1500
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
         VariousPropertyBits=   19
         Caption         =   "Delete"
         Size            =   "2646;635"
         TakeFocusOnClick=   0   'False
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAdd 
         Height          =   360
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
         VariousPropertyBits=   19
         Caption         =   "Add"
         Size            =   "2646;635"
         TakeFocusOnClick=   0   'False
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
      Align           =   3  'Align Left
      Height          =   4290
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9420
      _cx             =   16616
      _cy             =   7567
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
      BackColorFixed  =   8438015
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
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
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
   End
End
Attribute VB_Name = "frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(5) As New clsData

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    LOADCURRENCIES
End Sub

Private Sub cmdAdd_Click()
    CN(1).dbOpen "INSERT INTO CURRENCY (CURRENCY) VALUES (" & QT("$") & ")"
    LOADCURRENCIES
End Sub

Private Sub cmdDelete_Click()
    flxItem.RemoveItem flxItem.Row
End Sub

Private Sub flxItem_EnterCell()
    If flxItem.Col = 2 Then
        flxItem.Editable = flexEDKbdMouse
    Else
        flxItem.Editable = flexEDNone
    End If
End Sub

Private Sub flxItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then
        SaveGrid Me.flxItem
    End If
End Sub

Private Sub LOADCURRENCIES()
    CN(0).dbOpen "SELECT * FROM Currency ORDER BY 1 ASC"
    Set flxItem.DataSource = CN(0)
    flxItem.DataMode = flexDMBoundImmediate
    flxItem.AutoSearch = flexSearchNone
    flxItem.AutoSize 0, flxItem.COLS - 1, False
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
    '    flxItem.TextMatrix(flxItem.Row, flxItem.Col) = Clipboard.GetText
    End If
End Sub



