VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmShowLoadMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show Load Mains..."
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8UCtl.VSFlexGrid mshf 
      Align           =   3  'Align Left
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8955
      _cx             =   15796
      _cy             =   14843
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
      MousePointer    =   1
      BackColor       =   12632256
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12632319
      ForeColorSel    =   64
      BackColorBkg    =   12632256
      BackColorAlternate=   12632256
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
      Rows            =   1
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
      AutoSearchDelay =   5
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
      WallPaperAlignment=   10
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmShowLoadMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New clsData
Private showSQL, showSQLFull, showSQLQuery As String
Private showSQLFilter, showSQLOrdering As String
Private Const MyCAPTION = "Select Window..."

Public Sub Init(Optional ByVal sql As String)
    If mshf.COLS >= 1 Then mshf.Col = 1
    Me.Move Screen.Width - Me.Width, 0
    Me.Show vbModal
End Sub

Private Sub Form_Load()
    PopulateGrid
    mdiOne.SetFormFont Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 191 And Shift = vbCtrlMask Then            '/ hide one
        mshf.ColHidden(mshf.Col) = True
        If mshf.Col < mshf.COLS - 2 Then mshf.Col = mshf.Col + 1
    End If
    If KeyCode = 188 And Shift = vbCtrlMask Then  '< HIDE ROWS
        PopulateGrid
    End If
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW ROWS
        PopulateGrid
    End If
    If KeyCode = vbKeyEscape Then        '
        Erase sArray
        Unload Me
    End If
End Sub

Private Sub PopulateGrid()
    On Error Resume Next
    Erase sArray 'To clear dbopen array filling
    If MsgBox("Do you wish to load the main report?", vbYesNo + vbQuestion) = vbYes Then
        mdiOne.CDlg.FileName = ""
        mdiOne.CDlg.Filter = "Excel Main Report |*main*.xls"
        mdiOne.CDlg.ShowOpen
        If mdiOne.CDlg.CancelError = False Then mshf.LoadGrid mdiOne.CDlg.FileName, flexFileExcel
    End If
    mshf.AutoSize 1, mshf.COLS - 1
End Sub

Private Sub mshf_DblClick()
    R = mshf.Row
    If R > 0 And mshf.ROWS > 1 Then
        Erase sArray
        For i = 0 To mshf.COLS - 1
            sArray(i) = mshf.TextMatrix(R, i)
        Next
    End If
    Unload Me
End Sub

Private Sub mshf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        R = mshf.Row
        If R > 0 Then
            Erase sArray
            For i = 0 To mshf.COLS - 1
                sArray(i) = mshf.TextMatrix(R, i)
            Next
        End If
        Unload Me
    End If
End Sub

