VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmReportsGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERAL REPORT WINDOW..."
   ClientHeight    =   8565
   ClientLeft      =   150
   ClientTop       =   -30
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportsGen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   10215
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   8130
      Width           =   10215
      Begin VB.PictureBox Picture2 
         Height          =   405
         Left            =   5775
         ScaleHeight     =   345
         ScaleWidth      =   4335
         TabIndex        =   4
         Top             =   0
         Width           =   4395
         Begin VB.CommandButton cmdSaveGrid 
            Caption         =   "SaveGrid"
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
            TabIndex        =   7
            Top             =   0
            Width           =   1440
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
            Left            =   1440
            TabIndex        =   6
            Top             =   0
            Width           =   1440
         End
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
            Left            =   2895
            TabIndex        =   5
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.ComboBox cmbReportString 
         CausesValidation=   0   'False
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
         Left            =   4710
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   30
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox cmbReportHead 
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
         Left            =   0
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   30
         Width           =   4695
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   3  'Align Left
      Height          =   8130
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10215
      _cx             =   18018
      _cy             =   14340
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
      MultiTotals     =   -1  'True
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
      Editable        =   2
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
Attribute VB_Name = "frmReportsGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New clsData
Private Const MyCAPTION = "General Reports... "

Private Sub Form_Load()
    On Error Resume Next
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    
    For Each i In Split(mdiOne.sckGo.GReadINI("[frmReportsGeneral-flxReport-SQLStrings]", "[END]"), "(:-)")
        a = Split(i, ":")
        cmbReportHead.AddItem a(0): cmbReportString.AddItem a(1)
    Next
    'cmbReportHead.ListIndex = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW
        For i = 0 To flxReport.COLS - 1
            flxReport.ColHidden(i) = False
        Next
    End If
    If KeyCode = 191 And Shift = vbCtrlMask Then            '/ hide one
        flxReport.ColHidden(flxReport.Col) = True
        If flxReport.Col < flxReport.COLS - 2 Then flxReport.Col = flxReport.Col + 1
    End If
End Sub

Private Sub flxReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = MyCAPTION & FlxSum(Me.flxReport)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then SaveGrid Me.flxReport
End Sub

Private Sub cmbReportHead_Click()
    CN.dbOpen (cmbReportString.List(cmbReportHead.ListIndex))
    Set flxReport.DataSource = Nothing
    Set flxReport.DataSource = CN
    flxReport.GridLines = flexGridInset
    Me.Caption = MyCAPTION & cmbReportHead.List(cmbReportHead.ListIndex)
End Sub

Private Sub cmdSaveGrid_Click()
    flxReport.Row = 0
    If MsgBox("Do you wish to save the grid?", vbYesNo + vbQuestion) = vbYes Then
        mdiOne.CDlg.FileName = Me.Caption & " " & Format(Now, "DD-MMM-YY HHMMSS")
        mdiOne.CDlg.Filter = CompanyName & " Excel Report |*.xls"
        mdiOne.CDlg.ShowSave
        flxReport.FocusRect = flexFocusNone
        If mdiOne.CDlg.CancelError = False Then flxReport.SaveGrid mdiOne.CDlg.FileName, flexFileExcel, flexXLSaveFixedCells
    End If
End Sub

Private Sub cmdRefresh_Click()
    cmbReportHead_Click
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Picture1_DblClick()
    SQ = InputBox("ENTER SQL", "ON THE FLY REPORTS...")
    LoadReport SQ
End Sub

Public Sub LoadReport(ByVal SQ As String)
    On Error Resume Next
    CN.dbOpen SQ
    Set flxReport.DataSource = Nothing
    Set flxReport.DataSource = CN
    flxReport.GridLines = flexGridInset
    Me.Caption = MyCAPTION & "ON THE FLY REPORTS"
End Sub

Private Sub FLXREPORT_KeyDown(KeyCode%, Shift%)
        Dim Cpy As Boolean, Pst As Boolean
    ' copy: ctrl-C, ctrl-X, ctrl-ins
    If KeyCode = vbKeyC And Shift = 2 Then Cpy = True
    If KeyCode = vbKeyX And Shift = 2 Then Cpy = True
    If KeyCode = vbKeyInsert And Shift = 2 Then Cpy = True

    ' paste: ctrl-V, shift-ins
    If KeyCode = vbKeyV And Shift = 2 Then Pst = True
    If KeyCode = vbKeyInsert And Shift = 1 Then Pst = True
    ' do it
    If Cpy Then
        Clipboard.Clear
        Clipboard.SetText flxReport.Clip
    ElseIf Pst Then
        For i = 0 To flxReport.SelectedRows - 1
            flxReport.TextMatrix(flxReport.SelectedRow(i), flxReport.Col) = Clipboard.GetText
        Next
    End If
End Sub



