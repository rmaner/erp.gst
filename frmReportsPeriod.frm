VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportsPeriod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Period Reports..."
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmReportsPeriod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   8130
      Width           =   10215
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
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   30
         Width           =   2010
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
         Left            =   1485
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   30
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   405
         Left            =   5790
         ScaleHeight     =   345
         ScaleWidth      =   4335
         TabIndex        =   1
         Top             =   -15
         Width           =   4395
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
            TabIndex        =   4
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
            TabIndex        =   3
            Top             =   0
            Width           =   1440
         End
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
            TabIndex        =   2
            Top             =   0
            Width           =   1440
         End
      End
      Begin MSComCtl2.DTPicker txtDate1 
         Height          =   330
         Left            =   2145
         TabIndex        =   7
         Top             =   15
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "ddd, dd-MMM-yy"
         Format          =   80019459
         CurrentDate     =   38023
      End
      Begin MSComCtl2.DTPicker txtDate2 
         Height          =   330
         Left            =   4050
         TabIndex        =   9
         Top             =   15
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "ddd, dd-MMM-yy"
         Format          =   80019459
         CurrentDate     =   38023
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3915
         X2              =   4050
         Y1              =   180
         Y2              =   195
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   3  'Align Left
      Height          =   8130
      Left            =   0
      TabIndex        =   8
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
Attribute VB_Name = "frmReportsPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New clsData
Private Const MyCAPTION = "Period Reports..."

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    
    For Each i In Split(mdiOne.sckGo.GReadINI("[frmReportsPeriod-flxReport-SQLStrings]", "[END]"), "(:-)")
        a = Split(i, ":")
        cmbReportHead.AddItem a(0): cmbReportString.AddItem a(1)
    Next
    txtDate1.Value = Now: txtDate2.Value = Now
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
    If cmbReportHead.ListIndex <> -1 Then
        SQ = cmbReportString.List(cmbReportHead.ListIndex)
        If Right(SQ, 6) = "PERIOD" Then SQ = Replace(SQ, "PERIOD", QT(Format(txtDate1.Value, "dd-MMM-yy")) & ", " & QT(Format(txtDate2.Value, "dd-MMM-yy")))
        CN.dbOpen SQ, 1
        Set flxReport.DataSource = Nothing
        Set flxReport.DataSource = CN.recs
        Me.Caption = MyCAPTION & cmbReportHead.List(cmbReportHead.ListIndex)
    End If
End Sub

Private Sub cmdRefresh_Click()
    If cmbReportHead.ListIndex >= 0 Then cmbReportHead_Click
End Sub

Private Sub cmdSaveGrid_Click()
    flxReport.Row = 0
    If MsgBox("Do you wish to save the grid?", vbYesNo + vbQuestion) = vbYes Then
        mdiOne.CDlg.FileName = cmbReportHead.Text & " " & Format(Now, "DD-MMM-YY HHMMSS")
        mdiOne.CDlg.Filter = CompanyName & " Excel Report |*.xls"
        mdiOne.CDlg.ShowSave
        flxReport.FocusRect = flexFocusNone
        If mdiOne.CDlg.CancelError = False Then flxReport.SaveGrid mdiOne.CDlg.FileName, flexFileExcel, flexXLSaveFixedCells
    End If
End Sub

Private Sub txtDate1_Change()
    cmbReportHead_Click
End Sub

Private Sub txtDate2_Change()
    cmbReportHead_Click
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub
