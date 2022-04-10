VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportsDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATE REPORT WINDOW..."
   ClientHeight    =   8400
   ClientLeft      =   150
   ClientTop       =   -30
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportsDate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   13140
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   13080
      TabIndex        =   0
      Top             =   7965
      Width           =   13140
      Begin VB.PictureBox Picture2 
         Height          =   405
         Left            =   10125
         ScaleHeight     =   345
         ScaleWidth      =   2895
         TabIndex        =   6
         Top             =   0
         Width           =   2955
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
            Left            =   0
            TabIndex        =   8
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
            Left            =   1455
            TabIndex        =   7
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.CheckBox chkMonthly 
         Caption         =   "Monthly"
         Height          =   270
         Left            =   5835
         TabIndex        =   5
         Top             =   60
         Width           =   1350
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
         Left            =   2910
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   30
         Visible         =   0   'False
         Width           =   495
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
         Width           =   3720
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   330
         Left            =   3750
         TabIndex        =   4
         Top             =   30
         Width           =   1995
         _ExtentX        =   3519
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
         CustomFormat    =   "ddd, dd-MMM-yyyy"
         Format          =   80019459
         CurrentDate     =   38023
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   1  'Align Top
      Height          =   7935
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13140
      _cx             =   23177
      _cy             =   13996
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
Attribute VB_Name = "frmReportsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MyCAPTION = "Date Reports..."
Private MyData(5) As New clsData
Private sqlStr As String

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    Me.Move 0, 0
    For Each i In Split(mdiOne.sckGo.GReadINI("[frmReportsDate-flxReport-SQLStrings]", "[END]"), "(:-)")
        a = Split(i, ":")
        cmbReportHead.AddItem a(0): cmbReportString.AddItem a(1)
    Next
    txtDate.Value = Now
End Sub

Private Sub cmbReportHead_Click()
    If cmbReportHead.ListIndex <> -1 Then
        sqlStr = cmbReportString.List(cmbReportHead.ListIndex)
        
        If Right(sqlStr, 4) = "DATE" Then sqlStr = Replace(sqlStr, "DATE", QT(Format(txtDate.Value, "dd-MMM-yy")))
        If chkMonthly.Value = 1 Then sqlStr = sqlStr & ", " & QT("M")
        MyData(0).dbOpen sqlStr
        Set flxReport.DataSource = Nothing
        flxReport.Clear
        Set flxReport.DataSource = MyData(0)
        Me.Caption = MyCAPTION & cmbReportHead.List(cmbReportHead.ListIndex)
        For i = 0 To flxReport.COLS - 1
            If flxReport.ColDataType(i) = flexDTDate Then flxReport.ColFormat(i) = "dd-mmm-yy"
        Next
    End If
    flxReport.AutoSize 0, flxReport.COLS - 1
End Sub

Private Sub chkMonthly_Click()
    cmbReportHead_Click
End Sub

Private Sub cmdRefresh_Click()
    If cmbReportHead.ListIndex >= 0 Then cmbReportHead_Click
End Sub

Private Sub flxReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim sum As Double
    sum = 0
    For i = 0 To flxReport.SelectedRows - 1
        If flxReport.SelectedRow(i) >= 1 Then
            sum = sum + Val(flxReport.TextMatrix(flxReport.SelectedRow(i), flxReport.Col))
        End If
        Me.Caption = MyCAPTION & "Sum on " & str(flxReport.Col) & " = " & Format(sum, "##,##0.00")
    Next
    If Button = 2 And Shift = vbCtrlMask Then
        SaveGrid Me.flxReport, cmbReportHead.Text & Format(Now, "DD-MMM-YY HHMM")
    End If
End Sub

Private Sub txtDate_Change()
    cmbReportHead_Click
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub
