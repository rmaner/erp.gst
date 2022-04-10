VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReportProfit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Profit Report..."
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10080
      TabIndex        =   0
      Top             =   8400
      Width           =   10140
      Begin VB.PictureBox Picture2 
         Height          =   405
         Left            =   5700
         ScaleHeight     =   345
         ScaleWidth      =   4335
         TabIndex        =   1
         Top             =   -15
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
            TabIndex        =   2
            Top             =   0
            Width           =   1440
         End
      End
      Begin MSComCtl2.DTPicker txtDate1 
         Height          =   330
         Left            =   -15
         TabIndex        =   5
         Top             =   15
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   82378755
         CurrentDate     =   38023
      End
      Begin MSComCtl2.DTPicker txtDate2 
         Height          =   330
         Left            =   1980
         TabIndex        =   6
         Top             =   15
         Width           =   1830
         _ExtentX        =   3228
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
         Format          =   82378755
         CurrentDate     =   38023
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1755
         X2              =   1890
         Y1              =   180
         Y2              =   195
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   3  'Align Left
      Height          =   8400
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10215
      _cx             =   18018
      _cy             =   14817
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
Attribute VB_Name = "frmReportProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New clsData
Private Const MyCAPTION = "Profit Report..."

Private Sub flxReport_DblClick()
    On Error GoTo ERROR_REPORT
    SQ = "appproc_ProfitReport_itemWise " & QT(Format(txtDate1.Value, "dd-MMM-yy")) & ", " & QT(Format(txtDate2.Value, "dd-MMM-yy")) & ", " & QT(flxReport.Text)
    CN.dbOpen SQ, 1
    Set flxReport.DataSource = Nothing
    Set flxReport.DataSource = CN.recs
    flxReport.ColFormat(flxReport.COLS - 1) = "###.00"
    Exit Sub
ERROR_REPORT:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
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

Private Sub REPORT()
    On Error Resume Next
    SQ = "appproc_ProfitReport_PublisherWise " & QT(Format(txtDate1.Value, "dd-MMM-yy")) & ", " & QT(Format(txtDate2.Value, "dd-MMM-yy"))
    CN.dbOpen SQ, 1
    Set flxReport.DataSource = Nothing
    Set flxReport.DataSource = CN.recs
    flxReport.ColFormat(3) = "###.00"
End Sub

Private Sub flxReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = MyCAPTION & FlxSum(Me.flxReport)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then SaveGrid Me.flxReport
End Sub

Private Sub txtDate1_Change()
    REPORT
End Sub

Private Sub txtDate2_Change()
    REPORT
End Sub

Private Sub cmdRefresh_Click()
    REPORT
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

