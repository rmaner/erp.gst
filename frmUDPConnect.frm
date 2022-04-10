VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUDPConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UDP Connect..."
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmUDPConnect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   6690
      TabIndex        =   4
      Top             =   3120
      Width           =   6750
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   6
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4815
         TabIndex        =   5
         Top             =   0
         Width           =   945
      End
      Begin MSWinsockLib.Winsock tcpServer 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSComDlg.CommonDialog CDlg 
         Left            =   405
         Top             =   -30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   1
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxRecv 
      Align           =   1  'Align Top
      Height          =   2430
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Width           =   6750
      _cx             =   11906
      _cy             =   4286
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
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
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   2
      ExplorerBar     =   0
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   6690
      TabIndex        =   1
      Top             =   0
      Width           =   6750
      Begin VB.ComboBox cmbPoints 
         Height          =   315
         Left            =   3810
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   -15
         Width           =   2895
      End
      Begin VB.TextBox txtDestinationIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Text            =   "192.168.0.201"
         Top             =   0
         Width           =   1755
      End
   End
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Text            =   "Send"
      Top             =   2775
      Width           =   6750
   End
End
Attribute VB_Name = "frmUDPConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error Resume Next
    mdiOne.SetFormFont Me
    'Me.Width = 6870: Me.Height = 3870
    tcpServer.RemotePort = 1001
    tcpServer.Bind 1001
    LoadCombo
    cmdConnect_Click
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    tcpServer.RemoteHost = txtDestinationIP.Text
    If KeyAscii = vbKeyReturn Then tcpServer.SendData txtSend.Text
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strData As String
    tcpServer.GetData strData
    flxRecv.AddItem tcpServer.RemoteHost & ":" & tcpServer.RemoteHostIP & vbTab & Format(Now, "DD-MMM HHMMSS") & vbTab & strData
    flxRecv.ShowCell flxRecv.ROWS - 1, 0
    flxRecv.Row = flxRecv.ROWS - 1
    flxRecv.AutoSize 0, flxRecv.COLS - 1
    Me.Show
    Me.SetFocus
    Me.ZOrder 0
End Sub

Private Sub cmbPoints_Click()
    S = Split(cmbPoints.Text, ":")
    txtDestinationIP.Text = S(1)
    cmdConnect_Click
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next
    tcpServer.RemoteHost = txtDestinationIP.Text
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub LoadCombo()
    cmbPoints.AddItem "GMAC01:192.168.0.2"
    cmbPoints.AddItem "STUDENTS:192.168.0.3"
    cmbPoints.AddItem "SERVER:192.168.0.201"
End Sub
