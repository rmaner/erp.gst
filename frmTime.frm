VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTime 
   BorderStyle     =   0  'None
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   2550
      Top             =   -60
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   270
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    pb.Value = 0
    Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub Timer1_Timer()
    pb.Value = pb.Value + 20
    If pb.Value > pb.Max - 30 Then
        Unload Me
    End If
End Sub
