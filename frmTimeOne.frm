VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeOne 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   195
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   990
      Top             =   -150
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Max             =   800
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmTimeOne"
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
    If pb.Value > pb.Max - 30 Then Unload Me
End Sub

