VERSION 5.00
Begin VB.Form frmHelpContents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help..."
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpContents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5250
      Left            =   -15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   7830
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   3345
      TabIndex        =   0
      Top             =   5280
      Width           =   1200
   End
End
Attribute VB_Name = "frmHelpContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    Text1.Text = mdiOne.sckGo.GReadINI("[HELPSTART]", "[HELPSTOP]")
    Text1.Text = Replace(Text1.Text, "(:-)", vbCrLf)
End Sub
