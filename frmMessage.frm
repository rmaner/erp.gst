VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom  message box (the message you are seeing, is also on clipboard) ..."
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   6495
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   0
      Width           =   8820
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   8880
      ScaleHeight     =   6495
      ScaleWidth      =   930
      TabIndex        =   0
      Top             =   0
      Width           =   930
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    mdiOne.SetFormFont Me
End Sub

Public Sub GetMessage(msg As String)
    Clipboard.Clear
    Clipboard.SetText msg
    txtMessage.Text = msg
    Me.Show
    Me.ZOrder 0
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
