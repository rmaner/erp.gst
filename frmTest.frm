VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TEMP"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmTest.frx":114DA
   ScaleHeight     =   3600
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCrypt 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5715
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   420
      Width           =   3705
   End
   Begin VB.TextBox txtString 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1845
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   420
      Width           =   3705
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtCrypt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtString.Text = Decipher(txtCrypt.Text)
    End If
End Sub

Private Sub txtString_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       txtCrypt.Text = Cipher(txtString.Text)
    End If
End Sub
