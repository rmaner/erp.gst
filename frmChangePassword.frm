VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change your password..."
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNewPassword2 
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
      IMEMode         =   3  'DISABLE
      Left            =   2070
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1065
      Width           =   2610
   End
   Begin VB.TextBox txtNewPassword1 
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
      IMEMode         =   3  'DISABLE
      Left            =   2070
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   540
      Width           =   2610
   End
   Begin VB.TextBox txtOldPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2070
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   15
      Width           =   2610
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   4620
      TabIndex        =   5
      Top             =   1545
      Width           =   4680
      Begin MSForms.CommandButton cmdQuit 
         Height          =   435
         Left            =   2475
         TabIndex        =   4
         Top             =   -15
         Width           =   1980
         Caption         =   "Quit"
         Size            =   "3492;767"
         Accelerator     =   81
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdChange 
         Height          =   435
         Left            =   285
         TabIndex        =   3
         Top             =   -15
         Width           =   1980
         Caption         =   "Change"
         Size            =   "3492;767"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Reenter new password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   30
      TabIndex        =   8
      Top             =   1065
      Width           =   2145
   End
   Begin VB.Label Label2 
      Caption         =   "Enter new password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   30
      TabIndex        =   7
      Top             =   540
      Width           =   2145
   End
   Begin VB.Label Label1 
      Caption         =   "Enter old password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   2145
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN(2) As New clsData

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
End Sub

Private Sub cmdChange_Click()
    On Error Resume Next
    CN(0).dbOpen "SELECT PASSWD FROM Usarx WHERE LOGINID=" & QT(GUID)
    If Not CN(0).recs.EOF Then
        If (UCase(Decipher(CN(0).recs!passwd)) = UCase(txtOldPassword.Text)) And (UCase(txtNewPassword1.Text) = UCase(txtNewPassword2.Text)) And txtNewPassword1.Text <> "" Then
            CN(1).dbOpen "UPDATE Usarx SET PASSWD=" & QT(Cipher(UCase(txtNewPassword1.Text))) & " WHERE LOGINID=" & QT(GUID)
            MsgBox "change password successful.", vbOKOnly + vbInformation
        Else
            MsgBox "change password failed", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

