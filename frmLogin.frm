VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1500
   ClientLeft      =   2895
   ClientTop       =   3480
   ClientWidth     =   4485
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   886.25
   ScaleMode       =   0  'User
   ScaleWidth      =   4211.172
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPasswd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1672
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   735
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1672
      TabIndex        =   0
      Top             =   345
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1012
      TabIndex        =   2
      Top             =   1125
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2452
      TabIndex        =   3
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   15
      Width           =   4440
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   240
      Index           =   1
      Left            =   487
      TabIndex        =   5
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Login ID:"
      Height          =   240
      Index           =   0
      Left            =   487
      TabIndex        =   4
      Top             =   360
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CN As New clsData

Private Sub Form_Load()
    txtPasswd.Text = ""
    Label1.Caption = "Server=" & DataSource & "; " & initCatalog
End Sub

Private Sub cmdOK_Click()
On Error GoTo Error_subLogin
    'check for correct passwd
    GUID = ""
    mdiOne.menuCLOSE
    CN.dbOpen "SELECT Passwd, Rights from Users where LoginID=" & QT(txtUserName.Text), 1
    If UCase(txtPasswd.Text) = UCase(Decipher(CN.recs!passwd)) And txtPasswd.Text <> "" Then
        msgUITS "Login successful for " & txtUserName
        GUID = UCase(txtUserName.Text): UserRights = CN.recs!Rights
        mdiOne.menuOPEN CN.recs!Rights
    Else
        MsgBox "Invalid login/passwd!", vbOKOnly + vbCritical
    End If
    mdiOne.StatusBar1.Panels(1).Text = "Current user: " & GUID
    Unload Me

Exit_subLogin:
    Exit Sub
    
Error_subLogin:
    MsgBox Error
    Resume Exit_subLogin
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
