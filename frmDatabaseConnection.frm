VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatabaseConnection 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Connection..."
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmDatabaseConnection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCommit 
      Caption         =   "Commit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   2
      Top             =   1530
      Width           =   1890
   End
   Begin VB.TextBox txtLogin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   810
      TabIndex        =   0
      Top             =   255
      Width           =   3390
   End
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   810
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   3390
   End
   Begin MSComctlLib.ImageList img 
      Left            =   75
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabaseConnection.frx":0482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabaseConnection.frx":0914
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabaseConnection.frx":0DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabaseConnection.frx":1238
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   30
      Top             =   1440
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Database Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Database Passwd"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   735
      Width           =   1890
   End
End
Attribute VB_Name = "frmDatabaseConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim OldDCN As String
Dim ResumeConnection As Boolean

Private Sub Form_Load()
    If conn.State = 1 Then
        conn.Close: OldDCN = DCN
        ResumeConnection = True
    Else
        ResumeConnection = False
    End If
End Sub

Private Sub cmdCommit_Click()
    On Error Resume Next
    DCN = ""
    DCN = DCN & "Provider=" & Provider & ";"
    DCN = DCN & "Persist Security Info=" & PSI & ";"
    DCN = DCN & "Initial Catalog=" & initCatalog & ";"
    DCN = DCN & "Data Source=" & DataSource & ";"
    DCN = DCN & "User ID=" & txtLogin.Text & ";"
    DCN = DCN & "Password=" & txtPasswd.Text & ";"
    
    'Set & Open connection
    If conn.State = 1 Then conn.Close
    conn.ConnectionString = DCN
    conn.Open
    If conn.State = 1 Then
        SetNewValue HKEY_LOCAL_MACHINE, sSubKey, "String", "login", txtLogin.Text
        SetNewValue HKEY_LOCAL_MACHINE, sSubKey, "String", "passwd", Cipher(txtPasswd.Text)
        Unload Me
    Else
        MsgBox Error & vbCrLf & "Unable to connect to the database. Contact administrator!", vbOKOnly + vbCritical
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 188 And Shift = vbCtrlMask Then  '<
        Timer1.Interval = Abs(Timer1.Interval - 10)
    End If
    If KeyCode = 190 And Shift = vbCtrlMask Then  '>
        Timer1.Interval = Abs(Timer1.Interval + 10)
    End If
End Sub

Private Sub Timer1_Timer()
    i = (i Mod img.ListImages.Count) + 1
    Me.Icon = img.ListImages(i).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ResumeConnection = True Then
        If conn.State = 0 Then
            DCN = OldDCN
            conn.ConnectionString = DCN
            conn.Open
        End If
    End If
End Sub

