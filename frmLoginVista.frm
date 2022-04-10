VERSION 5.00
Begin VB.Form frmLoginVista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login to DHIKTA BDMS..."
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   0
      ScaleHeight     =   1965
      ScaleWidth      =   6180
      TabIndex        =   6
      Top             =   915
      Width           =   6240
      Begin VB.CommandButton cmdDatabaseServerChange 
         Caption         =   "DB"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4620
         TabIndex        =   15
         Top             =   30
         Width           =   405
      End
      Begin VB.TextBox txtDatabaseServer 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2325
         TabIndex        =   13
         Top             =   30
         Width           =   2325
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1530
         TabIndex        =   2
         Top             =   1530
         Width           =   1485
      End
      Begin VB.ComboBox cmbDatabases 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2340
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   375
         Width           =   2700
      End
      Begin VB.TextBox txtPasswd 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2340
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1095
         Width           =   2685
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         TabIndex        =   0
         Top             =   750
         Width           =   2685
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&CANCEL"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3225
         TabIndex        =   3
         Top             =   1530
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DB Server: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1200
         TabIndex        =   14
         Top             =   45
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Databases:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1215
         TabIndex        =   9
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1215
         TabIndex        =   8
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "&Login ID:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1215
         TabIndex        =   7
         Top             =   765
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6180
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6240
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Welcome to Dhikta"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   30
         Width           =   2385
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "(Items Distribution Management System)"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "(for assistance call UITS on 98354-55022)"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   585
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmLoginVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private CN(2) As New clsData


Private Sub Form_Load()
On Error GoTo Error_subFormLoad
    Me.Caption = "Welcome to DHIKTA Business Automation Software."
        
    cmbDatabases.Clear
    txtDatabaseServer.Text = DatabaseServer
    ConnectToDatabase "master", DatabaseServer
    defaultInitCatalogFoundAt = 0
    db = QT(Left(defaultInitCatalog, 2) & "%")
    CN(0).dbOpen "SELECT * FROM SYSDATABASES WHERE Name like " & db, 1
    If Not CN(0).recs.EOF Then CN(0).recs.MoveFirst
    Do Until CN(0).recs.EOF()
        cmbDatabases.AddItem CN(0).recs!Name
        CN(0).recs.MoveNext
    Loop
    txtPasswd.Text = ""
    For i = 0 To cmbDatabases.ListCount - 1
        If UCase(Trim(cmbDatabases.List(i))) = UCase(Trim(defaultInitCatalog)) Then defaultInitCatalogFoundAt = i
    Next
    cmbDatabases.ListIndex = defaultInitCatalogFoundAt
    
    txtUserName.Text = "scott": txtPasswd.Text = "tiger"
Exit_subFormLoad:
    Exit Sub
    
Error_subFormLoad:
    MsgBox Error
    Resume Exit_subFormLoad
End Sub

Private Sub cmdDatabaseServerChange_Click()
        DatabaseServer = txtDatabaseServer.Text
        ConnectToDatabase "master", DatabaseServer
        Form_Load
End Sub

Private Sub cmbDatabases_Click()
    initCatalog = cmbDatabases.Text
    ConnectToDatabase initCatalog, DatabaseServer
    CN(1).dbOpen "SELECT KeyOne from DBKey", 1
    If CN(1).recs!KeyOne <> CodeKey Then
        MsgBox "Invalid UITS database format!" & vbCrLf & "Contact UITS support for suggestions.", vbOKOnly
        ConnectToDatabase "master", DatabaseServer
    End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo Error_subLogin
    'check for correct passwd
    GUID = ""
    mdiOne.menuCLOSE
    CN(1).dbOpen "SELECT Passwd, Rights from Usarx where LoginID=" & QT(txtUserName.Text), 1
    passwddb = UCase(Decipher(CN(1).recs!passwd))
    If UCase(txtPasswd.Text) = passwddb And txtPasswd.Text <> "" Then
        msgUITS "Login successful for " & txtUserName
        GUID = UCase(txtUserName.Text): UserRights = CN(1).recs!Rights
        mdiOne.menuOPEN CN(1).recs!Rights
        mdiOne.StatusBar1.Panels(1).Text = "Current user: " & GUID
        Unload Me
    Else
        MsgBox "Invalid login/passwd!", vbOKOnly + vbCritical
        GUID = ""
        mdiOne.StatusBar1.Panels(1).Text = "Current user: " & GUID
    End If

Exit_subLogin:
    Exit Sub
    
Error_subLogin:
    MsgBox Error
    Resume Exit_subLogin
End Sub

Private Sub cmdCancel_Click()
    GUID = ""
    mdiOne.StatusBar1.Panels(1).Text = "Current user: " & GUID
    ConnectToDatabase "master", DatabaseServer
    Unload Me
End Sub
