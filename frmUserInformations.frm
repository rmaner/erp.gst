VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmUserInformations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Informations..."
   ClientHeight    =   3255
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmUserInformations.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3750
      TabIndex        =   9
      Top             =   2835
      Width           =   1250
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2505
      TabIndex        =   8
      Top             =   2835
      Width           =   1250
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1260
      TabIndex        =   7
      Top             =   2835
      Width           =   1250
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   15
      TabIndex        =   6
      Top             =   2835
      Width           =   1250
   End
   Begin VB.PictureBox SSFrame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2805
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   4950
      TabIndex        =   10
      Top             =   -15
      Width           =   5010
      Begin VB.ComboBox cmbRights 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   1065
         TabIndex        =   5
         Top             =   2385
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   4
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2034
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         Height          =   765
         Index           =   3
         Left            =   1065
         MaxLength       =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1203
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1065
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   852
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   1
         Top             =   501
         Width           =   3330
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   0
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   0
         Top             =   150
         Width           =   3855
      End
      Begin MSForms.CommandButton cmdSelect 
         Height          =   450
         Left            =   4395
         TabIndex        =   17
         Top             =   420
         Width           =   525
         VariousPropertyBits=   8388635
         PicturePosition =   262148
         Size            =   "926;794"
         Picture         =   "frmUserInformations.frx":114DA
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Rights:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   16
         Top             =   2415
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   15
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   14
         Top             =   2055
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   13
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   12
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "LoginID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   11
         Top             =   525
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmUserInformations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(5) As New clsData

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    cmbRights.Clear
    cmbRights.AddItem "Administrator"
    cmbRights.AddItem "SuperUser"
    cmbRights.AddItem "GeneralUser"
    cmbRights.AddItem "ReportViewer"
    cmbRights.AddItem "Accounts"
End Sub

Private Sub txtFields_Change(Index As Integer)
    If Index = 1 Then
        SQ = "SELECT * FROM Usarx WHERE LoginID=" & QT(txtFields(1).Text)
        CN(1).dbOpen SQ, 1
        If Not CN(1).recs.EOF Then
            txtFields(0).Text = CN(1).recs!UserName
            txtFields(2).Text = Decipher(CN(1).recs!passwd)
            txtFields(3).Text = CN(1).recs!Address
            txtFields(4).Text = CN(1).recs!Phones
            cmbRights.ListIndex = Val(CN(1).recs!Rights)
        Else
            txtFields(0).Text = "NONE"
            txtFields(2).Text = "NONE"
            txtFields(3).Text = "NONE"
            txtFields(4).Text = "NONE"
            cmbRights.ListIndex = 4
        End If
    End If
End Sub

Private Sub cmdSelect_Click()
    frmShow.Init "SELECT * FROM Usarx ORDER BY 1 DESC"
    If sArray(1) <> "" Then
        txtFields(1).Text = sArray(1)
        txtFields_Change 1
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    X = MsgBox("Add new user?", vbYesNo)
    If X = vbYes Then
        CN(1).dbOpen "ADDNEW_USER " & QT(txtFields(1).Text), 1
        Set CN(1).recs = CN(1).recs.NextRecordset
        If Not CN(1).recs.EOF Then
            txtFields(0).Text = CN(1).recs!UserName
            txtFields(1).Text = CN(1).recs!LoginID
            txtFields(2).Text = Decipher(CN(1).recs!passwd)
            txtFields(3).Text = CN(1).recs!Address
            txtFields(4).Text = CN(1).recs!Phones
            cmbRights.ListIndex = Val(CN(1).recs!Rights)
        End If
    End If
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    X = MsgBox("Delete LoginID=" & txtFields(1).Text & "?", vbYesNo)
    If X = vbYes Then
        CN(2).dbOpen "Delete Usarx Where LoginID=" & QT(txtFields(1).Text), 1
        cmdSelect_Click
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo err_handler
    If txtFields(1).Text <> "" Then
        SQ = "Update Usarx Set UserName=" & QT(txtFields(0).Text) & ", "
        SQ = SQ & " Passwd=" & QT(Cipher(LCase(txtFields(2).Text))) & ", "
        SQ = SQ & " Address=" & QT(txtFields(3).Text) & ", "
        SQ = SQ & " Phones=" & QT(txtFields(4).Text) & ", "
        SQ = SQ & " Rights=" & cmbRights.ListIndex & " Where LoginID=" & QT(txtFields(1).Text)
        CN(3).dbOpen SQ, 1
    End If
    Exit Sub
err_handler:
    MsgBox Error, vbOKOnly + vbCritical
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
