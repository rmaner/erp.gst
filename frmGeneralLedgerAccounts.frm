VERSION 5.00
Begin VB.Form frmGeneralLedgerAccounts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Ledger Accounts..."
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   1740
      ScaleHeight     =   3150
      ScaleWidth      =   5430
      TabIndex        =   7
      Top             =   0
      Width           =   5490
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3075
         TabIndex        =   21
         Top             =   2640
         Width           =   1740
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   780
         TabIndex        =   20
         Top             =   2640
         Width           =   1740
      End
      Begin VB.CheckBox Check1 
         Height          =   330
         Left            =   2070
         TabIndex        =   19
         Top             =   2115
         Width           =   285
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2055
         TabIndex        =   18
         Top             =   1695
         Width           =   3270
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2055
         TabIndex        =   17
         Top             =   1275
         Width           =   3270
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2055
         TabIndex        =   16
         Top             =   870
         Width           =   3270
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2055
         TabIndex        =   15
         Top             =   465
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2055
         TabIndex        =   14
         Top             =   60
         Width           =   1485
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Income/ Expense"
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
         Left            =   45
         TabIndex        =   13
         Top             =   2100
         Width           =   1875
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "Account Group2"
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
         Left            =   45
         TabIndex        =   12
         Top             =   1680
         Width           =   1875
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Account Group1"
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
         Left            =   45
         TabIndex        =   11
         Top             =   1275
         Width           =   1875
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "Description"
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
         Left            =   45
         TabIndex        =   10
         Top             =   855
         Width           =   1875
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Sub-Account:"
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
         Left            =   45
         TabIndex        =   9
         Top             =   465
         Width           =   1875
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "Account Number: "
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
         Left            =   45
         TabIndex        =   8
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   0
      ScaleHeight     =   3150
      ScaleWidth      =   1650
      TabIndex        =   0
      Top             =   0
      Width           =   1710
      Begin VB.CommandButton Command6 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   6
         Top             =   2625
         Width           =   1650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   5
         Top             =   2100
         Width           =   1650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   4
         Top             =   1575
         Width           =   1650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   3
         Top             =   1050
         Width           =   1650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Sub Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   2
         Top             =   525
         Width           =   1650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmGeneralLedgerAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
End Sub
