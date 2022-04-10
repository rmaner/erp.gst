VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPPSS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PPSS..."
   ClientHeight    =   9060
   ClientLeft      =   105
   ClientTop       =   345
   ClientWidth     =   14325
   Icon            =   "frmPPSS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pboxK 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   14265
      TabIndex        =   58
      Top             =   8610
      Width           =   14325
      Begin VB.CommandButton cmdPrintLong 
         Caption         =   "Prnt&T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10590
         TabIndex        =   146
         Top             =   15
         Width           =   660
      End
      Begin VB.CommandButton cmdPrintSmall 
         Caption         =   "PrntH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11235
         TabIndex        =   130
         Top             =   15
         Width           =   660
      End
      Begin VB.CommandButton cmdCancelBill 
         Caption         =   "C&ancelBill"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2100
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   0
         Width           =   1050
      End
      Begin VB.CommandButton cmdDeleteBill 
         Caption         =   "&DeleteBill"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1050
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
         Width           =   1050
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13005
         TabIndex        =   51
         Top             =   15
         Width           =   1245
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&PrintF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11925
         TabIndex        =   50
         Top             =   15
         Width           =   990
      End
      Begin VB.CommandButton cmdMemoFinalize 
         Caption         =   "&OnAccount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   9435
         TabIndex        =   49
         Top             =   15
         Width           =   1050
      End
      Begin VB.CommandButton cmdMemoFinalize 
         Caption         =   "Cas&h"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8325
         TabIndex        =   48
         Top             =   15
         Width           =   1110
      End
      Begin VB.CommandButton cmdMemoFinalize 
         Caption         =   "C&hallan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7215
         TabIndex        =   47
         Top             =   15
         Width           =   1110
      End
      Begin VB.CommandButton cmdSaveOrder 
         Caption         =   "&SaveOrder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6105
         TabIndex        =   46
         Top             =   15
         Width           =   1110
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "&Calculate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   1050
      End
      Begin MSForms.TextBox txtUserNo 
         Height          =   315
         Left            =   3195
         TabIndex        =   124
         Top             =   45
         Width           =   2850
         VariousPropertyBits=   746604567
         ForeColor       =   32768
         BorderStyle     =   1
         Size            =   "5027;556"
         Value           =   "- by "
         BorderColor     =   -2147483640
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Memo Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   15
      TabIndex        =   54
      Top             =   0
      Width           =   3000
      Begin VB.PictureBox pboxA 
         Height          =   2100
         Left            =   60
         ScaleHeight     =   2040
         ScaleWidth      =   2835
         TabIndex        =   55
         Top             =   195
         Width           =   2895
         Begin VB.TextBox txtStatus 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
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
            Left            =   -15
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "STATUS"
            Top             =   1215
            Width           =   2865
         End
         Begin VB.TextBox txtOrderRef 
            Height          =   330
            Left            =   600
            TabIndex        =   2
            Top             =   420
            Width           =   915
         End
         Begin VB.TextBox txtDBRef 
            Height          =   330
            Left            =   600
            MousePointer    =   10  'Up Arrow
            TabIndex        =   0
            ToolTipText     =   "Ctrl+0"
            Top             =   45
            Width           =   915
         End
         Begin MSComCtl2.DTPicker txtOrderDate 
            Height          =   330
            Left            =   1530
            TabIndex        =   3
            Top             =   420
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MMM-yy"
            Format          =   20316163
            CurrentDate     =   38023
         End
         Begin MSComCtl2.DTPicker txtDBDate 
            Height          =   330
            Left            =   1530
            TabIndex        =   1
            Top             =   45
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MMM-yy"
            Format          =   20316163
            CurrentDate     =   38023
         End
         Begin MSComCtl2.DTPicker txtInvDate 
            Height          =   345
            Left            =   1530
            TabIndex        =   5
            Top             =   795
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MMM-yy"
            Format          =   20316163
            CurrentDate     =   38023
         End
         Begin VB.TextBox txtInvRef 
            Height          =   345
            Left            =   600
            TabIndex        =   4
            Top             =   795
            Width           =   915
         End
         Begin MSForms.CommandButton cmdSelectDBRef 
            Height          =   360
            Left            =   1410
            TabIndex        =   89
            Top             =   1605
            Width           =   1440
            VariousPropertyBits=   8388635
            Caption         =   "OPEN"
            PicturePosition =   327683
            Size            =   "2540;635"
            Accelerator     =   79
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdNewDBRef 
            Height          =   360
            Left            =   0
            TabIndex        =   90
            Top             =   1605
            Width           =   1425
            VariousPropertyBits=   8388635
            Caption         =   "NEW"
            PicturePosition =   327683
            Size            =   "2514;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            Caption         =   "Inv:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   30
            TabIndex        =   83
            Top             =   855
            Width           =   810
         End
         Begin VB.Label lblOrder 
            Caption         =   "Order:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   15
            TabIndex        =   57
            Top             =   465
            Width           =   810
         End
         Begin VB.Label Label3 
            Caption         =   "DBRef:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   15
            TabIndex        =   56
            Top             =   105
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox pboxI 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   14265
      TabIndex        =   52
      Top             =   8190
      Width           =   14325
      Begin VB.TextBox txtComments 
         Height          =   315
         Left            =   1020
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Width           =   10785
      End
      Begin VB.Label Label9 
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   30
         TabIndex        =   53
         Top             =   30
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab sstOne 
      Height          =   2355
      Left            =   3045
      TabIndex        =   7
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   4154
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   450
      BackColor       =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1 Customer && Postage"
      TabPicture(0)   =   "frmPPSS.frx":4E0E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCustomer"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraPostage"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2 Transport && Good Recv"
      TabPicture(1)   =   "frmPPSS.frx":4E2A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraGoods"
      Tab(1).Control(1)=   "fraTransporter"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&3 Karter && Trace"
      TabPicture(2)   =   "frmPPSS.frx":4E46
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraKarter"
      Tab(2).Control(1)=   "fraSubDist"
      Tab(2).ControlCount=   2
      Begin VB.Frame fraGoods 
         Caption         =   "Good Package Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -70710
         TabIndex        =   102
         Top             =   45
         Width           =   4500
         Begin VB.PictureBox pboxF 
            Height          =   1305
            Left            =   75
            ScaleHeight     =   1245
            ScaleWidth      =   4320
            TabIndex        =   103
            Top             =   225
            Width           =   4380
            Begin VB.TextBox txtTerminal 
               Height          =   285
               Left            =   2985
               TabIndex        =   109
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtGRNo 
               Height          =   300
               Left            =   645
               TabIndex        =   108
               Top             =   0
               Width           =   1395
            End
            Begin VB.TextBox txtBundleWeight 
               Height          =   285
               Left            =   2985
               TabIndex        =   107
               Top             =   645
               Width           =   1335
            End
            Begin VB.TextBox txtBundleCount 
               Height          =   285
               Left            =   2985
               TabIndex        =   106
               Top             =   330
               Width           =   1335
            End
            Begin VB.ComboBox txtGRMode 
               Height          =   315
               Left            =   645
               TabIndex        =   105
               Top             =   615
               Width           =   1395
            End
            Begin VB.ComboBox txtToPayMode 
               Height          =   315
               Left            =   2985
               TabIndex        =   104
               Top             =   0
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker txtGRDate 
               Height          =   300
               Left            =   645
               TabIndex        =   110
               Top             =   300
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   20316163
               CurrentDate     =   38023
            End
            Begin MSMask.MaskEdBox txtGRAmount 
               Height          =   300
               Left            =   645
               TabIndex        =   111
               Top             =   945
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               Caption         =   "To Pay:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2040
               TabIndex        =   119
               Top             =   0
               Width           =   930
            End
            Begin VB.Label Label47 
               Caption         =   "GRNo:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   -15
               TabIndex        =   118
               Top             =   30
               Width           =   750
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Bundle:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2040
               TabIndex        =   117
               Top             =   315
               Width           =   930
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               Caption         =   "Weight:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2040
               TabIndex        =   116
               Top             =   660
               Width           =   930
            End
            Begin VB.Label Label4 
               Caption         =   "Mode:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   -15
               TabIndex        =   115
               Top             =   645
               Width           =   750
            End
            Begin VB.Label Label5 
               Caption         =   "GRDt:"
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
               Left            =   -15
               TabIndex        =   114
               Top             =   315
               Width           =   750
            End
            Begin VB.Label Label6 
               Caption         =   "GRAmt:"
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
               Left            =   -15
               TabIndex        =   113
               Top             =   960
               Width           =   750
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               Caption         =   "Terminal:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2040
               TabIndex        =   112
               Top             =   960
               Width           =   930
            End
         End
      End
      Begin VB.Frame fraKarter 
         Caption         =   "Karter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   -74925
         TabIndex        =   92
         Top             =   60
         Width           =   7425
         Begin VB.PictureBox pboxE 
            Height          =   1695
            Left            =   75
            ScaleHeight     =   1635
            ScaleWidth      =   7170
            TabIndex        =   93
            Top             =   180
            Width           =   7230
            Begin VB.TextBox txtKName 
               Height          =   315
               Left            =   1410
               TabIndex        =   97
               Top             =   45
               Width           =   5715
            End
            Begin VB.TextBox txtKID 
               Height          =   315
               Left            =   15
               TabIndex        =   96
               Top             =   45
               Width           =   1005
            End
            Begin VB.TextBox txtKAddress 
               Height          =   885
               Left            =   15
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   95
               Top             =   405
               Width           =   7110
            End
            Begin VB.CommandButton cmdSelectKID 
               DownPicture     =   "frmPPSS.frx":4E62
               Height          =   315
               Left            =   1020
               Picture         =   "frmPPSS.frx":51A5
               Style           =   1  'Graphical
               TabIndex        =   94
               ToolTipText     =   "Ctrl+4"
               Top             =   30
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin MSMask.MaskEdBox txtKAmount 
               Height          =   285
               Left            =   5835
               TabIndex        =   98
               Top             =   1335
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               _Version        =   393216
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtKRate 
               Height          =   270
               Left            =   1080
               TabIndex        =   99
               Top             =   1335
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   476
               _Version        =   393216
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
               Caption         =   "Karter Rate:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   30
               TabIndex        =   101
               Top             =   1335
               Width           =   1605
            End
            Begin VB.Label Label8 
               Caption         =   "Karter Amount:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   4440
               TabIndex        =   100
               Top             =   1335
               Width           =   1500
            End
         End
      End
      Begin VB.Frame fraTransporter 
         Caption         =   "Transporter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -74940
         TabIndex        =   66
         Top             =   45
         Width           =   4215
         Begin VB.PictureBox pboxD 
            Height          =   1335
            Left            =   45
            ScaleHeight     =   1275
            ScaleWidth      =   4050
            TabIndex        =   67
            Top             =   195
            Width           =   4110
            Begin VB.CommandButton cmdSelectTID 
               DownPicture     =   "frmPPSS.frx":54E8
               Height          =   285
               Left            =   825
               Picture         =   "frmPPSS.frx":582B
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Ctrl+3"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.TextBox txtTID 
               Height          =   285
               Left            =   0
               TabIndex        =   19
               Top             =   -15
               Width           =   825
            End
            Begin VB.TextBox txtTCity 
               Height          =   285
               Left            =   0
               TabIndex        =   23
               Top             =   990
               Width           =   1350
            End
            Begin VB.TextBox txtTPhones 
               Height          =   285
               Left            =   1365
               TabIndex        =   24
               Top             =   990
               Width           =   1350
            End
            Begin VB.TextBox txtTName 
               Height          =   285
               Left            =   1170
               TabIndex        =   21
               Top             =   -15
               Width           =   2880
            End
            Begin VB.TextBox txtTaddress 
               Height          =   690
               Left            =   0
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   285
               Width           =   4050
            End
            Begin VB.TextBox txtTEmail 
               Height          =   285
               Left            =   2730
               TabIndex        =   25
               Top             =   990
               Width           =   1350
            End
         End
      End
      Begin VB.Frame fraPostage 
         Caption         =   "Post/Courier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   7770
         TabIndex        =   64
         Top             =   45
         Width           =   3330
         Begin VB.PictureBox pboxC 
            Height          =   825
            Left            =   60
            ScaleHeight     =   765
            ScaleWidth      =   3120
            TabIndex        =   65
            Top             =   210
            Width           =   3180
            Begin VB.CommandButton cmdSelectPID 
               DownPicture     =   "frmPPSS.frx":5B6E
               Height          =   315
               Left            =   825
               Picture         =   "frmPPSS.frx":5EB1
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Ctrl+2"
               Top             =   30
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.TextBox txtPName 
               Enabled         =   0   'False
               Height          =   345
               Left            =   15
               TabIndex        =   17
               TabStop         =   0   'False
               ToolTipText     =   "SD NAME"
               Top             =   405
               Width           =   3090
            End
            Begin VB.TextBox txtPID 
               Height          =   330
               Left            =   30
               TabIndex        =   15
               ToolTipText     =   "SUB DISTRIBUTOR"
               Top             =   30
               Width           =   825
            End
         End
      End
      Begin VB.Frame fraSubDist 
         Caption         =   "Trace Items..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   -67410
         TabIndex        =   62
         Top             =   60
         Width           =   3585
         Begin VB.CheckBox chkGeneralTracing 
            Alignment       =   1  'Right Justify
            Caption         =   "General Trace of items"
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
            Left            =   90
            TabIndex        =   129
            Top             =   285
            Width           =   2295
         End
         Begin VB.CheckBox chkIncludeID 
            Alignment       =   1  'Right Justify
            Caption         =   "IncludeID:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2160
            TabIndex        =   127
            Top             =   870
            Width           =   1200
         End
         Begin VB.ComboBox txtTraceIn 
            Height          =   315
            Left            =   930
            TabIndex        =   125
            Top             =   855
            Width           =   1170
         End
         Begin VB.PictureBox pboxG 
            Height          =   60
            Left            =   30
            ScaleHeight     =   0
            ScaleWidth      =   3465
            TabIndex        =   63
            Top             =   720
            Width           =   3525
            Begin VB.CommandButton cmdSelectSDID 
               DownPicture     =   "frmPPSS.frx":61F4
               Height          =   285
               Left            =   1035
               Picture         =   "frmPPSS.frx":6537
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Ctrl+5"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.TextBox txtSDName 
               Enabled         =   0   'False
               Height          =   285
               Left            =   0
               TabIndex        =   28
               TabStop         =   0   'False
               ToolTipText     =   "SD NAME"
               Top             =   330
               Width           =   2280
            End
            Begin VB.TextBox txtSDID 
               Height          =   285
               Left            =   0
               TabIndex        =   26
               ToolTipText     =   "SUB DISTRIBUTOR"
               Top             =   0
               Width           =   1020
            End
         End
         Begin MSForms.CommandButton cmdTrace 
            Height          =   360
            Left            =   90
            TabIndex        =   128
            Top             =   1200
            Width           =   3405
            VariousPropertyBits=   8388635
            Caption         =   "TRACE"
            PicturePosition =   327683
            Size            =   "6006;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Trace In: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   60
            TabIndex        =   126
            Top             =   900
            Width           =   885
         End
      End
      Begin VB.Frame fraCustomer 
         Caption         =   "Party Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   75
         TabIndex        =   60
         Top             =   45
         Width           =   7680
         Begin VB.PictureBox pboxB 
            Height          =   1650
            Left            =   90
            ScaleHeight     =   1590
            ScaleWidth      =   7455
            TabIndex        =   61
            Top             =   210
            Width           =   7515
            Begin VB.TextBox txtState 
               Height          =   330
               Left            =   2910
               TabIndex        =   137
               ToolTipText     =   "CITY"
               Top             =   1095
               Width           =   630
            End
            Begin VB.TextBox txtShipState 
               Height          =   330
               Left            =   6810
               TabIndex        =   136
               ToolTipText     =   "CITY"
               Top             =   1080
               Width           =   630
            End
            Begin VB.TextBox txtShipAddress 
               Height          =   330
               Left            =   3735
               TabIndex        =   14
               ToolTipText     =   "SHIPPING ADDRESS"
               Top             =   1080
               Width           =   3060
            End
            Begin VB.TextBox txtPhones 
               Height          =   315
               Left            =   4770
               TabIndex        =   12
               ToolTipText     =   "PHONES"
               Top             =   390
               Width           =   2670
            End
            Begin VB.TextBox txtCity 
               Height          =   330
               Left            =   30
               TabIndex        =   11
               ToolTipText     =   "CITY"
               Top             =   1095
               Width           =   2880
            End
            Begin VB.TextBox txtEmail 
               Height          =   315
               Left            =   4755
               TabIndex        =   13
               ToolTipText     =   "EMAIL"
               Top             =   735
               Width           =   2670
            End
            Begin VB.TextBox txtName 
               Height          =   315
               Left            =   1500
               TabIndex        =   9
               ToolTipText     =   "NAME"
               Top             =   30
               Width           =   5910
            End
            Begin VB.TextBox txtAddress 
               Height          =   660
               Left            =   30
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               ToolTipText     =   "ADDRESS"
               Top             =   390
               Width           =   4710
            End
            Begin VB.TextBox txtID 
               Height          =   315
               Left            =   30
               TabIndex        =   8
               ToolTipText     =   "ID"
               Top             =   30
               Width           =   1035
            End
            Begin MSForms.CommandButton cmdSelectID 
               Height          =   405
               Left            =   1080
               TabIndex        =   91
               ToolTipText     =   "Ctrl+1"
               Top             =   0
               Width           =   405
               VariousPropertyBits=   8388635
               PicturePosition =   262148
               Size            =   "714;714"
               Picture         =   "frmPPSS.frx":687A
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current Ledger Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7785
         TabIndex        =   59
         Top             =   1260
         Width           =   3315
         Begin MSMask.MaskEdBox txtAccountBalance 
            Height          =   360
            Left            =   75
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   635
            _Version        =   393216
            BorderStyle     =   0
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   " #,##0.00 Dr; #,##0.00 Cr;#,##0.00"
            PromptChar      =   "_"
         End
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   15
      TabIndex        =   68
      Top             =   2385
      Width           =   14295
      Begin VB.PictureBox Picture1 
         Height          =   420
         Left            =   60
         ScaleHeight     =   360
         ScaleWidth      =   14085
         TabIndex        =   131
         Top             =   4290
         Width           =   14145
         Begin VB.CheckBox chkReturnCalculator 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "[ ReturnAssist: ]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6030
            TabIndex        =   135
            Top             =   60
            Width           =   5865
         End
         Begin MSForms.CommandButton CommandButton1 
            Height          =   345
            Left            =   2790
            TabIndex        =   134
            Top             =   0
            Width           =   540
            VariousPropertyBits=   8388635
            Caption         =   "A"
            PicturePosition =   327683
            Size            =   "952;609"
            Accelerator     =   36
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton CommandButton2 
            Height          =   345
            Left            =   3360
            TabIndex        =   133
            Top             =   0
            Width           =   525
            VariousPropertyBits=   8388635
            Caption         =   "B"
            PicturePosition =   327683
            Size            =   "926;609"
            Accelerator     =   36
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdExtraDiscountCreditNote 
            Height          =   360
            Left            =   0
            TabIndex        =   132
            Top             =   0
            Width           =   2445
            VariousPropertyBits=   8388635
            Caption         =   "ExtraDiscountCreditNote"
            PicturePosition =   327683
            Size            =   "4313;635"
            Accelerator     =   36
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.PictureBox pboxH 
         Height          =   420
         Left            =   60
         ScaleHeight     =   360
         ScaleWidth      =   14085
         TabIndex        =   69
         Top             =   3870
         Width           =   14145
         Begin MSMask.MaskEdBox txtItemAmount 
            Height          =   360
            Left            =   12420
            TabIndex        =   30
            Top             =   0
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdQuickItem 
            Caption         =   "Q&uickItem"
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
            Left            =   1365
            TabIndex        =   120
            Top             =   15
            Width           =   1380
         End
         Begin MSMask.MaskEdBox txtNetGSTAmt 
            Height          =   360
            Left            =   10140
            TabIndex        =   144
            Top             =   0
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "GST:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   9330
            TabIndex        =   145
            Top             =   60
            Width           =   765
         End
         Begin MSForms.CommandButton cmdDeleteItem 
            Height          =   360
            Left            =   2760
            TabIndex        =   86
            Top             =   0
            Width           =   1380
            VariousPropertyBits=   8388635
            Caption         =   "DELETE ITEM"
            PicturePosition =   327683
            Size            =   "2434;635"
            Accelerator     =   68
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdLoadGrid 
            Height          =   360
            Left            =   5550
            TabIndex        =   88
            Top             =   0
            Width           =   1380
            VariousPropertyBits=   8388635
            Caption         =   "Load Grid"
            PicturePosition =   327683
            Size            =   "2434;635"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSaveGrid 
            Height          =   360
            Left            =   4155
            TabIndex        =   87
            Top             =   0
            Width           =   1380
            VariousPropertyBits=   8388635
            Caption         =   "Save Grid"
            PicturePosition =   393216
            Size            =   "2434;635"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtItemCount 
            Height          =   330
            Left            =   7800
            TabIndex        =   123
            Top             =   15
            Width           =   1020
            VariousPropertyBits=   746604567
            ForeColor       =   32768
            BorderStyle     =   1
            Size            =   "1799;582"
            Value           =   "0"
            BorderColor     =   -2147483640
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label8 
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   11865
            TabIndex        =   70
            Top             =   60
            Width           =   765
         End
         Begin MSForms.CommandButton cmdSelectItem 
            Height          =   360
            Left            =   -15
            TabIndex        =   85
            Top             =   0
            Width           =   1380
            VariousPropertyBits=   8388635
            Caption         =   "SELECT ITEM"
            PicturePosition =   327683
            Size            =   "2434;635"
            Accelerator     =   69
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid flxOrder 
         Height          =   3690
         Left            =   60
         TabIndex        =   29
         Top             =   165
         Width           =   14145
         _cx             =   24950
         _cy             =   6509
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   1
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632319
         ForeColorSel    =   64
         BackColorBkg    =   16777215
         BackColorAlternate=   16776960
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   4
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   3
         Rows            =   1
         Cols            =   50
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   1
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   10
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame Frame1 
      DragMode        =   1  'Automatic
      Height          =   1110
      Left            =   30
      TabIndex        =   71
      Top             =   7080
      Width           =   14235
      Begin VB.PictureBox pboxJ 
         Height          =   900
         Left            =   75
         ScaleHeight     =   840
         ScaleWidth      =   14010
         TabIndex        =   72
         Top             =   150
         Width           =   14070
         Begin MSMask.MaskEdBox txtAddFreight 
            Height          =   270
            Left            =   4020
            TabIndex        =   35
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLessMisc 
            Height          =   270
            Left            =   6765
            TabIndex        =   39
            Top             =   555
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtNetAmount 
            Height          =   270
            Left            =   12405
            TabIndex        =   40
            Top             =   285
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   476
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLessFreight 
            Height          =   270
            Left            =   4020
            TabIndex        =   36
            Top             =   570
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAddMisc 
            Height          =   270
            Left            =   6765
            TabIndex        =   38
            Top             =   270
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtRoundOff 
            Height          =   270
            Left            =   12405
            TabIndex        =   41
            Top             =   570
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   476
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSDCommission 
            Height          =   270
            Left            =   1035
            TabIndex        =   32
            Top             =   285
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPostage 
            Height          =   270
            Left            =   1035
            TabIndex        =   33
            Top             =   555
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtBulkDisc 
            Height          =   270
            Left            =   6765
            TabIndex        =   37
            Top             =   -15
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSplDisc 
            Height          =   270
            Left            =   4020
            TabIndex        =   34
            Top             =   -15
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtNetCessAmt 
            Height          =   270
            Left            =   1035
            TabIndex        =   31
            Top             =   0
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDiscountAmt 
            Height          =   270
            Left            =   12405
            TabIndex        =   121
            Top             =   0
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   476
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtNetSGSTAmt 
            Height          =   270
            Left            =   9150
            TabIndex        =   138
            Top             =   285
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtNetIGSTAmt 
            Height          =   270
            Left            =   9150
            TabIndex        =   139
            Top             =   570
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtNetCGSTAmt 
            Height          =   270
            Left            =   9150
            TabIndex        =   140
            Top             =   0
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "IGST: "
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
            Left            =   8400
            TabIndex        =   143
            Top             =   585
            Width           =   780
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "SGST: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   8400
            TabIndex        =   142
            Top             =   315
            Width           =   780
         End
         Begin VB.Label lblCGST 
            Alignment       =   1  'Right Justify
            Caption         =   "CGST:  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   8400
            TabIndex        =   141
            Top             =   30
            Width           =   780
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "DiscountAmt: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   11235
            TabIndex        =   122
            Top             =   30
            Width           =   1200
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "NetCess:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -30
            TabIndex        =   84
            Top             =   30
            Width           =   1020
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Special Disc:"
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
            Left            =   2775
            TabIndex        =   82
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Bulk Disc:"
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
            Left            =   5550
            TabIndex        =   81
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            Caption         =   "Net Amount: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   11235
            TabIndex        =   80
            Top             =   315
            Width           =   1200
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            Caption         =   "-Misc Amt:"
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
            Left            =   5550
            TabIndex        =   79
            Top             =   570
            Width           =   1200
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Freight:"
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
            Left            =   2790
            TabIndex        =   78
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            Caption         =   "Less Freight:"
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
            Left            =   2820
            TabIndex        =   77
            Top             =   585
            Width           =   1185
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Caption         =   "+Misc Amt:"
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
            Left            =   5550
            TabIndex        =   76
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "RoundOff: "
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
            Left            =   11235
            TabIndex        =   75
            Top             =   585
            Width           =   1200
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "+ Postage:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -30
            TabIndex        =   74
            Top             =   555
            Width           =   1020
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "SDComm:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -30
            TabIndex        =   73
            Top             =   300
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "frmPPSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(25) As New clsData
Private cJr As New clsJournals

Public AutoCalculate As Boolean
Public Discs As Boolean
Public MyCAPTION As String
Public StockValidationRequired As Boolean
Public HiddenCols As String
Public Main As String
Public Support As String
Public CounterMain As String
Public CounterSupport As String
Public MainAddNew As String
Public MainSelectView As String
Public PartyView As String
Public PartyInitial As String
Public ItemSelectView As String
Public MemoFormat As String

Public ReturnSAStartDate As String
Public ReturnSAEndDate As String
Public ReturnSRStartDate As String
Public ReturnSREndDate As String

Public ReturnPUStartDate As String
Public ReturnPUEndDate As String
Public ReturnPRStartDate As String
Public ReturnPREndDate As String

Public Serial_Col, DBRefX_Col, ItemID_Col, ItemCode_Col, HSNCode_Col, ItemName_Col, MakerAuthor_Col, ProducerID_Col, ProducerName_Col, Version_Col, MfgDate_Col, ExpDate_Col, Packing_Col, Unit_Col, Qty_Col, Free_Col, MRP_Col, SRP_Col, Gross_Col, aDisc_Col, aDiscAmt_Col, bDisc_Col, bDiscAmt_Col, cDisc_Col, cDiscAmt_Col, GST_Col, GSTAmt_Col, Cess_Col, CessAmt_Col, Amount_Col, Stock_Col As Integer
Private DBRef As Long

Private Sub cmdTrace_Click()
    frmMessage.GetMessage Traceitems
End Sub

Private Sub flxOrder_AfterDataRefresh()
    'Calling ReturnAssist for sale/purchase return limits validation
    ReturnAssist
End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    Me.Caption = MyCAPTION
    mdiOne.SetFormFont Me
    
    Serial_Col = 0: DBRefX_Col = 1: ItemID_Col = 2: ItemCode_Col = 3: HSNCode_Col = 4: ItemName_Col = 5: MakerAuthor_Col = 6: ProducerID_Col = 7: ProducerName_Col = 8: Version_Col = 9: MfgDate_Col = 10: ExpDate_Col = 11: Packing_Col = 12: Unit_Col = 13: Qty_Col = 14: Free_Col = 15: MRP_Col = 16: SRP_Col = 17: Gross_Col = 18: aDisc_Col = 19: aDiscAmt_Col = 20: bDisc_Col = 21: bDiscAmt_Col = 22: cDisc_Col = 23: cDiscAmt_Col = 24: GST_Col = 25: GSTAmt_Col = 26: Cess_Col = 27: CessAmt_Col = 28: Amount_Col = 29: Stock_Col = 30
    
    txtDBDate.Value = Now: txtOrderDate.Value = Now: txtInvDate.Value = Now
    txtGRMode.AddItem "DIRECT": txtGRMode.AddItem "BANK": txtGRMode.AddItem "HOLD"
    txtToPayMode.AddItem "Paid-Full": txtToPayMode.AddItem "Paid-Half": txtToPayMode.AddItem "Paid-Zero": txtToPayMode.AddItem "ToPay-Full": txtToPayMode.AddItem "ToPay-Half": txtToPayMode.AddItem "ToPay-Zero"
    txtTraceIn.Clear: txtTraceIn.AddItem "SA": txtTraceIn.AddItem "SR": txtTraceIn.AddItem "PU": txtTraceIn.AddItem "PR": txtTraceIn.AddItem "TI": txtTraceIn.AddItem "TO": txtTraceIn.ListIndex = 0
    
    ReturnSAStartDate = mdiOne.sckGo.GReadINI("[SABIG-PERIOD-START]"): ReturnSAEndDate = mdiOne.sckGo.GReadINI("[SABIG-PERIOD-END]")
    ReturnSRStartDate = mdiOne.sckGo.GReadINI("[SRBIG-PERIOD-START]"): ReturnSREndDate = mdiOne.sckGo.GReadINI("[SRBIG-PERIOD-END]")
    ReturnPUStartDate = mdiOne.sckGo.GReadINI("[PUBIG-PERIOD-START]"): ReturnPUEndDate = mdiOne.sckGo.GReadINI("[PUBIG-PERIOD-END]")
    ReturnPRStartDate = mdiOne.sckGo.GReadINI("[PRBIG-PERIOD-START]"): ReturnPREndDate = mdiOne.sckGo.GReadINI("[PRBIG-PERIOD-END]")
    
    Call PictureBoxStatus(False)
    
End Sub

Private Sub Form_Activate()
    If UserRights = 0 Or UserRights = 1 Then
        cmdDeleteBill.Enabled = True
    Else
        cmdDeleteBill.Enabled = False
    End If
    
    If UserRights = 0 Then
        cmdExtraDiscountCreditNote.Enabled = True
    Else
        cmdExtraDiscountCreditNote.Enabled = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 188 And Shift = vbCtrlMask Then Call ColShowHide(1)             '< Hide
    If KeyCode = 190 And Shift = vbCtrlMask Then Call ColShowHide(0)             '> Show
    If KeyCode = 191 And Shift = vbCtrlMask Then                                 '/ hide one
        flxOrder.ColHidden(flxOrder.Col) = True
        If flxOrder.Col < flxOrder.COLS - 2 Then flxOrder.Col = flxOrder.Col + 1
    End If
    If KeyCode = vbKeyD And Shift = vbCtrlMask Then ApplyDiscounts "SPL"                'Special Discount
    If KeyCode = vbKeyD And Shift = vbCtrlMask + vbShiftMask Then ApplyDiscounts "GEN"  'General Discount
    If KeyCode = vbKeyK And Shift = vbCtrlMask Then SaveDiscounts               'Save Discounts
    If KeyCode = vbKeyU And Shift = vbCtrlMask Then SavePersonalSettings        'Save Personal Settings
    If KeyCode = vbKeyM And Shift = vbCtrlMask Then MergeRows                   'Merge Rows
    If KeyCode = vbKeyN And Shift = vbCtrlMask Then Call AddBlankRows           'Add Blank Rows
    If KeyCode = vbKeyF12 Then cmdSaveOrder_Click
    If KeyCode = vbKeyT And Shift = vbCtrlMask + vbShiftMask Then
        frmReportsFly.LoadReport "appproc_Traceitems " & Val(flxOrder.TextMatrix(flxOrder.Row, ItemID_Col)) & ", " & QT(Format(txtDBDate.Value, "dd-MMM-yyyy"))
    End If
    If KeyCode = vbKeyT And Shift = vbCtrlMask Then
        frmReportsFly.LoadReport "appproc_TraceitemsForPDiscount " & Val(flxOrder.TextMatrix(flxOrder.Row, ItemID_Col)), 5925
    End If
    If KeyCode = 38 And Shift = vbAltMask Then  'up arrow
        SQ = "SELECT ItemID, ItemCode, ItemName, AUTHORS, PUBNAME, PRICE FROM " & ItemSelectView & " WHERE ItemCode=" & QT(flxOrder.TextMatrix(flxOrder.Row, flxOrder.Col)) & " ORDER BY 1 Desc"
        frmShow.Init SQ
        If Val(sArray(0)) <> 0 Then
            PopulateGridRow Val(sArray(0))
        End If
    End If
    If (KeyCode = 96 Or KeyCode = 48) And Shift = vbCtrlMask Then txtDBRef.SetFocus      'ZERO
    If (KeyCode = 97 Or KeyCode = 49) And Shift = vbCtrlMask Then cmdSelectID_Click      'ONE
    If (KeyCode = 98 Or KeyCode = 50) And Shift = vbCtrlMask Then cmdSelectPID_Click     'TWO
    If (KeyCode = 99 Or KeyCode = 51) And Shift = vbCtrlMask Then cmdSelectTID_Click     'THREE
    If (KeyCode = 100 Or KeyCode = 52) And Shift = vbCtrlMask Then cmdSelectKID_Click    'FOUR
    If (KeyCode = 101 Or KeyCode = 53) And Shift = vbCtrlMask Then txtDBRef.SetFocus     'FIVE
    If (KeyCode = 102 Or KeyCode = 54) And Shift = vbCtrlMask Then txtDBRef.SetFocus     'SIX
End Sub

Private Sub txtDBDate_DblClick()
    txtOrderDate.Value = txtDBDate.Value: txtInvDate.Value = txtDBDate.Value
End Sub

Private Sub txtDBRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtDBRef.Text = Val(txtDBRef.Text) + 1
        txtDBRef_LostFocus
    End If
    If KeyCode = vbKeyDown Then
        txtDBRef.Text = Val(txtDBRef.Text) - 1
        txtDBRef_LostFocus
    End If
End Sub

Public Sub txtDBRef_LostFocus()
    CN(0).dbOpen "Select * from " & Main & " where DBRef=" & Val(txtDBRef.Text)
    DBRef = Val(sArray(0))
    If DBRef = 0 Then
        PictureBoxStatus (False): txtDBRef = ""
    End If
    If Val(sArray(0)) <> 0 Then
        FillFormTextFromMain 'Filling of Forms's text boxes.
        CN(1).dbOpen "Select * from " & Support & " where DBRefX=" & Val(txtDBRef.Text) & " Order by Serial", 1
        Set flxOrder.DataSource = CN(1).recs
        'ColShowHide (1)
        PictureBoxStatus (True)
    End If
    
    ' Col Width
    flxOrder.ColWidth(Serial_Col) = 350
    flxOrder.ColWidth(ItemID_Col) = 450
    flxOrder.ColWidth(ItemName_Col) = 3600
    flxOrder.ColWidth(ProducerName_Col) = 1000
    flxOrder.ColWidth(Qty_Col) = 500
    flxOrder.ColWidth(MRP_Col) = 500
    flxOrder.ColWidth(SRP_Col) = 500
    flxOrder.ColWidth(Gross_Col) = 900
    flxOrder.ColWidth(aDisc_Col) = 900
    flxOrder.ColWidth(aDiscAmt_Col) = 900
    flxOrder.ColWidth(Amount_Col) = 1000
    ColShowHide (1)
    sstOne.Tab = 0
    'Calculate
    
    ' Col Format
    'flxOrder.ColFormat(aDisc_Col) = "###0.00%"
    flxOrder.ColFormat(aDiscAmt_Col) = "###0.00"

    txtStatus.Text = GetOrderStatus(Main, Val(txtDBRef.Text))
    txtAccountBalance.Text = WhatIsLedgerBalance(txtID.Text)
End Sub

Private Sub cmdNewDBRef_Click()
    X = MsgBox("Create new order?", vbYesNo + vbQuestion)
    If X = vbYes Then
        txtDBDate.Value = Now: txtInvDate.Value = Now
        CN(2).dbOpen MainAddNew & " " & QT(GUID), 1
        Set CN(2).recs = CN(2).recs.NextRecordset
        txtDBRef.Text = CN(2).recs!DBRef
        If CN(2).recs!DBRef Mod 10 = 0 Then
            mdiOne.sckGo.LogIt
        End If
    End If
    
    txtDBRef_LostFocus
End Sub

Private Sub cmdSelectDBRef_Click()
    SQ = "SELECT * FROM " & MainSelectView & " Where DateDiff(D, DBDate," & QT(Format(Now, "DD-MMM-YY")) & ")=0 ORDER BY 1 Desc"
    frmShow.Init SQ
    If sArray(0) <> "" Then
        txtDBRef.Text = sArray(0)
    End If
    txtDBRef_LostFocus
End Sub

Private Sub txtID_LostFocus()
    txtID.Enabled = True
    CN(4).dbOpen "Select * from " & PartyView & " WHERE ID=" & QT(txtID.Text), 0
    txtID.Text = sArray(0): txtName.Text = sArray(1): txtAddress.Text = sArray(2): txtCity.Text = sArray(3): txtState.Text = sArray(4): txtPhones.Text = sArray(5): txtEmail.Text = sArray(6): txtShipAddress.Text = sArray(7)
    txtTerminal.Text = txtCity.Text
    txtAccountBalance.Text = WhatIsLedgerBalance(txtID.Text)
End Sub

Private Sub cmdSelectID_Click()
    SQ = "Select * from " & PartyView & " WHERE ID LIKE " & QT(PartyInitial) & " ORDER BY 2"
    frmShow.Init SQ
    If sArray(0) <> "" Then
        txtID.Text = sArray(0)
        txtID_LostFocus
    End If
End Sub

Private Sub txtPID_LostFocus()
    CN(7).dbOpen "Select * from " & PartyView & " WHERE ID=" & QT(txtPID.Text), 0
    txtPID.Text = sArray(0): txtPName.Text = sArray(1)
End Sub

Private Sub cmdSelectPID_Click()
    SQ = "Select * from " & PartyView & " WHERE ID LIKE " & QT("O%") & " ORDER BY 2"
    frmShow.Init SQ
    If sArray(0) <> "" Then
        txtPID.Text = sArray(0)
        txtPID_LostFocus
    End If
End Sub

Private Sub txtStatus_Change()
    If txtStatus = "NEW" Then txtStatus.BackColor = vbWhite
    If txtStatus = "ORDER" Then txtStatus.BackColor = vbYellow
    If txtStatus = "CANCELLED" Then txtStatus.BackColor = vbBlue
    If txtStatus = "CHALLAN" Then txtStatus.BackColor = vbCyan
    If txtStatus = "CASH" Then txtStatus.BackColor = vbGreen
    If txtStatus = "CREDIT" Then txtStatus.BackColor = vbRed
End Sub

Private Sub txtTID_LostFocus()
    CN(7).dbOpen "Select * from " & PartyView & " WHERE ID=" & QT(txtTID.Text), 0
    txtTID.Text = sArray(0): txtTName.Text = sArray(1): txtTaddress.Text = sArray(2): txtTCity.Text = sArray(3): txtTPhones.Text = sArray(4): txtTEmail.Text = sArray(5)
End Sub

Private Sub cmdSelectTID_Click()
    SQ = "Select * from " & PartyView & " WHERE ID LIKE " & QT("T%") & " ORDER BY 2"
    frmShow.Init SQ
    If sArray(0) <> "" Then
        txtTID.Text = sArray(0)
        txtTID_LostFocus
    End If
End Sub

Private Sub txtKID_LostFocus()
    CN(7).dbOpen "Select * from " & PartyView & " WHERE ID=" & QT(txtKID.Text), 0
    txtKID.Text = sArray(0): txtKName.Text = sArray(1): txtKAddress.Text = sArray(2)
End Sub

Private Sub cmdSelectKID_Click()
    SQ = "Select * from " & PartyView & " WHERE ID LIKE " & QT("K%") & " ORDER BY 2"
    frmShow.Init SQ
    If sArray(0) <> "" Then
        txtKID.Text = sArray(0)
        txtKID_LostFocus
    End If
End Sub

Private Sub txtSDID_LostFocus()
    CN(7).dbOpen "Select * from " & PartyView & " WHERE ID=" & QT(txtSDID.Text), 0
    txtSDID.Text = sArray(0): txtSDName.Text = sArray(1)
End Sub

Private Sub cmdSelectSDID_Click()
    SQ = "Select * from " & PartyView & " WHERE ID LIKE " & QT("U%") & " ORDER BY 2"
    frmShow.Init SQ
    If sArray(0) <> "" Then
        txtSDID.Text = sArray(0)
        txtSDID_LostFocus
    End If
End Sub

Private Sub cmbGRMode_Click()
    txtGRMode.Text = cmbGRMode.Text
End Sub

Private Sub txtGRAmount_Change()
    txtToPayMode_Click
End Sub

Private Sub txtToPayMode_Click()
    Select Case txtToPayMode.ListIndex
        Case 0
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
        Case 1
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0.5: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
        Case 2
            txtAddFreight.Text = Val(txtGRAmount.Text) * 1: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
        Case 3
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0: txtLessFreight.Text = Val(txtGRAmount.Text) * 1
        Case 4
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0: txtLessFreight.Text = Val(txtGRAmount.Text) * 0.5
        Case 5
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
    End Select
End Sub

Private Sub chkGeneralTracing_Click()
    If chkGeneralTracing.Value = 1 Then
        frmReportsDiscount.Show
    Else
        frmReportsDiscount.Hide
    End If
End Sub

Private Sub flxOrder_DblClick()
    If flxOrder.Col <> Qty_Col Then cmdQuickItem_Click
End Sub

' Making flxOrder Grid Editable
Private Sub flxOrder_EnterCell()
    With flxOrder
        c = .Col
        Select Case c
            Case ItemID_Col, ItemCode_Col, Version_Col, SRP_Col, Qty_Col, Free_Col, aDisc_Col, aDiscAmt_Col, bDisc_Col, bDiscAmt_Col, cDisc_Col, cDiscAmt_Col, Cess_Col, Amount_Col
                .Editable = flexEDKbdMouse
            Case Else
                .Editable = flexEDNone
        End Select
    End With
    If chkGeneralTracing.Value = 1 Then
        frmReportsDiscount.LoadReport "appproc_TraceitemsForPDisc " & Val(flxOrder.TextMatrix(flxOrder.Row, ItemID_Col)), "appproc_TraceitemsForCDiscount " & Val(flxOrder.TextMatrix(flxOrder.Row, ItemID_Col)) & ", " & QT(txtID.Text), "appproc_TraceitemsForCDiscountGeneral " & Val(flxOrder.TextMatrix(flxOrder.Row, ItemID_Col))
    End If
    
    'Calling ReturnAssist for sale/purchase return limits validation
    ReturnAssist
End Sub

Private Sub flxOrder_LeaveCell()
    With flxOrder
    For R = 1 To .ROWS - 1
        For c = 0 To .COLS - 1
            Select Case c
                Case Qty_Col, Free_Col, SRP_Col, Gross_Col, aDiscAmt_Col, bDiscAmt_Col, cDiscAmt_Col, GSTAmt_Col, CessAmt_Col, Amount_Col, Stock_Col
                    .TextMatrix(R, c) = Val(.TextMatrix(R, c))
            End Select
        Next
        If StockValidationRequired Then
            If (Val(.TextMatrix(R, Qty_Col)) + Val(.TextMatrix(R, Free_Col))) > Val(.TextMatrix(R, Stock_Col)) Then
                .Cell(flexcpBackColor, R, Qty_Col, R, Free_Col) = vbRed
            Else
                .Cell(flexcpBackColor, R, Qty_Col, R, Free_Col) = .Cell(flexcpBackColor, R, 1)
            End If
        End If
    Next
    End With
    flxOrder.Editable = flexEDNone
End Sub

Private Sub flxOrder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sum As Double
    sum = 0
    For i = 0 To flxOrder.SelectedRows - 1
        If flxOrder.SelectedRow(i) >= 1 Then sum = sum + Val(flxOrder.TextMatrix(flxOrder.SelectedRow(i), flxOrder.Col))
    Next
    Me.Caption = MyCAPTION & "Sum on col " & str(flxOrder.Col) & " = " & sum
    If Button = 2 And Shift = vbCtrlMask Then
        SaveGrid Me.flxOrder
    End If
End Sub

Private Sub flxOrder_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And (Col = ItemID_Col Or Col = ItemCode_Col) Then
            flxOrder.Col = Qty_Col: flxOrder.EditCell
    End If
    
    If KeyAscii = vbKeyReturn And Col = Qty_Col Then
        flxOrder.Col = ItemCode_Col
        If flxOrder.Row < flxOrder.ROWS - 1 Then flxOrder.Row = flxOrder.Row + 1
    End If
    EnumerateGrid
    Calculate
End Sub

Private Sub flxOrder_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim A As Double
    Dim P As Double
    Dim d As Double
    Dim G As Double
    
    If Col = ItemID_Col Then
        If Val(flxOrder.TextMatrix(Row, Col)) = 0 Then
            frmSearchItem.GetLink Me
            ItemID = Val(flxOrder.TextMatrix(Row, Col))
            PopulateGridRow ItemID
        Else
            ItemID = Val(flxOrder.TextMatrix(Row, Col))
            PopulateGridRow ItemID
        End If
    End If
    
    If Col = ItemCode_Col Then
        CN(5).dbOpen "SELECT * from " & ItemSelectView & " Where ItemCode=" & QT(flxOrder.TextMatrix(Row, Col)) & " ORDER BY ItemID DESC"
        msgUITS CN(5).recs.RecordCount & " rows found."
        If CN(5).recs.RecordCount >= 1 Then
            ItemID = CN(5).recs!ItemID
            PopulateGridRow ItemID
        End If
    End If
    
    'This feature needs revisit since we have included multiple discount
    'If AutoCalculate = True Then 'REVERSE DISCOUNT CALCULATIONS
    '    If flxOrder.Col = Amount_col Then
    '        Gross = flxOrder.TextMatrix(Row, Gross_col): Net = flxOrder.TextMatrix(Row, Amount_col)
    '        flxOrder.TextMatrix(Row, Disc_Col) = (1 - (Net / Gross)) * 100
    '    End If
    'End If
    ' D = (1 + G / 100 - A / P) / (0.01 + G * 0.0001)
    
    If flxOrder.Col = Amount_Col Then
        A = Val(flxOrder.TextMatrix(Row, Amount_Col))
        P = Val(flxOrder.TextMatrix(Row, Gross_Col))
        G = Val(flxOrder.TextMatrix(Row, GST_Col))
        d = (1 + G / 100 - A / P) / (0.01 + G * 0.0001)
        flxOrder.TextMatrix(Row, aDisc_Col) = Val(d)
    End If
   
    If flxOrder.Col = aDiscAmt_Col Then
        aa = Val(flxOrder.TextMatrix(Row, aDiscAmt_Col))
        G = Val(flxOrder.TextMatrix(Row, Gross_Col))
        d = (aa / G * 100)
        flxOrder.TextMatrix(Row, aDisc_Col) = Val(d)
    End If
    
    
    
    
    If flxOrder.Col = Stock_Col Then   'STOCK UPDATION
        If MsgBox("Confirm initial stock updation?", vbYesNo + vbQuestion) = vbYes Then
            CN(6).dbOpen "appproc_Updateitems_InitStock " & flxOrder.TextMatrix(flxOrder.Row, ItemID_Col) & ", " & flxOrder.TextMatrix(flxOrder.Row, Stock_Col)
        End If
    End If
    EnumerateGrid
    Calculate
End Sub

Private Sub flxOrder_AfterSort(ByVal Col As Long, Order As Integer)
    EnumerateGrid
End Sub

Private Sub flxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim Cpy As Boolean, Pst As Boolean
    ' copy: ctrl-C, ctrl-X, ctrl-ins
    If KeyCode = vbKeyC And Shift = 2 Then Cpy = True

    ' paste: ctrl-V, shift-ins
    If KeyCode = vbKeyV And Shift = 2 Then Pst = True
    
    ' do it
    If Cpy Then
        Clipboard.Clear
        Clipboard.SetText flxOrder.TextMatrix(flxOrder.Row, flxOrder.Col)
    ElseIf Pst Then
        For i = 0 To flxOrder.SelectedRows - 1
            If flxOrder.Col = aDisc_Col Then flxOrder.TextMatrix(flxOrder.SelectedRow(i), flxOrder.Col) = Clipboard.GetText
        Next
    End If
    
    If KeyCode = vbKeyI And Shift = vbCtrlMask And (flxOrder.Row > 1 And flxOrder.Row - 1 > 0) Then
        flxOrder.RowPosition(flxOrder.Row) = flxOrder.Row - 1
        flxOrder.Row = flxOrder.Row - 1
    End If
    If KeyCode = vbKeyJ And Shift = vbCtrlMask And (flxOrder.Row < flxOrder.ROWS - 1 And flxOrder.Row + 1 <= flxOrder.ROWS - 1) Then
        flxOrder.RowPosition(flxOrder.Row) = flxOrder.Row + 1
        flxOrder.Row = flxOrder.Row + 1
    End If
    
    If KeyCode = vbKeyI And Shift = vbCtrlMask + vbShiftMask Then
       flxOrder.AddItem "0", flxOrder.Row    'New Row
       flxOrder.Row = flxOrder.Row
       flxOrder.ShowCell flxOrder.Row, 0
    End If
    If KeyCode = vbKeyJ And Shift = vbCtrlMask + vbShiftMask Then
       flxOrder.AddItem "0", flxOrder.Row + 1   'New Row
       flxOrder.Row = flxOrder.Row + 1
       flxOrder.ShowCell flxOrder.Row, 0
    End If
    
    If KeyCode = vbKey9 And Shift = vbCtrlMask Then
        frmShow.Init "SELECT ItemID, PRICE, ItemCode, ItemName, MakerAuthor, ProducerNAME FROM Items WHERE ItemCode=" & QT(flxOrder.TextMatrix(flxOrder.Row, ItemCode_Col)) & " ORDER BY ItemID ASC"
        If Val(sArray(0)) <> 0 Then
            flxOrder.TextMatrix(flxOrder.Row, ItemID_Col) = sArray(0)
            flxOrder_AfterEdit flxOrder.Row, ItemID_Col
        End If
    End If
    
    EnumerateGrid
End Sub

Private Sub cmdSelectitem_Click()
    frmSelectItem.GetLink Me
    EnumerateGrid
    Calculate
End Sub

Private Sub cmdQuickItem_Click()
    If flxOrder.Row <> -1 Then
        frmQuickItem.GetLink Me, Val(flxOrder.TextMatrix(flxOrder.Row, ItemID_Col))
    Else
        frmQuickItem.GetLink Me
    End If
End Sub

Private Sub cmdDeleteitem_Click()
    For deleterow = 0 To flxOrder.SelectedRows - 1
        If flxOrder.SelectedRow(i) >= 1 Then flxOrder.RemoveItem flxOrder.SelectedRow(i)
    Next
    EnumerateGrid
    Calculate
End Sub

Public Sub cmdSaveGrid_Click()
    On Error Resume Next
    If MsgBox("Do you wish to save the grid?", vbYesNo + vbQuestion) = vbYes Then
        mdiOne.CDlg.FileName = Format(txtDBRef.Text, Me.MemoFormat) & "-" & txtID.Text & "-" & Format(Now, "DD-MMM-YY HHMM")
        mdiOne.CDlg.Filter = CompanyName & " Excel Report |*.xls"
        mdiOne.CDlg.ShowSave
        If mdiOne.CDlg.CancelError = False Then flxOrder.SaveGrid mdiOne.CDlg.FileName, flexFileExcel
    End If
End Sub

Private Sub cmdLoadGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If MsgBox("Do you wish to load the grid?", vbYesNo + vbQuestion) = vbYes Then
            mdiOne.CDlg.FileName = ""
            mdiOne.CDlg.Filter = Format(txtDBRef.Text, Me.MemoFormat) & " Excel Report |*.xls"
            mdiOne.CDlg.ShowOpen
            If mdiOne.CDlg.CancelError = False Then flxOrder.LoadGrid mdiOne.CDlg.FileName, flexFileExcel
        End If
    End If
    
    If Button = 2 Then
        frmShowLoadMain.Init
        FillFormTextFromMain
    End If

End Sub

Private Sub cmdLoadGrid_Click()
    On Error Resume Next
End Sub

Private Sub cmdExtraDiscountCreditNote_Click()
    On Error Resume Next
    
    
    CN(20).dbOpen "Select NetAmount from " & Main & " where DBRef=" & Val(txtDBRef.Text)
    
    EDPreviousTotal = CN(20).recs!NetAmount
    EDCurrentTotal = Val(txtNetAmount.Text)
    ExtraDiscount = EDPreviousTotal - EDCurrentTotal
    CNSerialNumber = 0
    EDString = EDPreviousTotal & " - " & EDCurrentTotal & " = " & ExtraDiscount
    If MsgBox("Are you sure to update?" & EDString, vbYesNo) = vbYes Then
        CN(21).dbOpen "ADDNEW_VOUCHERS JRNL": Set CN(21).recs = CN(21).recs.NextRecordset
        CNSerialNumber = CN(21).recs!MaxSerial
            
        EDSql = "UPDATE_VOUCHERS " & Val(CNSerialNumber) & ", " & QT(Format(Now, "dd-MMM-yy HH:MM")) _
        & ", " & QT(UCase("N011")) & ", " & QT("DISCOUNTS AND REBATES") & ", " & QT("X") _
        & ", " & QT(UCase(txtID.Text)) & ", " & QT(txtName.Text) & ", " & QT(txtCity.Text) _
        & ", " & Val(ExtraDiscount) & ", " & QT("Credit Note towards Sale#" & txtDBRef.Text) & ", " & QT("JRNL")

        CN(22).dbOpen EDSql
        
        cJr.VoucherToJournalEntries Val(CNSerialNumber), "\V\R0"
        
        MsgBox "Created credit note number " & CNSerialNumber, vbOKOnly
    End If
End Sub

Private Sub CommandButton1_Click()
    APPLYPREVIOUSDISCOUNT
End Sub

Private Sub CommandButton2_Click()
    APPLYCOUNTERDISCOUNT
End Sub

Private Sub chkReturnCalculator_Click()
    ReturnAssist
End Sub

Private Sub cmdCalculate_Click()
    Dim aDiscStr, bDiscStr, cDiscStr As String
    Dim aDisc, bDisc, cDisc As Double
    Dim Gross, aDiscAmt, bDiscAmt, cDiscAmt, CGSTAmt, SGSTAmt, IGSTAmt, CessAmt, Amount As Currency
    Dim Gross_Total, aDiscAmt_Total, bDiscAmt_Total, cDiscAmt_Total, GSTAmt_Total, CessAmt_Total, Amount_Total, NetAmount, RoundOff As Currency
    Dim TempCount, AddItemCount, LessItemCount, CGST, SGST, IGST, Cess As Double
    
    itemCount = 0
    txtToPayMode_Click
    Gross_Total = aDiscAmt_Total = bDiscAmt_Total = cDiscAmt_Total = GSTAmt_Total = CessAmt_Total = Amount_Total = NetAmount = RoundOff = 0
    
    For R = 1 To flxOrder.ROWS - 1
        Gross = Val(flxOrder.TextMatrix(R, Qty_Col)) * Val(flxOrder.TextMatrix(R, SRP_Col))
        
        aDiscStr = flxOrder.TextMatrix(R, aDisc_Col): aDisc = CircularDiscount(aDiscStr)
        aDiscAmt = Gross * (aDisc / 100): aDiscAmt_Total = aDiscAmt_Total + aDiscAmt
        
        bDiscStr = flxOrder.TextMatrix(R, bDisc_Col): bDisc = CircularDiscount(bDiscStr)
        bDiscAmt = Gross * (bDisc / 100): bDiscAmt_Total = bDiscAmt_Total + bDiscAmt
        
        cDiscStr = flxOrder.TextMatrix(R, cDisc_Col): cDisc = CircularDiscount(cDiscStr)
        cDiscAmt = Gross * (cDisc / 100): cDiscAmt_Total = cDiscAmt_Total + cDiscAmt
        
        Amount = Gross - aDiscAmt - bDiscAmt - cDiscAmt
        
        GST = Val(flxOrder.TextMatrix(R, GST_Col))
        Cess = Val(flxOrder.TextMatrix(R, Cess_Col))
        
        GSTAmt = Amount * (GST / 100)
        CessAmt = GSTAmt * Cess / 100
        
        Amount = Amount + GSTAmt + CessAmt
        
        
        GSTAmt_Total = GSTAmt_Total + GSTAmt
        CessAmt_Total = CessAmt_Total + CessAmt
        Amount_Total = Amount_Total + Amount
                
        flxOrder.TextMatrix(R, Gross_Col) = Format(Gross, "###0.00")
        flxOrder.TextMatrix(R, aDiscAmt_Col) = Format(aDiscAmt, "###0.00")
        flxOrder.TextMatrix(R, bDiscAmt_Col) = Format(bDiscAmt, "###0.00")
        flxOrder.TextMatrix(R, cDiscAmt_Col) = Format(cDiscAmt, "###0.00")
        flxOrder.TextMatrix(R, GSTAmt_Col) = Format(GSTAmt, "###0.00")
        flxOrder.TextMatrix(R, CessAmt_Col) = Format(CessAmt, "###0.00")
        flxOrder.TextMatrix(R, Amount_Col) = Format(Amount, "###0.00")
        
        
        TempCount = Val(flxOrder.TextMatrix(R, Qty_Col)) + Val(flxOrder.TextMatrix(R, Free_Col))
        If TempCount >= 0 Then
            AddItemCount = AddItemCount + TempCount
        Else
            LessItemCount = LessItemCount + TempCount
        End If
    Next
    
    txtKAmount.Text = RamjeeRound(Val(txtBundleCount.Text) * Val(txtKRate.Text))
    
    NetAmount = Amount_Total
    NetAmount = NetAmount - (NetAmount * Val(txtSplDisc.Text) / 100)
    NetAmount = NetAmount - (NetAmount * Val(txtBulkDisc.Text) / 100)
    
    NetAmount = NetAmount + Val(txtAddFreight.Text) + Val(txtAddMisc.Text)
    NetAmount = NetAmount - Val(txtLessFreight.Text) - Val(txtLessMisc.Text)
    NetAmount = NetAmount + Val(txtKAmount) + Val(txtPostage.Text)
    
    
    RoundOff = Round(NetAmount) - NetAmount
    NetAmount = Round(NetAmount)
    
    DiscAmt_Total = aDiscAmt_Total + bDiscAmt_Total + cDiscAmt_Total ' adding discount totals to a variable
    txtDiscountAmt.Text = Format(DiscAmt_Total, "###0.00")
    txtItemAmount.Text = Format(Amount_Total, "###0.00")
    
    txtNetGSTAmt.Text = Format(GSTAmt_Total, "###0.00")
    
    ' CALCULATE GST distribution to State/Center
    If Trim(CompanyState) = Trim(txtState.Text) Then
        CGSTAmt_Total = GSTAmt_Total / 2: SGSTAmt_Total = GSTAmt_Total / 2: IGSTAmt_Total = 0
    Else
        CGSTAmt_Total = 0: SGSTAmt_Total = 0: IGSTAmt_Total = GSTAmt_Total
    End If
    
    txtNetCGSTAmt.Text = Format(CGSTAmt_Total, "###0.00")
    txtNetSGSTAmt.Text = Format(SGSTAmt_Total, "###0.00")
    txtNetIGSTAmt.Text = Format(IGSTAmt_Total, "###0.00")
    
    txtNetCessAmt.Text = Format(CessAmt_Total, "###0.00")

    txtNetAmount.Text = Format(NetAmount, "###0.00")
    txtRoundOff.Text = Format(RoundOff, "###0.00")
    txtItemCount.Text = str(AddItemCount) & IIf(LessItemCount < 0, str(LessItemCount), "")
End Sub

Private Sub cmdDeleteBill_Click()
    CN(12).dbOpen "Select Status from " & Main & " where dbref=" & DBRef, 0
    If UCase(CN(12).recs!Status) = "CHALLAN" Then
 '   If (UCase(CN(12).recs!Status) = "CASH" Or UCase(CN(12).recs!Status) = "CREDIT") And UserRights = 0 Then
        CN(13).dbOpen "INSERT INTO CTMAIN SELECT * FROM " & Main & " WHERE DBREF=" & DBRef, 0
        CN(13).dbOpen "INSERT INTO CHANGETRACE SELECT * FROM " & Support & " WHERE DBREFX=" & DBRef, 0
        
        If MsgBox("Confirm deletion?", vbYesNo + vbQuestion) = vbYes Then
            'Delete Memo and its related journal entries; but leaves the order intact.
            CN(8).dbOpen "DELETE JOURNAL WHERE MemoRef=" & QT(Format(Val(txtDBRef.Text), MemoFormat)), 1
            MsgBox "Journal reference " & Format(Val(txtDBRef.Text), MemoFormat) & " deleted!", vbOKOnly + vbCritical
            CN(8).dbOpen "UPDATE " & Main & " SET STATUS=" & QT("ORDER") & " WHERE DBRef=" & Val(txtDBRef.Text), 1
            CN(21).dbOpen "UPDATE " & Main & " SET UserNo = " & QT(Left("XX " & Format(Now, "DDMMM hhmm") & " " & GUID & "|" & txtUserNo.Text, 196)) & " WHERE DBREF=" & Val(txtDBRef.Text), 1
            txtDBRef_LostFocus
        End If
    End If
End Sub

Private Sub cmdCancelBill_Click()
    If MsgBox("Confirm bill cancellation?", vbYesNo + vbQuestion) = vbYes Then
        'Delete Memo and its related journal entries; but leaves the order intact.
        CN(8).dbOpen "DELETE JOURNAL WHERE MemoRef=" & QT(Format(Val(txtDBRef.Text), MemoFormat)), 1
        MsgBox "Journal reference " & Format(Val(txtDBRef.Text), MemoFormat) & " deleted!", vbOKOnly + vbCritical
        CN(8).dbOpen "UPDATE " & Main & " SET STATUS=" & QT("CANCELLED") & " WHERE DBRef=" & Val(txtDBRef.Text), 1
        CN(21).dbOpen "UPDATE " & Main & " SET UserNo = " & QT(Left("XX " & Format(Now, "DDMMM hhmm") & " " & GUID & "|" & txtUserNo.Text, 196)) & " WHERE DBREF=" & Val(txtDBRef.Text), 1
        txtDBRef_LostFocus
    End If
End Sub

Public Sub cmdSaveOrder_Click()
    Call flxOrder_LeaveCell
    
    DeleteString = "Delete " & Support & " where DBRefX=" & Val(txtDBRef.Text)
    SaveString = "Select * from " & Support & " where DBRefX=" & Val(txtDBRef.Text) & " order by Serial"
    
    bs = GetOrderStatus(Main, Val(txtDBRef.Text))
    If UCase(bs) = "NEW" Or UCase(bs) = "ORDER" Then
        toMain "ORDER"
        toSupport DeleteString, SaveString, flxOrder
        frmTimeOne.Show
        msgUITS Support & " Order Data Saved!"
    Else
        MsgBox "Current status is " & Trim(bs) & vbCrLf & Operation & " operation failed!", vbOKOnly + vbCritical
    End If
    Call txtDBRef_LostFocus
End Sub

Private Sub cmdMemoFinalize_Click(Index As Integer)
    Dim Operation As String
    
    Select Case Index
        Case 0: Operation = "CHALLAN"
        Case 1: Operation = "CASH"
        Case 2: Operation = "CREDIT"
    End Select
    
    bs = GetOrderStatus(Main, Val(txtDBRef.Text))
    If UCase(bs) = "NEW" Or UCase(bs) = "ORDER" Then
        cmdSaveOrder_Click
        bs = GetOrderStatus(Main, Val(txtDBRef.Text))
    End If
    
    If (UCase(bs) = "ORDER") Then
        If MaterialiseOrder(Val(txtDBRef.Text)) Then
            CN(9).dbOpen "Update " & Main & " SET Status=" & QT(Operation) & " WHERE DBRef=" & Val(txtDBRef.Text), 1
            If cJr.PPSSJournalEntries(Support, Val(txtDBRef.Text), MemoFormat) Then
                CN(21).dbOpen "UPDATE " & Main & " SET UserNo = " & QT(Left(Left(Operation, 2) & " " & Format(flxOrder.ROWS - 1, "00") & " " & Format(Now, "DDMMM hhmm") & " " & Left(GUID, 4) & "|" & txtUserNo.Text, 198)) & " WHERE DBREF=" & Val(txtDBRef.Text), 1
                CN(10).dbOpen "appproc_SetInvNo " & Val(txtDBRef.Text) & ", " & QT(Main), 1
                msgUITS Operation & " operation successful!"
            Else
                CN(10).dbOpen "Update " & Main & " SET Status=" & QT("ORDER") & " WHERE DBRef=" & Val(txtDBRef.Text), 1
            End If
        End If
        Call txtDBRef_LostFocus
    Else
        MsgBox "Current status is " & Trim(bs) & vbCrLf & Operation & " operation failed!", vbOKOnly + vbCritical
    End If
End Sub

Private Sub cmdPrintLong_Click()
    If Val(txtDBRef.Text) <> 0 Then
        frmPrintMemoLongFormat.PrintIT Val(txtDBRef.Text), Support, 0
    End If
End Sub

Private Sub cmdPrintSmall_Click()
    If Val(txtDBRef.Text) <> 0 Then
        frmPrintMemoSmallFormat.PrintIT Val(txtDBRef.Text), Support, printFormat
    End If
End Sub

Private Sub cmdPrint_Click()
    If Val(txtDBRef.Text) <> 0 Then
        Select Case Support
            Case "SALE": printFormat = 0
            Case "SALERETURN": printFormat = 0
            Case "PURCHASE": printFormat = 4
            Case "PURCHASERETURN": printFormat = 5
            Case "TOUT": printFormat = 2
            Case "TIN": printFormat = 2
        End Select
        If BILL_PRINT_FORMAT = "FIXED" Then
            frmPrintMemoFixedFormat.PrintIT Val(txtDBRef.Text), Support, printFormat
        Else
            frmPrintMemoFixedFormat.PrintIT Val(txtDBRef.Text), Support, printFormat
        End If
    End If
End Sub

Private Sub cmdQuit_Click()
    For i = 0 To 4
        sSQL(i) = ""
    Next
    Erase sArray
    Unload Me
End Sub

'================================================
' SUB ROUTINES
'================================================

Private Sub FillFormTextFromMain()
    i = 1 ' txtDBRef.Text = sArray(i): i = 1
    txtDBDate.Value = sArray(i): i = i + 1
    txtStatus.Text = sArray(i): i = i + 1
    txtOrderRef.Text = sArray(i): i = i + 1
    txtOrderDate.Value = sArray(i): i = i + 1
    txtInvRef.Text = sArray(i): i = i + 1
    txtInvDate.Value = sArray(i): i = i + 1
    
    txtID.Text = sArray(i): i = i + 1           ' Customer
    txtName.Text = sArray(i): i = i + 1
    txtAddress.Text = sArray(i): i = i + 1
    txtCity.Text = sArray(i): i = i + 1
    txtState.Text = sArray(i): i = i + 1
    txtPhones.Text = sArray(i): i = i + 1
    txtEmail.Text = sArray(i): i = i + 1
    txtShipAddress.Text = sArray(i): i = i + 1
    txtShipState.Text = sArray(i): i = i + 1
    
    txtTID.Text = sArray(i): i = i + 1          ' Transporter
    txtTName.Text = sArray(i): i = i + 1
    txtTaddress.Text = sArray(i): i = i + 1
    txtTCity.Text = sArray(i): i = i + 1
    txtTPhones.Text = sArray(i): i = i + 1
    txtTEmail.Text = sArray(i): i = i + 1
    txtTerminal.Text = sArray(i): i = i + 1
    
    txtKID.Text = sArray(i): i = i + 1          ' Karter
    txtKName.Text = sArray(i): i = i + 1
    txtKRate.Text = sArray(i): i = i + 1
    txtKAmount = sArray(i): i = i + 1
    
    txtPID.Text = sArray(i): i = i + 1          ' Postage
    txtPName.Text = sArray(i): i = i + 1
    txtPostage.Text = sArray(i): i = i + 1
    
    txtSDID.Text = sArray(i): i = i + 1
    txtSDName.Text = sArray(i): i = i + 1
    txtSDCommission.Text = sArray(i): i = i + 1
    
    txtGRNo.Text = sArray(i): i = i + 1
    txtGRDate.Value = sArray(i): i = i + 1
    txtGRMode.ListIndex = Val(sArray(i)): i = i + 1
    txtGRAmount.Text = sArray(i): i = i + 1
    txtToPayMode.ListIndex = sArray(i): i = i + 1
    txtBundleCount.Text = sArray(i): i = i + 1
    txtBundleWeight.Text = sArray(i): i = i + 1
    txtItemCount.Text = sArray(i): i = i + 1
    
    txtDiscountAmt.Text = Val(sArray(i)): i = i + 1
    txtItemAmount.Text = sArray(i): i = i + 1
    txtSplDisc.Text = sArray(i): i = i + 1
    txtBulkDisc.Text = sArray(i): i = i + 1
    txtAddMisc.Text = sArray(i): i = i + 1
    txtLessMisc.Text = sArray(i): i = i + 1
    txtAddFreight.Text = sArray(i): i = i + 1
    txtLessFreight.Text = sArray(i): i = i + 1
    txtNetGSTAmt.Text = sArray(i): i = i + 1
    txtNetCGSTAmt.Text = sArray(i): i = i + 1
    txtNetSGSTAmt.Text = sArray(i): i = i + 1
    txtNetIGSTAmt.Text = sArray(i): i = i + 1
    txtNetCessAmt.Text = sArray(i): i = i + 1
    txtRoundOff.Text = sArray(i): i = i + 1
    txtNetAmount.Text = sArray(i): i = i + 1
    txtUserNo.Text = sArray(i): i = i + 1
    txtComments.Text = sArray(i): i = i + 1
End Sub

Public Sub toMain(Status As String)
    CN(18).dbOpen "Select * from " & Main & " where DBRef=" & DBRef, 1
'       cn(18).recs!DBRef = txtDBRef.Text
        CN(18).recs!DBDate = txtDBDate.Value
        CN(18).recs!Status = Status
        CN(18).recs!OrderRef = txtOrderRef.Text
        CN(18).recs!OrderDate = txtOrderDate.Value
        CN(18).recs!InvRef = txtInvRef.Text
        CN(18).recs!INVDate = txtInvDate.Value
        CN(18).recs!id = txtID.Text
        CN(18).recs!Name = txtName.Text
        CN(18).recs!Address = txtAddress.Text
        CN(18).recs!City = txtCity.Text
        CN(18).recs!State = txtState.Text
        CN(18).recs!Phones = txtPhones.Text
        CN(18).recs!Email = txtEmail.Text
        CN(18).recs!ShipAddress = txtShipAddress.Text
        CN(18).recs!ShipState = txtShipState.Text
        
        CN(18).recs!TID = txtTID.Text
        CN(18).recs!TName = txtTName.Text
        CN(18).recs!Taddress = txtTaddress.Text
        CN(18).recs!TCity = txtTCity.Text
        CN(18).recs!TPhones = txtTPhones.Text
        CN(18).recs!TEmail = txtTEmail.Text
        CN(18).recs!Terminal = txtTerminal.Text
        
        CN(18).recs!KID = txtKID.Text
        CN(18).recs!KName = txtKName.Text
        CN(18).recs!KRate = Val(txtKRate.Text)
        CN(18).recs!KAmount = Val(txtKAmount.Text)
        
        CN(18).recs!PID = txtPID.Text
        CN(18).recs!PName = txtPName.Text
        CN(18).recs!Postage = Val(txtPostage.Text)
        
        CN(18).recs!SDID = txtSDID.Text
        CN(18).recs!SDName = txtSDName.Text
        CN(18).recs!SDCommission = Val(txtSDCommission.Text)
        
        CN(18).recs!GRNo = txtGRNo.Text
        CN(18).recs!GRDate = txtGRDate.Value
        CN(18).recs!GRMode = txtGRMode.ListIndex
        CN(18).recs!GRAmount = Val(txtGRAmount.Text)
        CN(18).recs!ToPayMode = txtToPayMode.ListIndex
        CN(18).recs!BundleCount = Val(txtBundleCount.Text)
        CN(18).recs!BundleWeight = Val(txtBundleWeight.Text)
        
        CN(18).recs!itemCount = Val(txtItemCount.Text)
        CN(18).recs!DiscAmt = Val(txtDiscountAmt.Text)
        CN(18).recs!itemAmount = Val(txtItemAmount.Text)
        CN(18).recs!SplDisc = Val(txtSplDisc.Text)
        CN(18).recs!BulkDisc = Val(txtBulkDisc.Text)
        CN(18).recs!AddMisc = Val(txtAddMisc.Text)
        CN(18).recs!LessMisc = Val(txtLessMisc.Text)
        CN(18).recs!AddFreight = Val(txtAddFreight.Text)
        CN(18).recs!LessFreight = Val(txtLessFreight.Text)
        
        CN(18).recs!NetGSTAmt = Val(txtNetGSTAmt.Text)
        CN(18).recs!NetCGSTAmt = Val(txtNetCGSTAmt.Text)
        CN(18).recs!NetSGSTAmt = Val(txtNetSGSTAmt.Text)
        CN(18).recs!NetIGSTAmt = Val(txtNetIGSTAmt.Text)
        CN(18).recs!NetCessAmt = Val(txtNetCessAmt.Text)
        
        CN(18).recs!RoundOff = Val(txtRoundOff.Text)
        CN(18).recs!NetAmount = Val(txtNetAmount.Text)
        CN(18).recs!UserNo = Left(Left(Status, 2) & " " & Format(flxOrder.ROWS - 1, "00") & " " & Format(Now, "DDMMM hhmm") & " " & Left(GUID, 4) & "|" & CN(18).recs!UserNo, 198)
        CN(18).recs!Comments = txtComments.Text
    CN(18).recs.Update
End Sub

Public Sub toSupport(ByVal DeleteStr As String, ByVal SaveStr As String, MyFlex As Control)
    Dim i As Integer
    CN(19).dbOpen DeleteStr, 1
    CN(19).dbOpen SaveStr, 1
    With MyFlex
        For i = 1 To .ROWS - 1
            If ValidateItemID(Val(.TextMatrix(i, 2))) = True Then
                CN(19).recs.AddNew
                CN(19).recs!Serial = i              '.TextMatrix(i, 0)
                CN(19).recs!DBRefX = Val(txtDBRef.Text)
                CN(19).recs!ItemID = Val(.TextMatrix(i, ItemID_Col))
                CN(19).recs!ItemCode = Trim(.TextMatrix(i, ItemCode_Col))
                CN(19).recs!HSNCode = Trim(.TextMatrix(i, HSNCode_Col))
                CN(19).recs!ItemName = Left(.TextMatrix(i, ItemName_Col), 200)
                CN(19).recs!MakerAuthor = Left(.TextMatrix(i, MakerAuthor_Col), 200)
                CN(19).recs!ProducerID = Left(.TextMatrix(i, ProducerID_Col), 10)
                CN(19).recs!ProducerName = Left(.TextMatrix(i, ProducerName_Col), 200)
                CN(19).recs!Version = Left(.TextMatrix(i, Version_Col), 10)
                CN(19).recs!MfgDate = Left(.TextMatrix(i, MfgDate_Col), 10)
                CN(19).recs!ExpDate = Left(.TextMatrix(i, ExpDate_Col), 10)
                CN(19).recs!Packing = Left(.TextMatrix(i, Packing_Col), 10)
                CN(19).recs!Unit = Val(.TextMatrix(i, Unit_Col))
                
                CN(19).recs!Qty = Val(.TextMatrix(i, Qty_Col))
                CN(19).recs!Free = Val(.TextMatrix(i, Free_Col))
                CN(19).recs!MRP = Val(.TextMatrix(i, MRP_Col))
                CN(19).recs!SRP = Val(.TextMatrix(i, SRP_Col))
                CN(19).recs!Gross = Val(.TextMatrix(i, Gross_Col))
                
                CN(19).recs!aDisc = Trim(.TextMatrix(i, aDisc_Col))
                CN(19).recs!aDiscAmt = Trim(.TextMatrix(i, aDiscAmt_Col))
                CN(19).recs!bDisc = Trim(.TextMatrix(i, bDisc_Col))
                CN(19).recs!bDiscAmt = Trim(.TextMatrix(i, bDiscAmt_Col))
                CN(19).recs!cDisc = Trim(.TextMatrix(i, cDisc_Col))
                CN(19).recs!cDiscAmt = Trim(.TextMatrix(i, cDiscAmt_Col))
                                
                CN(19).recs!GST = Trim(.TextMatrix(i, GST_Col))
                CN(19).recs!GSTAmt = Val(.TextMatrix(i, GSTAmt_Col))
                CN(19).recs!Cess = Trim(.TextMatrix(i, Cess_Col))
                CN(19).recs!CessAmt = Val(.TextMatrix(i, CessAmt_Col))
                
                CN(19).recs!Amount = Val(.TextMatrix(i, Amount_Col))
                CN(19).recs!Stock = Val(.TextMatrix(i, Stock_Col))
            End If
        Next
    End With
    CN(19).recs.UpdateBatch
    'CN(19).recs.Update
End Sub

Public Function MaterialiseOrder(ByVal DBRef As Long) As Boolean
    Dim B As Boolean
    Dim msg As String
    B = True
   
    If StockValidationRequired Then
        CN(11).dbOpen "SELECT A.Serial, A.ItemID, (A.QTY + A.FREE) AS REQD, B.AVLBL, (B.AVLBL - (A.QTY + A.FREE)) AS DIFF from " & Support & " A, appview_Stock B where A.ItemID = B.ItemID AND A.DBRefX=" & DBRef & " ORDER BY SERIAL", 1
        If CN(11).recs.RecordCount <> 0 Then CN(11).recs.MoveFirst
        Do Until CN(11).recs.EOF
            If Val(CN(11).recs!Diff) < 0 Then
                msg = msg & vbCrLf & CN(11).recs!Serial & ". ItemID: " & CN(11).recs!ItemID & " is short on stock. Required=" & CN(11).recs!Reqd & " & Available=" & CN(11).recs!AVLBL
                B = False
            End If
            CN(11).recs.MoveNext
        Loop
        If B = False Then
            MsgBox msg, vbOKOnly + vbCritical
        Else
            If MsgBox("Order passed! Are you sure to save it.", vbYesNo + vbQuestion) = vbNo Then
                B = False
            End If
        End If
        MaterialiseOrder = B
    Else
        MaterialiseOrder = True
    End If
End Function
'Updated on 20 Aug 2008 with appproc_ReturnDiscount
Private Sub ApplyDiscounts(ByVal opt As String)
    On Error Resume Next
    If Discs = True Then
        For i = 1 To flxOrder.ROWS - 1
            'PARTY
            Party = txtID.Text
            ItemID = Val(flxOrder.TextMatrix(i, ItemID_Col))
            If ItemID <> 0 Then
                procsql = "appproc_ReturnDiscount " & QT(Party) & ", " & ItemID & ", " & QT(opt) & ", " & QT(Support)
                CN(12).dbOpen procsql, 1
                If Not CN(12).recs.EOF Then flxOrder.TextMatrix(i, aDisc_Col) = CN(12).recs!aDisc
            End If
        Next
        msgUITS "discounts applied"
        Calculate
    End If
End Sub

Private Sub SaveDiscounts()
    On Error Resume Next
    If Discs = True Then
        For i = 1 To flxOrder.ROWS - 1
            'PARTY
            Party = txtID.Text
            ItemID = Val(flxOrder.TextMatrix(i, ItemID_Col))
            If ItemID <> 0 Then CN(12).dbOpen "appproc_SaveDiscount " & QT(Party) & ", " & ItemID & ", " & QT(flxOrder.TextMatrix(i, aDisc_Col)), 1
        Next
        msgUITS "Discounts saved."
    End If
End Sub

Private Sub SavePersonalSettings()
    If HasAccount(txtID.Text) = True Then
        If MsgBox("Do you wish to save user settings with ID=" & QT(txtID.Text) & " TID=" & QT(txtTID.Text) & " KID=" & QT(txtKID.Text) & " PID=" & QT(txtPID.Text) & " SDID=" & QT(txtSDID.Text), vbYesNo) = vbYes Then
            CN(20).dbOpen "UPDATE PERSONAL SET TID=" & QT(txtTID.Text) & ", KID=" & QT(txtKID.Text) & ", PID=" & QT(txtPID.Text) & ", SDID=" & QT(txtSDID.Text) & ", DISCOUNT=" & QT(txtSplDisc.Text) & " WHERE ID=" & QT(txtID.Text), 1
            msgUITS "USER SETTINGS SAVED!"
        End If
    Else
        msgUITS txtID.Text & " has no account, cannot save user settings"
    End If
End Sub

Private Sub MergeRows()
ResumeMerge:
    For i = 1 To flxOrder.ROWS - 1
        For j = i + 1 To flxOrder.ROWS - 1
            If flxOrder.TextMatrix(i, ItemID_Col) = flxOrder.TextMatrix(j, ItemID_Col) Then
                flxOrder.TextMatrix(i, Qty_Col) = Val(flxOrder.TextMatrix(i, Qty_Col)) + Val(flxOrder.TextMatrix(j, Qty_Col))
                flxOrder.TextMatrix(i, Qty_Col + 1) = Val(flxOrder.TextMatrix(i, Qty_Col + 1)) + Val(flxOrder.TextMatrix(j, Qty_Col + 1))
                flxOrder.RemoveItem (j)
                msgUITS "redundancy found at " & str(j)
                GoTo ResumeMerge
            End If
        Next
    Next
End Sub

Private Sub PictureBoxStatus(S As Boolean)
    pboxB.Enabled = S
    pboxC.Enabled = S
    pboxD.Enabled = S
    pboxE.Enabled = S
    pboxF.Enabled = S
    pboxG.Enabled = S
    pboxH.Enabled = S
    pboxI.Enabled = S
    pboxJ.Enabled = S
    pboxK.Enabled = S
End Sub

Private Sub EnumerateGrid()
    For i = 1 To flxOrder.ROWS - 1
        flxOrder.TextMatrix(i, 0) = i
    Next
End Sub

Public Sub LinkOpen()
    Me.SetFocus
    txtDBRef.Text = XREF: txtDBRef_LostFocus
End Sub

Private Function ValidateItemID(ByVal ItemID As Long) As Boolean
    CN(16).dbOpen "SELECT ItemID FROM items WHERE ItemID=" & ItemID, 1
    If CN(16).recs.RecordCount = 1 Then
        ValidateItemID = True
    Else
        ValidateItemID = False
    End If
End Function

Public Function PopulateGridRow(ByVal ItemID As Long, Optional Qty As Long) As Boolean
    CN(15).dbOpen "SELECT * FROM " & ItemSelectView & " WHERE ItemID=" & ItemID
    If CN(15).recs.RecordCount >= 1 Then
        For i = 0 To CN(15).recs.FIELDS.Count - 1
            If i <> Qty_Col Then flxOrder.TextMatrix(flxOrder.Row, i) = CN(15).recs.FIELDS(i)
        Next
        If Qty <> 0 Then flxOrder.TextMatrix(flxOrder.Row, Qty_Col) = Qty
        If Val(flxOrder.TextMatrix(flxOrder.Row, Qty_Col)) = 0 Then flxOrder.TextMatrix(flxOrder.Row, Qty_Col) = 1
    Else
        For i = 0 To flxOrder.COLS - 1
            flxOrder.TextMatrix(flxOrder.Row, i) = ""
        Next
    End If
End Function

Private Sub AddBlankRows()
    For i = 0 To 10
       flxOrder.AddItem "0"      'New Row
    Next
End Sub

Private Sub ColShowHide(ByVal opt As Integer)
    If opt = 0 Then     'Show
        For i = 0 To flxOrder.COLS - 1
            flxOrder.ColHidden(i) = False
        Next
    End If
    If opt = 1 Then     'Hide
        For Each i In Split(mdiOne.sckGo.GReadINI(HiddenCols), ",")
            flxOrder.ColHidden(i) = True
        Next
    End If
End Sub

Public Sub SelectItem(ByVal ItemID As Long, ByVal Qty As Long)
    flxOrder.AddItem "0": Row = flxOrder.ROWS - 1
'   flxOrder.SaveGrid "C:\" & Support & ".xls", flexFileExcel
    CN(17).dbOpen "SELECT * from " & ItemSelectView & " Where ItemID=" & ItemID, 1
    If CN(17).recs.RecordCount >= 1 Then
        For i = 0 To CN(17).recs.FIELDS.Count - 1
            If i <> Qty_Col Then
                flxOrder.TextMatrix(Row, i) = CN(17).recs.FIELDS(i)
            Else
                flxOrder.TextMatrix(Row, i) = Val(flxOrder.TextMatrix(Row, i))
            End If
        Next
        flxOrder.TextMatrix(Row, Qty_Col) = Qty
    Else
        For i = 0 To flxOrder.COLS - 1
            flxOrder.TextMatrix(Row, i) = ""
        Next
    End If
    flxOrder.ShowCell flxOrder.Row, ItemCode_Col
    Calculate
End Sub

Private Sub Calculate()
    If AutoCalculate = True Then cmdCalculate_Click
End Sub

Private Sub txtState_Change()
    Calculate
End Sub

Private Sub txtID_DblClick()
    XID = txtID.Text
    frmLedger.LinkOpen
End Sub

Private Sub txtTID_DblClick()
    XID = txtTID.Text
    frmLedger.LinkOpen
End Sub

Private Sub txtKID_DblClick()
    XID = txtKID.Text
    frmLedger.LinkOpen
End Sub

Private Sub txtSDID_DblClick()
    XID = txtSDID.Text
    frmLedger.LinkOpen
End Sub

Private Sub txtPID_DblClick()
    XID = txtPID.Text
    frmLedger.LinkOpen
End Sub


Private Sub APPLYPREVIOUSDISCOUNT()
    On Error Resume Next
    If Discs = True Then
        For i = 1 To flxOrder.ROWS - 1
            'PARTY
            Party = txtID.Text
            ItemID = Val(flxOrder.TextMatrix(i, ItemID_Col))
            If ItemID <> 0 Then
                If CompanyDivision = "CHAS" Then
                    CN(12).dbOpen "appproc_ReturnCounterDiscount " & QT(Party) & ", " & ItemID & ", " & QT(Support), 1
                Else
                    CN(12).dbOpen "appproc_ReturnCounterDiscount " & QT(Party) & ", " & ItemID & ", " & QT(Support & "_BRANCH"), 1
                End If
                If Not CN(12).recs.EOF Then flxOrder.TextMatrix(i, aDisc_Col) = CN(12).recs!aDisc
            End If
            flxOrder.Row = i: flxOrder.ShowCell i, aDisc_Col
            DoEvents
        Next
        msgUITS "previous discounts applied"
        Calculate
    End If
End Sub

Private Sub APPLYCOUNTERDISCOUNT()
    On Error Resume Next
    If Discs = True Then
        For i = 1 To flxOrder.ROWS - 1
            'PARTY
            Party = txtID.Text
            ItemID = Val(flxOrder.TextMatrix(i, ItemID_Col))
            If ItemID <> 0 Then
                If CompanyDivision = "CHAS" Then
                    CN(12).dbOpen "appproc_ReturnCounterDiscount " & QT(Party) & ", " & ItemID & ", " & QT(CounterSupport), 1
                Else
                    CN(12).dbOpen "appproc_ReturnCounterDiscount " & QT(Party) & ", " & ItemID & ", " & QT(CounterSupport & "_BRANCH"), 1
                End If
                If Not CN(12).recs.EOF Then flxOrder.TextMatrix(i, aDisc_Col) = CN(12).recs!aDisc
            End If
            flxOrder.Row = i: flxOrder.ShowCell i, aDisc_Col
            DoEvents
        Next
        msgUITS "counter discounts applied"
        Calculate
    End If
End Sub

Private Function Traceitems() As String
    On Error GoTo ErrorSub
    For i = 1 To flxOrder.ROWS - 1
        Party = txtID.Text  'Party
        ItemID = Val(flxOrder.TextMatrix(i, ItemID_Col))
        If ItemID <> 0 Then
            If chkIncludeID.Value = 0 Then
                CN(12).dbOpen "appproc_TraceitemsIn " & QT(txtTraceIn.Text) & ", " & ItemID, 1
            Else
                CN(12).dbOpen "appproc_TraceitemsIn " & QT(txtTraceIn.Text) & ", " & ItemID & ", " & QT(Party), 1
            End If
            If Not CN(12).recs.EOF Then
                While Not CN(12).recs.EOF
                    STRTRACE = STRTRACE & vbCrLf & CN(12).recs!Type & " : " & CN(12).recs!Name & " : DBRef=" & CN(12).recs!DBRef & " : " & Trim(flxOrder.TextMatrix(i, Serial_Col)) & ". " & Trim(flxOrder.TextMatrix(i, ItemName_Col)) & ", Inv: " & CN(12).recs!InvRef & " Dt: " & Format(CN(12).recs!INVDate, "dd-mm-yy") & " Qty: " & CN(12).recs!Qty & " Price: " & CN(12).recs!Price
                    CN(12).recs.MoveNext
                Wend
                STRTRACE = STRTRACE & vbCrLf
            End If
        End If
    Next
    Traceitems = STRTRACE
    Exit Function

ErrorSub:
    Traceitems = Error
End Function

Private Sub ReturnAssist()
    If chkReturnCalculator.Value = 1 Then
        ReturnAssistPub = flxOrder.TextMatrix(flxOrder.Row, ProducerID_Col)
        Select Case Support
            Case "SALE", "SALERETURN"
                CN(20).dbOpen "appproc_itemReturnEnhancer " & txtID.Text & ", " & ReturnAssistPub & ", SABIG " & ", " & QT(ReturnSAStartDate) & ", " & QT(ReturnSAEndDate), 0
                XValue = Val(sArray(0))
                CN(21).dbOpen "appproc_itemReturnEnhancer " & txtID.Text & ", " & ReturnAssistPub & ", SRBIG " & ", " & QT(ReturnSRStartDate) & ", " & QT(ReturnSREndDate), 0
                XReturnValue = Val(sArray(0))
            Case "PURCHASE", "PURCHASERETURN"
                CN(20).dbOpen "appproc_itemReturnEnhancer " & txtID.Text & ", " & ReturnAssistPub & ", PUBIG " & ", " & QT(ReturnPUStartDate) & ", " & QT(ReturnPUEndDate), 0
                XValue = Val(sArray(0))
                CN(21).dbOpen "appproc_itemReturnEnhancer " & txtID.Text & ", " & ReturnAssistPub & ", PRBIG " & ", " & QT(ReturnPRStartDate) & ", " & QT(ReturnPREndDate), 0
                XReturnValue = Val(sArray(0))
        End Select
        If XValue <> 0 Then
            XPerc = (XReturnValue / XValue) * 100
            chkReturnCalculator.Caption = "ReturnAssist: [ " + Format(XPerc, "#00.00") + "% ] [" + Trim(str(XValue)) + " - " + Trim(str(XReturnValue)) + "]:" + ReturnAssistPub
        Else
            chkReturnCalculator.Caption = "*** For ReturnAssist help, contact vendor ***"
        End If
    Else
        chkReturnCalculator.Caption = "*** For ReturnAssist help, contact vendor ***"
    End If
End Sub
