VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPurchaseReturn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PurchaseReturns..."
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   Icon            =   "frmPurchaseReturn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      DragMode        =   1  'Automatic
      Height          =   1110
      Left            =   15
      TabIndex        =   92
      Top             =   5925
      Width           =   11685
      Begin VB.PictureBox pboxJ 
         Height          =   900
         Left            =   75
         ScaleHeight     =   840
         ScaleWidth      =   11520
         TabIndex        =   93
         Top             =   150
         Width           =   11580
         Begin MSMask.MaskEdBox txtAddFreight 
            Height          =   270
            Left            =   4320
            TabIndex        =   94
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
            Left            =   7065
            TabIndex        =   95
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
            Left            =   9945
            TabIndex        =   96
            Top             =   285
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Left            =   4320
            TabIndex        =   97
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
            Left            =   7065
            TabIndex        =   98
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
            Left            =   9945
            TabIndex        =   99
            Top             =   570
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Left            =   1455
            TabIndex        =   100
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
            Left            =   1455
            TabIndex        =   101
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
            Left            =   7065
            TabIndex        =   102
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
            Left            =   4320
            TabIndex        =   103
            Top             =   -15
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSubject 
            Height          =   270
            Left            =   1455
            TabIndex        =   104
            Top             =   0
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   476
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "SDCommission:"
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
            Left            =   150
            TabIndex        =   115
            Top             =   300
            Width           =   1320
         End
         Begin VB.Label Label10 
            Caption         =   "Add Postage:"
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
            Left            =   330
            TabIndex        =   114
            Top             =   555
            Width           =   1260
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
            Left            =   8760
            TabIndex        =   113
            Top             =   570
            Width           =   1200
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Caption         =   "+ Misc Amt:"
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
            Left            =   5850
            TabIndex        =   112
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
            Left            =   3120
            TabIndex        =   111
            Top             =   585
            Width           =   1185
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
            Left            =   3090
            TabIndex        =   110
            Top             =   285
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
            Left            =   5850
            TabIndex        =   109
            Top             =   570
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
            Height          =   285
            Left            =   8760
            TabIndex        =   108
            Top             =   300
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
            Left            =   5850
            TabIndex        =   107
            Top             =   0
            Width           =   1200
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
            Left            =   3075
            TabIndex        =   106
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Subject:"
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
            Left            =   120
            TabIndex        =   105
            Top             =   15
            Width           =   1320
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
      Height          =   3990
      Left            =   15
      TabIndex        =   83
      Top             =   1995
      Width           =   11685
      Begin VB.PictureBox pboxH 
         Height          =   330
         Left            =   45
         ScaleHeight     =   270
         ScaleWidth      =   11535
         TabIndex        =   84
         Top             =   3615
         Width           =   11595
         Begin VB.CommandButton cmdSelectItem 
            Caption         =   "S&ELECT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -30
            TabIndex        =   87
            Top             =   0
            Width           =   1305
         End
         Begin VB.CommandButton cmdDeleteItem 
            Caption         =   "&DELETE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1275
            TabIndex        =   86
            Top             =   0
            Width           =   1305
         End
         Begin VB.TextBox txtItemCount 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Enabled         =   0   'False
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
            Left            =   8070
            TabIndex        =   85
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   -15
            Width           =   1005
         End
         Begin MSMask.MaskEdBox txtAmount 
            Height          =   285
            Left            =   9930
            TabIndex        =   88
            Top             =   0
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "###0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   9330
            TabIndex        =   90
            Top             =   15
            Width           =   765
         End
         Begin VB.Label lblItemCount 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ItemCount:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7095
            TabIndex        =   89
            Top             =   0
            Width           =   960
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid flxOrder 
         Height          =   3435
         Left            =   60
         TabIndex        =   91
         Top             =   165
         Width           =   11580
         _cx             =   20426
         _cy             =   6059
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
         BackColor       =   12632256
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632319
         ForeColorSel    =   64
         BackColorBkg    =   12632256
         BackColorAlternate=   12632256
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
         GridLineWidth   =   1
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
         PictureType     =   0
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
   Begin VB.PictureBox pboxI 
      Height          =   360
      Left            =   30
      ScaleHeight     =   300
      ScaleWidth      =   11595
      TabIndex        =   23
      Top             =   7065
      Width           =   11655
      Begin VB.TextBox txtComments 
         Height          =   315
         Left            =   1065
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   10530
      End
      Begin VB.Label Label9 
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   25
         Top             =   30
         Width           =   1695
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
      Height          =   2040
      Left            =   15
      TabIndex        =   9
      Top             =   0
      Width           =   3000
      Begin VB.PictureBox pboxA 
         Height          =   1770
         Left            =   60
         ScaleHeight     =   1710
         ScaleWidth      =   2835
         TabIndex        =   10
         Top             =   195
         Width           =   2895
         Begin VB.TextBox txtInvRef 
            Height          =   345
            Left            =   600
            TabIndex        =   19
            Top             =   705
            Width           =   915
         End
         Begin VB.TextBox txtDBRef 
            Height          =   330
            Left            =   600
            TabIndex        =   15
            Top             =   15
            Width           =   915
         End
         Begin VB.TextBox txtOrderRef 
            Height          =   330
            Left            =   600
            TabIndex        =   14
            Top             =   360
            Width           =   915
         End
         Begin VB.CommandButton cmdNewDBRef 
            Caption         =   "&NEW"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   0
            TabIndex        =   13
            Top             =   1380
            Width           =   1440
         End
         Begin VB.CommandButton cmdSelectDBRef 
            Height          =   345
            Left            =   1425
            Picture         =   "frmPurchaseReturn.frx":4E0E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1380
            Width           =   1425
         End
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
            Height          =   255
            Left            =   -15
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "STATUS"
            Top             =   1125
            Width           =   2865
         End
         Begin MSComCtl2.DTPicker txtOrderDate 
            Height          =   330
            Left            =   1530
            TabIndex        =   16
            Top             =   360
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
            Format          =   22675459
            CurrentDate     =   38023
         End
         Begin MSComCtl2.DTPicker txtDBDate 
            Height          =   330
            Left            =   1530
            TabIndex        =   17
            Top             =   15
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
            Format          =   22675459
            CurrentDate     =   38023
         End
         Begin MSComCtl2.DTPicker txtInvDate 
            Height          =   345
            Left            =   1530
            TabIndex        =   18
            Top             =   705
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
            Format          =   22675459
            CurrentDate     =   38023
         End
         Begin VB.Label Label3 
            Caption         =   "DBRef:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   15
            TabIndex        =   22
            Top             =   75
            Width           =   810
         End
         Begin VB.Label Label35 
            Caption         =   "Order:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   15
            TabIndex        =   21
            Top             =   405
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Inv:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   30
            TabIndex        =   20
            Top             =   765
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox pboxK 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11640
      TabIndex        =   0
      Top             =   7440
      Width           =   11700
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
         Height          =   285
         Left            =   1230
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   30
         Width           =   1230
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
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   1230
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
         Height          =   315
         Left            =   4260
         TabIndex        =   6
         Top             =   15
         Width           =   1230
      End
      Begin VB.CommandButton cmdMemoFinalize 
         Caption         =   "Challan"
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
         Index           =   0
         Left            =   5490
         TabIndex        =   5
         Top             =   15
         Width           =   1230
      End
      Begin VB.CommandButton cmdMemoFinalize 
         Caption         =   "OnAccount"
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
         Index           =   2
         Left            =   7950
         TabIndex        =   4
         Top             =   15
         Width           =   1230
      End
      Begin VB.CommandButton cmdMemoFinalize 
         Caption         =   "&Cash"
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
         Index           =   1
         Left            =   6720
         TabIndex        =   3
         Top             =   15
         Width           =   1230
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
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
         Left            =   9180
         TabIndex        =   2
         Top             =   15
         Width           =   1230
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
         Height          =   315
         Left            =   10410
         TabIndex        =   1
         Top             =   15
         Width           =   1230
      End
   End
   Begin TabDlg.SSTab sstOne 
      Height          =   2040
      Left            =   3015
      TabIndex        =   26
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   3598
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
      TabCaption(0)   =   "Customer && Postage"
      TabPicture(0)   =   "frmPurchaseReturn.frx":5151
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPostage"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCustomer"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Transporter && Karter"
      TabPicture(1)   =   "frmPurchaseReturn.frx":516D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraKarter"
      Tab(1).Control(1)=   "fraTransporter"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Good Recv && SubDistributor"
      TabPicture(2)   =   "frmPurchaseReturn.frx":5189
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraGoods"
      Tab(2).Control(1)=   "fraSubDist"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Account Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   6195
         TabIndex        =   81
         Top             =   1020
         Width           =   2400
         Begin MSMask.MaskEdBox txtAccountBalance 
            Height          =   300
            Left            =   75
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   210
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   529
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
         Height          =   1560
         Left            =   75
         TabIndex        =   71
         Top             =   45
         Width           =   6090
         Begin VB.PictureBox pboxB 
            Height          =   1260
            Left            =   60
            ScaleHeight     =   1200
            ScaleWidth      =   5895
            TabIndex        =   72
            Top             =   225
            Width           =   5955
            Begin VB.CommandButton cmdSelectID 
               DownPicture     =   "frmPurchaseReturn.frx":51A5
               Height          =   285
               Left            =   975
               Picture         =   "frmPurchaseReturn.frx":54E8
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   570
            End
            Begin VB.TextBox txtID 
               Height          =   285
               Left            =   0
               TabIndex        =   79
               ToolTipText     =   "ID"
               Top             =   0
               Width           =   975
            End
            Begin VB.TextBox txtAddress 
               Height          =   600
               Left            =   0
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   78
               ToolTipText     =   "ADDRESS"
               Top             =   300
               Width           =   3315
            End
            Begin VB.TextBox txtName 
               Height          =   280
               Left            =   1530
               TabIndex        =   77
               ToolTipText     =   "NAME"
               Top             =   0
               Width           =   4395
            End
            Begin VB.TextBox txtEmail 
               Height          =   280
               Left            =   3315
               TabIndex        =   76
               ToolTipText     =   "EMAIL"
               Top             =   615
               Width           =   2610
            End
            Begin VB.TextBox txtWebsite 
               Height          =   300
               Left            =   3315
               TabIndex        =   75
               ToolTipText     =   "WEB"
               Top             =   915
               Width           =   2610
            End
            Begin VB.TextBox txtCity 
               Height          =   300
               Left            =   0
               TabIndex        =   74
               ToolTipText     =   "CITY"
               Top             =   915
               Width           =   3300
            End
            Begin VB.TextBox txtPhones 
               Height          =   280
               Left            =   3315
               TabIndex        =   73
               ToolTipText     =   "PHONES"
               Top             =   300
               Width           =   2610
            End
         End
      End
      Begin VB.Frame fraSubDist 
         Caption         =   "Sub Distributor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   -68850
         TabIndex        =   67
         Top             =   60
         Width           =   2475
         Begin VB.PictureBox pboxG 
            Height          =   675
            Left            =   60
            ScaleHeight     =   615
            ScaleWidth      =   2280
            TabIndex        =   68
            Top             =   210
            Width           =   2340
            Begin VB.TextBox txtSDID 
               Height          =   285
               Left            =   0
               TabIndex        =   70
               ToolTipText     =   "SUB DISTRIBUTOR"
               Top             =   0
               Width           =   1290
            End
            Begin VB.TextBox txtSDName 
               Enabled         =   0   'False
               Height          =   285
               Left            =   0
               TabIndex        =   69
               TabStop         =   0   'False
               ToolTipText     =   "SD NAME"
               Top             =   330
               Width           =   2280
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
         Height          =   975
         Left            =   6180
         TabIndex        =   63
         Top             =   45
         Width           =   2460
         Begin VB.PictureBox pboxC 
            Height          =   675
            Left            =   60
            ScaleHeight     =   615
            ScaleWidth      =   2280
            TabIndex        =   64
            Top             =   210
            Width           =   2340
            Begin VB.TextBox txtPID 
               Height          =   285
               Left            =   0
               TabIndex        =   66
               ToolTipText     =   "SUB DISTRIBUTOR"
               Top             =   0
               Width           =   1290
            End
            Begin VB.TextBox txtPName 
               Enabled         =   0   'False
               Height          =   285
               Left            =   0
               TabIndex        =   65
               TabStop         =   0   'False
               ToolTipText     =   "SD NAME"
               Top             =   330
               Width           =   2280
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
         Height          =   1710
         Left            =   -75000
         TabIndex        =   54
         Top             =   30
         Width           =   4320
         Begin VB.PictureBox pboxD 
            Height          =   1485
            Left            =   105
            ScaleHeight     =   1425
            ScaleWidth      =   4050
            TabIndex        =   55
            Top             =   180
            Width           =   4110
            Begin VB.TextBox txtTEmail 
               Height          =   285
               Left            =   0
               TabIndex        =   62
               Top             =   1140
               Width           =   2010
            End
            Begin VB.TextBox txtTWebsite 
               Height          =   285
               Left            =   2040
               TabIndex        =   61
               Top             =   1140
               Width           =   2010
            End
            Begin VB.TextBox txtTaddress 
               Height          =   540
               Left            =   0
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   60
               Top             =   285
               Width           =   4050
            End
            Begin VB.TextBox txtTName 
               Height          =   285
               Left            =   825
               TabIndex        =   59
               Top             =   -15
               Width           =   3225
            End
            Begin VB.TextBox txtTPhones 
               Height          =   285
               Left            =   2040
               TabIndex        =   58
               Top             =   840
               Width           =   2010
            End
            Begin VB.TextBox txtTCity 
               Height          =   285
               Left            =   0
               TabIndex        =   57
               Top             =   840
               Width           =   2010
            End
            Begin VB.TextBox txtTID 
               Height          =   285
               Left            =   0
               TabIndex        =   56
               Top             =   -15
               Width           =   825
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
         Height          =   1710
         Left            =   -70665
         TabIndex        =   45
         Top             =   30
         Width           =   4275
         Begin VB.PictureBox pboxE 
            Height          =   1470
            Left            =   75
            ScaleHeight     =   1410
            ScaleWidth      =   4035
            TabIndex        =   46
            Top             =   180
            Width           =   4095
            Begin VB.TextBox txtKName 
               Height          =   285
               Left            =   810
               TabIndex        =   49
               Top             =   -15
               Width           =   3225
            End
            Begin VB.TextBox txtKID 
               Height          =   285
               Left            =   -15
               TabIndex        =   48
               Top             =   -15
               Width           =   825
            End
            Begin VB.TextBox Text4 
               Height          =   540
               Left            =   -15
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   47
               Top             =   285
               Width           =   4050
            End
            Begin MSMask.MaskEdBox txtKAmount 
               Height          =   285
               Left            =   1485
               TabIndex        =   50
               Top             =   1125
               Width           =   2550
               _ExtentX        =   4498
               _ExtentY        =   503
               _Version        =   393216
               Format          =   "###0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtKRate 
               Height          =   285
               Left            =   1485
               TabIndex        =   51
               Top             =   840
               Width           =   2550
               _ExtentX        =   4498
               _ExtentY        =   503
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
               TabIndex        =   53
               Top             =   840
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
               Left            =   30
               TabIndex        =   52
               Top             =   1125
               Width           =   1605
            End
         End
      End
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
         Height          =   1650
         Left            =   -74910
         TabIndex        =   27
         Top             =   60
         Width           =   6030
         Begin VB.PictureBox pboxF 
            Height          =   1305
            Left            =   75
            ScaleHeight     =   1245
            ScaleWidth      =   5820
            TabIndex        =   28
            Top             =   225
            Width           =   5880
            Begin VB.TextBox txtTerminal 
               Height          =   285
               Left            =   3750
               TabIndex        =   34
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox txtGRNo 
               Height          =   285
               Left            =   675
               TabIndex        =   33
               Top             =   0
               Width           =   2055
            End
            Begin VB.TextBox txtBundleWeight 
               Height          =   285
               Left            =   3750
               TabIndex        =   32
               Top             =   645
               Width           =   2055
            End
            Begin VB.TextBox txtBundleCount 
               Height          =   285
               Left            =   3750
               TabIndex        =   31
               Top             =   330
               Width           =   2055
            End
            Begin VB.ComboBox txtGRMode 
               Height          =   315
               Left            =   660
               TabIndex        =   30
               Top             =   615
               Width           =   2055
            End
            Begin VB.ComboBox txtToPayMode 
               Height          =   315
               Left            =   3750
               TabIndex        =   29
               Top             =   0
               Width           =   2055
            End
            Begin MSComCtl2.DTPicker txtGRDate 
               Height          =   300
               Left            =   675
               TabIndex        =   35
               Top             =   300
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "dd-MM-yy hh:mm tt"
               Format          =   22675459
               CurrentDate     =   38023
            End
            Begin MSMask.MaskEdBox txtGRAmount 
               Height          =   300
               Left            =   660
               TabIndex        =   36
               Top             =   945
               Width           =   2055
               _ExtentX        =   3625
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
               Left            =   2790
               TabIndex        =   44
               Top             =   15
               Width           =   750
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
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
               Left            =   -195
               TabIndex        =   43
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
               Left            =   2805
               TabIndex        =   42
               Top             =   330
               Width           =   750
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
               Height          =   285
               Left            =   2835
               TabIndex        =   41
               Top             =   675
               Width           =   750
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
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
               Left            =   -195
               TabIndex        =   40
               Top             =   600
               Width           =   750
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
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
               Left            =   -210
               TabIndex        =   39
               Top             =   300
               Width           =   750
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
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
               Left            =   -60
               TabIndex        =   38
               Top             =   915
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
               Height          =   285
               Left            =   2925
               TabIndex        =   37
               Top             =   975
               Width           =   840
            End
         End
      End
   End
End
Attribute VB_Name = "frmPurchaseReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MyCAPTION = "PurchaseReturn "
Private Const Stk = True

Private Const HiddenCols = "[frmSale-flxItems-HiddenCols]"
Private Const Main = "PRETURNMAIN"
Private Const Support = "PURCHASERETURN"
Private Const MainAddNew = "ADDNEW_PRETURNMain"
Private Const MainSelectView = "appview_PRETURNMain_Select_View"
Private Const PartyView = "appview_AllAccounts"
Private Const PartyInitial = "D%"
Private Const ItemSelectView = "appview_SelectItemPurchase"

Private Const MemoFormat = "\P\R0"
Dim DBRef As Integer
Dim StockValidationRequired As Boolean

Private Sub Form_Load()
    StockValidationRequired = Stk: Me.Move 0, 0
    txtDBDate.Value = Now: txtOrderDate.Value = Now: txtInvDate.Value = Now
    txtGRMode.AddItem "DIRECT": txtGRMode.AddItem "BANK": txtGRMode.AddItem "HOLD"
    txtToPayMode.AddItem "Paid-Full": txtToPayMode.AddItem "Paid-Half": txtToPayMode.AddItem "Paid-Zero": txtToPayMode.AddItem "ToPay-Full": txtToPayMode.AddItem "ToPay-Half": txtToPayMode.AddItem "ToPay-Zero"
    Call PictureBoxStatus(False)
    t = ReadFont("[flxItems-Font]", 0): flxOrder.FontName = t
    t = ReadFont("[flxItems-Font]", 1): flxOrder.FontSize = t
    t = ReadFont("[flxItems-Font]", 2): flxOrder.FontBold = t
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 188 And Shift = vbCtrlMask Then  '<
        For Each i In Split(GReadINI(HiddenCols), ",")
            flxOrder.ColHidden(i) = True
        Next
    End If
    If KeyCode = 190 And Shift = vbCtrlMask Then  '>
        For i = 0 To flxOrder.COLS - 1
            flxOrder.ColHidden(i) = False
        Next
    End If
    If KeyCode = vbKeyD And Shift = vbCtrlMask Then
        Call ApplyDiscounts      'Special Discount
    End If
    If KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call ApplyCurrencyPrice   'ApplyCurrencyPrice
    End If
    If KeyCode = vbKeyN And Shift = vbCtrlMask Then
        For i = 0 To 10
            flxOrder.AddItem "0"      'New Row
        Next
    End If
    If KeyCode = vbKeyF12 Then
        cmdSaveOrder_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPrintMemo
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

Private Sub txtDBRef_LostFocus()
    sSQL(0) = "Select * from " & Main & " where DBRef=" & Val(txtDBRef.Text)
    Call dbOpen(0): Call ClearsArray(0): Call FillsArray(0): Call dbClose(0)
    
    DBRef = Val(sArray(0))
    If DBRef = 0 Then
        PictureBoxStatus (False): txtDBRef = ""
    End If
    If Val(sArray(0)) <> 0 Then
        FillFormText 'Filling of Forms's text boxes.
        sSQL(1) = "Select * from " & Support & " where DBRef=" & Val(txtDBRef.Text) & " Order by Serial"
        Call dbOpen(1): Set flxOrder.DataSource = recs(1): Call dbClose(1)
        For Each i In Split(GReadINI(HiddenCols), ",")
            flxOrder.ColHidden(i) = True
        Next
        PictureBoxStatus (True)
    End If
    flxOrder.ColWidth(2) = 500: flxOrder.ColWidth(3) = 1500: flxOrder.ColWidth(4) = 4000: flxOrder.ColWidth(6) = 800: flxOrder.ColWidth(7) = 1500
    sstOne.Tab = 0
    cmdCalculate_Click
    txtStatus.Text = GetOrderStatus(Main, Val(txtDBRef.Text))
    txtAccountBalance.Text = WhatIsLedgerBalance(txtID.Text)
End Sub

Private Sub cmdNewDBRef_Click()
    X = MsgBox("Create new order?", vbYesNo)
    If X = vbYes Then
        txtDBDate.Value = Now: txtInvDate.Value = Now
        sSQL(0) = MainAddNew: dbOpen (0)
        Set recs(0) = recs(0).NextRecordset
        txtDBRef.Text = recs(0)!DBRef
        dbClose (0)
        sSQL(0) = ""
    End If
    txtDBRef_LostFocus
End Sub

Private Sub cmdSelectDBRef_Click()
    sSQL(0) = "SELECT * FROM " & MainSelectView & " ORDER BY 1 Desc"
    frmShow.Init sSQL(0): sSQL(0) = ""
    If sArray(0) <> "" Then
        txtDBRef.Text = sArray(0)
    End If
    txtDBRef_LostFocus
    txtAccountBalance.Text = WhatIsLedgerBalance(txtID.Text)
End Sub

Private Sub txtID_LostFocus()
    txtID.Enabled = True
    sSQL(0) = "Select * from " & PartyView & " where ID=" & Chr(39) & txtID.Text & Chr(39)
    Call dbOpen(0): ClearsArray (0): Call FillsArray(0)
    Call dbClose(0): sSQL(0) = ""
    txtID.Text = sArray(0): txtName.Text = sArray(1): txtAddress.Text = sArray(2): txtCity.Text = sArray(3): txtPhones.Text = sArray(4): txtEmail.Text = sArray(5): txtWebsite.Text = sArray(6)
    txtSDID.Text = sArray(8): txtTID.Text = sArray(9): txtKID.Text = sArray(10)
    txtTerminal.Text = "Patna-to-" & txtCity.Text
    txtAccountBalance.Text = WhatIsLedgerBalance(txtID.Text)
End Sub

Private Sub cmdSelectID_Click()
    sSQL(0) = "Select * from " & PartyView & " WHERE ID LIKE " & QT(PartyInitial) & " ORDER BY 2"
    frmShow.Init sSQL(0): sSQL(0) = ""
    If sArray(0) <> "" Then
        txtID.Text = sArray(0): txtName.Text = sArray(1): txtAddress.Text = sArray(2): txtCity.Text = sArray(3): txtPhones.Text = sArray(4): txtEmail.Text = sArray(5): txtWebsite.Text = sArray(6)
        txtSDID.Text = sArray(8): txtTID.Text = sArray(9): txtKID.Text = sArray(10)
        txtTerminal.Text = "Patna-to-" & txtCity.Text
    End If
    txtAccountBalance.Text = WhatIsLedgerBalance(txtID.Text)
End Sub

Private Sub cmbGRMode_Click()
    txtGRMode.Text = cmbGRMode.Text
End Sub

Private Sub txtGRAmount_Change()
    txtToPayMode_Click
End Sub

Private Sub txtSDID_Change()
    'disabled
End Sub

Private Sub txtToPayMode_Click()
    Select Case txtToPayMode.ListIndex
        Case 0
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0#: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
        Case 1
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0.5: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
        Case 2
            txtAddFreight.Text = Val(txtGRAmount.Text) * 1: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
        Case 3
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0: txtLessFreight.Text = Val(txtGRAmount.Text) * 1#
        Case 4
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0: txtLessFreight.Text = Val(txtGRAmount.Text) * 0.5
        Case 5
            txtAddFreight.Text = Val(txtGRAmount.Text) * 0: txtLessFreight.Text = Val(txtGRAmount.Text) * 0
    End Select
End Sub

Private Sub flxOrder_EnterCell()
    With flxOrder
        c = .Col
        Select Case c
            Case 1 To 21
                .Editable = flexEDKbdMouse
            Case Else
                .Editable = flexEDNone
        End Select
    End With
End Sub

Private Sub flxOrder_LeaveCell()
    With flxOrder
    For R = 1 To .ROWS - 1
        For c = 0 To .COLS - 1
        Select Case c
            Case 9 To 11, 13 To 16
                .TextMatrix(R, c) = Val(.TextMatrix(R, c))
        End Select
        Next
        If StockValidationRequired Then
            If (Val(.TextMatrix(R, 10)) + Val(.TextMatrix(R, 11))) > Val(.TextMatrix(R, 22)) Then
                .Cell(flexcpBackColor, R, 10, R, 11) = vbRed
            Else
                .Cell(flexcpBackColor, R, 10, R, 11) = .Cell(flexcpBackColor, R, 1)
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
        If flxOrder.SelectedRow(i) >= 1 Then
            sum = sum + Val(flxOrder.TextMatrix(flxOrder.SelectedRow(i), flxOrder.Col))
        End If
    Next
    Me.Caption = MyCAPTION & "Sum on col " & str(flxOrder.Col) & " = " & sum
End Sub

Private Sub flxOrder_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        If Val(flxOrder.TextMatrix(Row, Col)) = 0 Then
            frmSearchItem.Show vbModal
        Else
            ItemID = Val(flxOrder.TextMatrix(Row, Col))
        End If
        sSQL(10) = "SELECT * from " & ItemSelectView & " Where ItemID=" & ItemID
        dbOpen (10)
        If recs(10).RecordCount >= 1 Then
            For i = 0 To recs(10).FIELDS.Count - 1
                If i <> 10 And i <> 17 Then
                    flxOrder.TextMatrix(Row, i) = recs(10).FIELDS(i)
                Else
                    flxOrder.TextMatrix(Row, i) = Val(flxOrder.TextMatrix(Row, i))
                End If
            Next
            flxOrder.Col = 10: flxOrder.EditCell
        Else
            For i = 0 To flxOrder.COLS - 1
                flxOrder.TextMatrix(Row, i) = ""
            Next
        End If
    End If
    If Col = 3 Then
        sSQL(10) = "SELECT * from " & ItemSelectView & " Where ISBN=" & QT(flxOrder.TextMatrix(Row, Col))
        dbOpen (10)
        If recs(10).RecordCount >= 1 Then
            For i = 0 To recs(10).FIELDS.Count - 1
                If i <> 10 And i <> 17 Then
                    flxOrder.TextMatrix(Row, i) = recs(10).FIELDS(i)
                Else
                    flxOrder.TextMatrix(Row, i) = Val(flxOrder.TextMatrix(Row, i))
                End If
            Next
            flxOrder.Col = 10: flxOrder.EditCell
        Else
            For i = 0 To flxOrder.COLS - 1
                flxOrder.TextMatrix(Row, i) = ""
            Next
        End If
    End If
    If Col = 10 Then
        flxOrder.Col = 17: flxOrder.EditCell
    End If
    If Col = 17 Then
        flxOrder.Col = 3
        If flxOrder.Row < flxOrder.ROWS - 1 Then flxOrder.Row = flxOrder.Row + 1
    End If
    EnumerateGrid
    cmdCalculate_Click
End Sub

Private Sub flxOrder_AfterSort(ByVal Col As Long, Order As Integer)
    EnumerateGrid
End Sub

Private Sub flxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyI And Shift = vbCtrlMask And (flxOrder.Row > 1 And flxOrder.Row - 1 > 0) Then
        flxOrder.RowPosition(flxOrder.Row) = flxOrder.Row - 1
        flxOrder.Row = flxOrder.Row - 1
    End If
    If KeyCode = vbKeyJ And Shift = vbCtrlMask And (flxOrder.Row < flxOrder.ROWS - 1 And flxOrder.Row + 1 <= flxOrder.ROWS - 1) Then
        flxOrder.RowPosition(flxOrder.Row) = flxOrder.Row + 1
        flxOrder.Row = flxOrder.Row + 1
    End If
    EnumerateGrid
End Sub

Private Sub cmdSelectItem_Click()
    frmSelectItem.GetLink Me.flxOrder, Support
    EnumerateGrid
    Call ApplyDiscounts
    Call ApplyCurrencyPrice
    cmdCalculate_Click
End Sub

Private Sub cmdDeleteItem_Click()
    R = flxOrder.Row
    If flxOrder.ROWS > 1 And flxOrder.Row <> 0 Then flxOrder.RemoveItem R
    For i = 1 To flxOrder.ROWS - 1
        flxOrder.TextMatrix(i, 0) = i
    Next
    cmdCalculate_Click
End Sub

Private Sub cmdApplyCurrencyPrice_Click()
    
End Sub

Private Sub txtTID_LostFocus()
    sSQL(0) = "Select * from appview_AllAccounts where ID=" & Chr(39) & txtTID.Text & Chr(39)
    Call dbOpen(0): ClearsArray (0): FillsArray (0): Call dbClose(0): sSQL(0) = ""
    txtTID.Text = sArray(0)
    txtTName.Text = sArray(1)
    txtTaddress.Text = sArray(2)
    txtTCity.Text = sArray(3)
    txtTPhones.Text = sArray(4)
    txtTEmail.Text = sArray(5)
    txtTWebsite.Text = sArray(6)
End Sub

Private Sub txtKID_LostFocus()
    sSQL(0) = "Select * from appview_AllAccounts where ID=" & Chr(39) & txtKID.Text & Chr(39)
    Call dbOpen(0): ClearsArray (0): FillsArray (0): Call dbClose(0): sSQL(0) = ""
    txtKID.Text = sArray(0)
    txtKName.Text = sArray(1)
    txtKRate.Text = sArray(2)
End Sub

Private Sub cmdCalculate_Click()
    Dim Gross, Amt, DiscAmt, SDCommission, SDTotalCommission, TotalAmount, NetAmount, RoundOff As Currency
    Dim TempCount, AddItemCount, LessItemCount, Disc, SDDisc As Double
    ItemCount = 0
    txtToPayMode_Click
    TotalAmount = 0: NetAmount = 0: SDTotalCommission = 0
    For R = 1 To flxOrder.ROWS - 1
        flxOrder.TextMatrix(R, 15) = Val(flxOrder.TextMatrix(R, 13)) * Val(flxOrder.TextMatrix(R, 14))
        Gross = Val(flxOrder.TextMatrix(R, 10)) * Val(flxOrder.TextMatrix(R, 15))
        Disc = Val(flxOrder.TextMatrix(R, 17))
        DiscAmt = Gross * (Disc / 100)
        SDDisc = Val(flxOrder.TextMatrix(R, 20))
        SDDisc = Disc           ' Disabling SDDisc
        SDCommission = Gross * ((SDDisc - Disc) / 100)
        Amt = Gross - DiscAmt
        flxOrder.TextMatrix(R, 16) = Format(Gross, "###0.00")
        flxOrder.TextMatrix(R, 18) = Format(DiscAmt, "###0.00")
        flxOrder.TextMatrix(R, 19) = Format(Amt, "###0.00")
        flxOrder.TextMatrix(R, 21) = Format(SDCommission, "###0.00")
        TotalAmount = TotalAmount + Amt
        SDTotalCommission = SDTotalCommission + SDCommission
        'ItemCount = ItemCount + ((Val(flxOrder.TextMatrix(R, 10)) + Val(flxOrder.TextMatrix(R, 11)))) * Val(flxOrder.TextMatrix(R, 9))
        TempCount = ((Val(flxOrder.TextMatrix(R, 10)) + Val(flxOrder.TextMatrix(R, 11)))) * Val(flxOrder.TextMatrix(R, 9))
        If TempCount >= 0 Then
            AddItemCount = AddItemCount + TempCount
        Else
            LessItemCount = LessItemCount + TempCount
        End If
    Next
    
    txtKAmount.Text = RamjeeRound(Val(txtBundleCount.Text) * Val(txtKRate.Text))
    
    NetAmount = TotalAmount
    NetAmount = NetAmount - (NetAmount * Val(txtSplDisc.Text) / 100)
    NetAmount = NetAmount - (NetAmount * Val(txtBulkDisc.Text) / 100)
    
    NetAmount = NetAmount + Val(txtAddFreight.Text) + Val(txtAddMisc.Text)
    NetAmount = NetAmount - Val(txtLessFreight.Text) - Val(txtLessMisc.Text)
    NetAmount = NetAmount + Val(txtKAmount) + Val(txtPostage.Text)
    RoundOff = Round(NetAmount) - NetAmount
    NetAmount = Round(NetAmount)
    txtAmount.Text = Format(TotalAmount, "###0.00")
    txtSDCommission.Text = Format(SDTotalCommission, "###0.00")
    txtNetAmount.Text = Format(NetAmount, "###0.00")
    txtRoundOff.Text = Format(RoundOff, "###0.00")
    txtItemCount.Text = str(AddItemCount) & IIf(LessItemCount < 0, str(LessItemCount), "")
End Sub

Private Sub cmdDeleteBill_Click()
    If MsgBox("Confirm deletion?", vbYesNo + vbQuestion) = vbYes Then
        'Delete Memo and its related journal entries; but leaves the order intact.
        sSQL(0) = "DELETE JOURNAL WHERE MemoRef=" & QT(Format(Val(txtDBRef.Text), MemoFormat)): dbOpen (0): dbClose (0)
        MsgBox "Journal reference " & Format(Val(txtDBRef.Text), MemoFormat) & " deleted!", vbOKOnly + vbCritical
        sSQL(0) = "UPDATE " & Main & " SET STATUS=" & QT("ORDER") & " WHERE DBRef=" & Val(txtDBRef.Text)
        dbOpen (0): dbClose (0)
        txtDBRef_LostFocus
    End If
End Sub

Private Sub cmdSaveOrder_Click()
    Call flxOrder_LeaveCell: Call ApplyCurrencyPrice
    
    DeleteString = "Delete " & Support & " where DBRef=" & Val(txtDBRef.Text)
    SaveString = "Select * from " & Support & " where DBRef=" & Val(txtDBRef.Text) & " order by Serial"
    
    bs = GetOrderStatus(Main, Val(txtDBRef.Text))
    If UCase(bs) = "NEW" Or UCase(bs) = "ORDER" Then
        If flxOrder.ROWS > 1 Then
            toMain "ORDER"
            toSupport DeleteString, SaveString, flxOrder
            msgUITS Support & " Order Data Saved!"
        Else
            msgUITS "Blank Order Cannot be Saved!"
        End If
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
            sSQL(0) = "Update " & Main & " SET Status=" & QT(Operation) & " WHERE DBRef=" & Val(txtDBRef.Text): dbOpen (0): dbClose (0)
            If MakeMultipleJournalEntries(Support, Val(txtDBRef.Text), MemoFormat) Then
                sSQL(0) = "appproc_SetInvNo " & Val(txtDBRef.Text) & ", " & QT(Main): dbOpen (0): dbClose (0)
                msgUITS Operation & " operation successful!"
            Else
                sSQL(0) = "Update " & Main & " SET Status=" & QT("ORDER") & " WHERE DBRef=" & Val(txtDBRef.Text): dbOpen (0): dbClose (0)
            End If
        End If
        Call txtDBRef_LostFocus
    Else
        MsgBox "Current status is " & Trim(bs) & vbCrLf & Operation & " operation failed!", vbOKOnly + vbCritical
    End If
End Sub

Private Sub cmdPrint_Click()
    If Val(txtDBRef.Text) <> 0 Then
        frmPrintMemo.PrintIT Val(txtDBRef.Text), Support
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

Private Sub FillFormText()
    i = 1
'   txtDBRef.Text = sArray(i): i = i + 1
    txtDBDate.Value = sArray(i): i = i + 1
    txtStatus.Text = sArray(i): i = i + 1
    txtOrderRef.Text = sArray(i): i = i + 1
    txtOrderDate.Value = sArray(i): i = i + 1
    txtInvRef.Text = sArray(i): i = i + 1
    txtInvDate.Value = sArray(i): i = i + 1
    txtID.Text = sArray(i): i = i + 1
    txtName.Text = sArray(i): i = i + 1
    txtAddress.Text = sArray(i): i = i + 1
    txtCity.Text = sArray(i): i = i + 1
    txtPhones.Text = sArray(i): i = i + 1
    txtEmail.Text = sArray(i): i = i + 1
    txtWebsite.Text = sArray(i): i = i + 1
    txtTID.Text = sArray(i): i = i + 1
    txtTName.Text = sArray(i): i = i + 1
    txtTaddress.Text = sArray(i): i = i + 1
    txtTCity.Text = sArray(i): i = i + 1
    txtTPhones.Text = sArray(i): i = i + 1
    txtTEmail.Text = sArray(i): i = i + 1
    txtTWebsite.Text = sArray(i): i = i + 1
    txtTerminal.Text = sArray(i): i = i + 1
    txtKID.Text = sArray(i): i = i + 1
    txtKName.Text = sArray(i): i = i + 1
    txtKRate.Text = sArray(i): i = i + 1
    txtKAmount = sArray(i): i = i + 1
    txtPID.Text = sArray(i): i = i + 1
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
    txtAmount.Text = sArray(i): i = i + 1
    txtSplDisc.Text = sArray(i): i = i + 1
    txtBulkDisc.Text = sArray(i): i = i + 1
    txtAddMisc.Text = sArray(i): i = i + 1
    txtLessMisc.Text = sArray(i): i = i + 1
    txtAddFreight.Text = sArray(i): i = i + 1
    txtLessFreight.Text = sArray(i): i = i + 1
    txtRoundOff.Text = sArray(i): i = i + 1
    txtNetAmount.Text = sArray(i): i = i + 1
'   txtUserNo.Text = sArray(i): i = i + 1
    i = i + 1
    txtSubject.Text = sArray(i): i = i + 1
    txtComments.Text = sArray(i): i = i + 1
End Sub

Public Sub toMain(Status As String)
    sSQL(0) = "Select * from " & Main & " where DBRef=" & DBRef
    dbOpen (0)
'       recs(0)!DBRef = txtDBRef.Text
        recs(0)!DBDate = txtDBDate.Value
        recs(0)!Status = Status
        recs(0)!OrderRef = txtOrderRef.Text
        recs(0)!OrderDate = txtOrderDate.Value
        recs(0)!InvRef = txtInvRef.Text
        recs(0)!InvDate = txtInvDate.Value
        recs(0)!ID = txtID.Text
        recs(0)!Name = txtName.Text
        recs(0)!Address = txtAddress.Text
        recs(0)!City = txtCity.Text
        recs(0)!Phones = txtPhones.Text
        recs(0)!Email = txtEmail.Text
        recs(0)!Website = txtWebsite.Text
        recs(0)!TID = txtTID.Text
        recs(0)!TName = txtTName.Text
        recs(0)!Taddress = txtTaddress.Text
        recs(0)!TCity = txtTCity.Text
        recs(0)!TPhones = txtTPhones.Text
        recs(0)!TEmail = txtTEmail.Text
        recs(0)!TWebsite = txtTWebsite.Text
        recs(0)!Terminal = txtTerminal.Text
        recs(0)!KID = txtKID.Text
        recs(0)!KName = txtKName.Text
        recs(0)!KRate = Val(txtKRate.Text)
        recs(0)!KAmount = Val(txtKAmount.Text)
        recs(0)!PID = txtPID.Text
        recs(0)!PName = txtPName.Text
        recs(0)!Postage = Val(txtPostage.Text)
        recs(0)!SDID = txtSDID.Text
        recs(0)!SDName = txtSDName.Text
        recs(0)!SDCommission = Val(txtSDCommission.Text)
        recs(0)!GRNo = txtGRNo.Text
        recs(0)!GRDate = txtGRDate.Value
        recs(0)!GRMode = txtGRMode.ListIndex
        recs(0)!GRAmount = Val(txtGRAmount.Text)
        recs(0)!ToPayMode = txtToPayMode.ListIndex
        recs(0)!BundleCount = Val(txtBundleCount.Text)
        recs(0)!BundleWeight = Val(txtBundleWeight.Text)
        recs(0)!ItemCount = txtItemCount.Text
        recs(0)!Amount = Val(txtAmount.Text)
        recs(0)!SplDisc = Val(txtSplDisc.Text)
        recs(0)!BulkDisc = Val(txtBulkDisc.Text)
        recs(0)!AddMisc = Val(txtAddMisc.Text)
        recs(0)!LessMisc = Val(txtLessMisc.Text)
        recs(0)!AddFreight = Val(txtAddFreight.Text)
        recs(0)!LessFreight = Val(txtLessFreight.Text)
        recs(0)!RoundOff = Val(txtRoundOff.Text)
        recs(0)!NetAmount = Val(txtNetAmount.Text)
        recs(0)!UserNo = GUID
        recs(0)!Subject = txtSubject.Text
        recs(0)!Comments = txtComments.Text
    recs(0).Update: dbClose (0)
End Sub

Public Sub toSupport(ByVal DeleteStr As String, ByVal SaveStr As String, MyFlex As Control)
    Dim i As Integer
    sSQL(1) = DeleteStr
    Call dbOpen(1): Call dbClose(1)
    sSQL(1) = SaveStr
    Call dbOpen(1)
    With MyFlex
        For i = 1 To .ROWS - 1
            If ValidateItemID(Val(.TextMatrix(i, 2))) = True Then
                recs(1).AddNew
                recs(1)!Serial = i              '.TextMatrix(i, 0)
                recs(1)!DBRef = Val(txtDBRef.Text)
                recs(1)!ItemID = Left(.TextMatrix(i, 2), 20)
                recs(1)!ISBN = Trim(.TextMatrix(i, 3))
                recs(1)!ItemName = Left(.TextMatrix(i, 4), 200)
                recs(1)!PublisherID = Left(.TextMatrix(i, 5), 10)
                recs(1)!PublisherName = Left(.TextMatrix(i, 6), 200)
                recs(1)!Authors = Left(.TextMatrix(i, 7), 200)
                recs(1)!Edition = Left(.TextMatrix(i, 8), 10)
                recs(1)!Pkg = Val(.TextMatrix(i, 9))
                recs(1)!Qty = Val(.TextMatrix(i, 10))
                recs(1)!Free = Val(.TextMatrix(i, 11))
                recs(1)!Currency = .TextMatrix(i, 12)
                recs(1)!Price = Val(.TextMatrix(i, 13))
                recs(1)!CurrPrice = Val(.TextMatrix(i, 14))
                recs(1)!INRPrice = Val(.TextMatrix(i, 15))
                recs(1)!Gross = Val(.TextMatrix(i, 16))
                recs(1)!Discount = Val(.TextMatrix(i, 17))
                recs(1)!DiscountAmt = Val(.TextMatrix(i, 18))
                recs(1)!Amount = Val(.TextMatrix(i, 19))
                recs(1)!SDDiscount = Val(.TextMatrix(i, 20))
                recs(1)!SDCommission = Val(.TextMatrix(i, 21))
                recs(1)!Stock = Val(.TextMatrix(i, 22))
            End If
        Next
    End With
    recs(1).UpdateBatch
    dbClose (1)
    sSQL(1) = ""
End Sub

Public Function MaterialiseOrder(ByVal DBRef As Integer) As Boolean
    Dim B As Boolean
    Dim msg As String
    B = True
   
    If StockValidationRequired Then
        sSQL(1) = "SELECT A.Serial, A.ItemID, (A.QTY + A.FREE) AS REQD, B.AVLBL, (B.AVLBL - (A.QTY + A.FREE)) AS DIFF from " & Support & " A, Stock_View B where A.ItemID = B.ItemID AND A.DBRef=" & DBRef & " ORDER BY SERIAL"
        Call dbOpen(1)
        If recs(1).RecordCount <> 0 Then recs(1).MoveFirst
        Do Until recs(1).EOF
            If Val(recs(1)!Diff) < 0 Then
                msg = msg & vbCrLf & recs(1)!Serial & ". ItemID: " & recs(1)!ItemID & " is short on stock. Required=" & recs(1)!Reqd & " & Available=" & recs(1)!Avlbl
                B = False
            End If
            recs(1).MoveNext
        Loop
        Call dbClose(1)
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

Private Sub ApplyDiscounts()
    'apply discounts
End Sub

Private Sub ApplyCurrencyPrice()
    For i = 1 To flxOrder.ROWS - 1
        sSQL(9) = "SELECT CurrPrice from Currency Where Currency=" & QT(flxOrder.TextMatrix(i, 12))
        dbOpen (9)
        If Not recs(9).EOF Then flxOrder.TextMatrix(i, 14) = recs(9)!CurrPrice
        dbClose (9)
    Next
    cmdCalculate_Click
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

Private Function ValidateItemID(ByVal ItemID As Long) As Boolean
    sSQL(10) = "SELECT ITEMID FROM Items WHERE ItemID=" & ItemID
    dbOpen (10)
    If recs(10).RecordCount = 1 Then
        ValidateItemID = True
    Else
        ValidateItemID = False
    End If
    dbClose (10)
End Function

