VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.4#0"; "UniBox10.ocx"
Begin VB.Form frmItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items..."
   ClientHeight    =   8670
   ClientLeft      =   135
   ClientTop       =   330
   ClientWidth     =   12165
   Icon            =   "frmItems.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   12105
      TabIndex        =   60
      Top             =   2820
      Width           =   12165
      Begin UniToolbox.UniText txtProducerID 
         Height          =   315
         Left            =   1605
         TabIndex        =   34
         Top             =   0
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483648
         BorderStyle     =   1
         Text            =   "0050003000300030"
         BackColor       =   -2147483648
      End
      Begin UniToolbox.UniText txtProducerName 
         Height          =   315
         Left            =   2700
         TabIndex        =   35
         Top             =   0
         Width           =   4515
         _Version        =   65540
         _ExtentX        =   7964
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483648
         BorderStyle     =   1
         Text            =   "00470045004e004500520041004c"
         BackColor       =   -2147483648
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Producer: "
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
         Index           =   11
         Left            =   330
         TabIndex        =   61
         Top             =   30
         Width           =   1245
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   12105
      TabIndex        =   51
      Top             =   8220
      Width           =   12165
      Begin VB.CommandButton cmdDuplicate 
         Caption         =   "&Duplicate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5089
         TabIndex        =   30
         Top             =   30
         Width           =   1620
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3319
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   30
         Width           =   1620
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10399
         TabIndex        =   33
         Top             =   30
         Width           =   1620
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8629
         TabIndex        =   32
         Top             =   30
         Width           =   1620
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6859
         TabIndex        =   31
         Top             =   30
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1556
         TabIndex        =   28
         Top             =   30
         Width           =   1620
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   5040
      Left            =   0
      ScaleHeight     =   4980
      ScaleWidth      =   12105
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3180
      Width           =   12165
      Begin VB.PictureBox Picture3 
         Height          =   315
         Left            =   4380
         ScaleHeight     =   255
         ScaleWidth      =   2790
         TabIndex        =   56
         Top             =   435
         Width           =   2850
         Begin UniToolbox.UniText txtSearch 
            Height          =   255
            Left            =   945
            TabIndex        =   57
            Top             =   0
            Width           =   1380
            _Version        =   65540
            _ExtentX        =   2434
            _ExtentY        =   450
            _StockProps     =   109
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
         End
         Begin UniToolbox.UniText txtSearchCol 
            Height          =   255
            Left            =   2340
            TabIndex        =   58
            Top             =   0
            Width           =   450
            _Version        =   65540
            _ExtentX        =   794
            _ExtentY        =   450
            _StockProps     =   109
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   1
            Text            =   "0031"
            Alignment       =   2
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Search for "
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
            Index           =   21
            Left            =   0
            TabIndex        =   59
            Top             =   30
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   330
         Left            =   4440
         ScaleHeight     =   270
         ScaleWidth      =   2760
         TabIndex        =   53
         Top             =   3870
         Width           =   2820
         Begin UniToolbox.UniText txtCurrentStock 
            Height          =   315
            Left            =   1440
            TabIndex        =   54
            Top             =   -30
            Width           =   1155
            _Version        =   65540
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   109
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            BackColor       =   -2147483624
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Current Stk: "
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
            Index           =   4
            Left            =   180
            TabIndex        =   55
            Top             =   30
            Width           =   1110
         End
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Top             =   75
         Width           =   1365
         _Version        =   65540
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   1
         Left            =   1620
         TabIndex        =   2
         Top             =   435
         Width           =   1380
         _Version        =   65540
         _ExtentX        =   2434
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   3
         Left            =   1620
         TabIndex        =   4
         Top             =   795
         Width           =   5625
         _Version        =   65540
         _ExtentX        =   9922
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   4
         Left            =   1620
         TabIndex        =   5
         Top             =   1155
         Width           =   5625
         _Version        =   65540
         _ExtentX        =   9922
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   5
         Left            =   1620
         TabIndex        =   6
         Top             =   1500
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   7
         Left            =   1620
         TabIndex        =   8
         Top             =   1830
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   8
         Left            =   2700
         TabIndex        =   9
         Top             =   1830
         Width           =   1200
         _Version        =   65540
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   9
         Left            =   1620
         TabIndex        =   10
         Top             =   2220
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   10
         Left            =   2700
         TabIndex        =   11
         Top             =   2220
         Width           =   1200
         _Version        =   65540
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   12
         Left            =   1620
         TabIndex        =   13
         Top             =   2550
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   14
         Left            =   1620
         TabIndex        =   15
         Top             =   2880
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   15
         Left            =   2700
         TabIndex        =   16
         Top             =   2880
         Width           =   1200
         _Version        =   65540
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   16
         Left            =   1620
         TabIndex        =   17
         Top             =   3210
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   17
         Left            =   1620
         TabIndex        =   18
         Top             =   3540
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   18
         Left            =   2700
         TabIndex        =   19
         Top             =   3540
         Width           =   1200
         _Version        =   65540
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   19
         Left            =   1620
         TabIndex        =   20
         Top             =   3870
         Width           =   2805
         _Version        =   65540
         _ExtentX        =   4960
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   20
         Left            =   1620
         TabIndex        =   21
         Top             =   4200
         Width           =   2805
         _Version        =   65540
         _ExtentX        =   4960
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   21
         Left            =   1620
         TabIndex        =   22
         Top             =   4590
         Width           =   1065
         _Version        =   65540
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   22
         Left            =   5370
         TabIndex        =   23
         Top             =   4470
         Visible         =   0   'False
         Width           =   630
         _Version        =   65540
         _ExtentX        =   1111
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Locked          =   -1  'True
      End
      Begin UniToolbox.UniText txtActiveCol 
         Height          =   315
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   450
         _Version        =   65540
         _ExtentX        =   794
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Enabled         =   0   'False
         Text            =   "0031"
         Alignment       =   2
      End
      Begin VSFlex8UCtl.VSFlexGrid flx 
         Height          =   4770
         Left            =   7425
         TabIndex        =   26
         Top             =   120
         Width           =   4500
         _cx             =   7937
         _cy             =   8414
         Appearance      =   1
         BorderStyle     =   0
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
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   255
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
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
         Rows            =   50
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
         AutoSearch      =   2
         AutoSearchDelay =   3
         MultiTotals     =   0   'False
         SubtotalPosition=   1
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   23
         Left            =   6000
         TabIndex        =   24
         Top             =   4470
         Visible         =   0   'False
         Width           =   630
         _Version        =   65540
         _ExtentX        =   1111
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
         Locked          =   -1  'True
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   24
         Left            =   6630
         TabIndex        =   25
         Top             =   4470
         Visible         =   0   'False
         Width           =   630
         _Version        =   65540
         _ExtentX        =   1111
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
         Locked          =   -1  'True
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Top             =   435
         Width           =   1380
         _Version        =   65540
         _ExtentX        =   2434
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   6
         Left            =   2700
         TabIndex        =   7
         Top             =   1500
         Width           =   4545
         _Version        =   65540
         _ExtentX        =   8017
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   11
         Left            =   3930
         TabIndex        =   12
         Top             =   2220
         Width           =   1200
         _Version        =   65540
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   345
         Index           =   13
         Left            =   2715
         TabIndex        =   14
         Top             =   2535
         Width           =   1200
         _Version        =   65540
         _ExtentX        =   2117
         _ExtentY        =   609
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin MSForms.CommandButton cmdGSTDiscCalc 
         Height          =   375
         Left            =   3930
         TabIndex        =   27
         Top             =   2535
         Width           =   3310
         Caption         =   "IncludingTax-DiscCalulator"
         Size            =   "5838;661"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "InitStock: "
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
         Index           =   10
         Left            =   120
         TabIndex        =   50
         Top             =   3900
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "GST/Cess: "
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
         Index           =   9
         Left            =   90
         TabIndex        =   49
         Top             =   2880
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Curr/MRP/SRP:"
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
         Index           =   7
         Left            =   15
         TabIndex        =   48
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Item/HSNCode:"
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
         Index           =   1
         Left            =   15
         TabIndex        =   47
         Top             =   465
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "ItemID:"
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
         Index           =   0
         Left            =   15
         TabIndex        =   46
         Top             =   105
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Packing/ Unit: "
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
         Index           =   6
         Left            =   15
         TabIndex        =   45
         Top             =   1890
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Producer: "
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
         Index           =   5
         Left            =   15
         TabIndex        =   44
         Top             =   1545
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "MakerAuthor:"
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
         Index           =   3
         Left            =   15
         TabIndex        =   43
         Top             =   1170
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "ItemName:"
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
         Index           =   2
         Left            =   15
         TabIndex        =   42
         Top             =   810
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "PDisc/SDisc:"
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
         Index           =   12
         Left            =   30
         TabIndex        =   41
         Top             =   2580
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Version:"
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
         Index           =   14
         Left            =   45
         TabIndex        =   40
         Top             =   3210
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Mfg/Exp Date:"
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
         Index           =   15
         Left            =   75
         TabIndex        =   39
         Top             =   3570
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "ItemID1: "
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
         Index           =   16
         Left            =   120
         TabIndex        =   38
         Top             =   4620
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "WareLocation: "
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
         Index           =   17
         Left            =   120
         TabIndex        =   37
         Top             =   4245
         Width           =   1515
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxTable 
      Align           =   3  'Align Left
      Height          =   2820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13560
      _cx             =   23918
      _cy             =   4974
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
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   4
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   30
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bLoadLists As Boolean
Private CN(10) As New clsData

Private Sub Form_Load()
    Me.Move 0, 0
    bLoadLists = True
    mdiOne.SetFormFont Me
    
    CN(0).dbOpen "SELECT * FROM Items ORDER BY 1 DESC"
    Set flxTable.DataSource = CN(0)
    flxTable.DataMode = flexDMBoundImmediate
    If flxTable.ROWS > 1 Then flxTable.Row = 1
    flxTable_EnterCell
End Sub

Private Sub Form_Activate()
    If UserRights = 0 Or UserRights = 1 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    flx.Width = Me.ScaleWidth
    flxTable.Width = Me.ScaleWidth
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 188 And Shift = vbCtrlMask Then Call ColShowHide(1)             '< Hide
    If KeyCode = 190 And Shift = vbCtrlMask Then Call ColShowHide(0)             '> Show
    If KeyCode = vbKeyN And Shift = vbCtrlMask Then
        cmdUpdate_Click
        cmdAdd_Click
    End If
    If KeyCode = vbKeyF12 Then cmdUpdate_Click
    If KeyCode = vbKeyF1 Then txtProducerID_DblClick
End Sub

Private Sub cmdGSTDiscCalc_Click()
    A = Val(txtFields(11).Text)
    P = A
    G = Val(txtFields(14).Text)
    d = (1 + G / 100 - A / P) / (0.01 + G * 0.0001)
    txtFields(13).Text = Val(d)
End Sub

Private Sub flxTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = vbCtrlMask Then
        SaveGrid Me.flxTable
    End If
End Sub

Private Sub flxTable_DblClick()
    CL = flxTable.Col
    If txtFields(CL).Enabled = True And txtFields(CL).Visible = True Then txtFields(CL).SetFocus
End Sub

Private Sub flxTable_EnterCell()
    txtSearchCol.Text = flxTable.Col
    For i = 0 To flxTable.COLS - 1
        txtFields(i).Text = flxTable.TextMatrix(flxTable.Row, i)
    Next
    CN(5).dbOpen "appproc_ReturnStock " & Val(flxTable.TextMatrix(flxTable.Row, 0))
    txtCurrentStock.Text = CN(5).recs!AVLBL
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    CN(1).dbOpen "INSERT INTO Items (ProducerID, ProducerName) VALUES (" & QT(txtProducerID.Text) & ", " & QT(txtProducerName.Text) & ")"
    cmdRefresh_Click
    flxTable_EnterCell
    txtFields(1).SetFocus
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    For i = 1 To flxTable.COLS - 1
        flxTable.TextMatrix(flxTable.Row, i) = txtFields(i).Text
    Next
    msgUITS "Updated"
End Sub

Private Sub cmdDuplicate_Click()
    CN(2).dbOpen "appproc_DuplicateItem " & Val(txtFields(0).Text)
    cmdRefresh_Click
End Sub

Private Sub cmdDelete_Click()
    Dim rw As Integer
    rw = flxTable.Row
    If MsgBox("Confirm deletion?", vbYesNo + vbQuestion) = vbYes Then
        CN(3).dbOpen "DELETE Items WHERE " & flxTable.TextMatrix(0, 0) & "=" & Val(txtFields(0).Text)
        cmdRefresh_Click
        If rw - 1 > 0 Then flxTable.Row = rw - 1
    End If
End Sub

Private Sub cmdRefresh_Click()
    CN(0).Requery
    flxTable.DataRefresh
    If flxTable.ROWS > 1 Then
        flxTable.Row = 1
        flxTable.ShowCell 1, 0
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set flxTable.DataSource = Nothing
End Sub

Private Sub flx_DblClick()
    AF = Val(txtActiveCol.Text)
    If AF >= txtFields.LBound And AF <= txtFields.UBound And CHECKSIDEFIELDS(AF) Then
        For i = 0 To flx.COLS - 1
            txtFields(AF + i).Text = flx.TextMatrix(flx.Row, i)
        Next
    End If
    'SendKeys vbTab
End Sub


Private Sub txtFields_GotFocus(Index As Integer)
    txtActiveCol.Text = Index
    txtFields(Index).SelStart = 0: txtFields(Index).SelLength = Len(txtFields(Index).Text)
    If bLoadLists = True Then
        flx.Clear
        Select Case Index
            Case 5:  CN(4).dbOpen "SELECT ID, Name FROM Personal Where ID like " & QT("P%") & " ORDER BY 1": Set flx.DataSource = CN(4)
        End Select
        rw = flx.FindRow(txtFields(Index).Text, 1, 0, False, False)
        If rw >= 1 Then
            flx.Row = rw
            flx.ShowCell rw, 0
        End If
    End If
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    rw = flx.FindRow(txtFields(Index).Text, 1, 0, False, False)
    If rw >= 1 Then
        flx.Row = rw: flx.ShowCell rw, 0
    End If
    If KeyCode = vbKeyUp And Shift = vbCtrlMask Then
        If flxTable.Row > 1 Then flxTable.Row = flxTable.Row - 1
        flxTable.ShowCell flxTable.Row, 0
        KeyCode = vbKeyA
    End If
    If KeyCode = vbKeyDown And Shift = vbCtrlMask Then
        If flxTable.Row < flxTable.ROWS - 1 Then flxTable.Row = flxTable.Row + 1
        flxTable.ShowCell flxTable.Row, 0
        KeyCode = vbKeyA
    End If

    If KeyCode = vbKeyLeft And Shift = vbCtrlMask Then
        txtFields(Index).Text = flxTable.TextMatrix(flxTable.Row, Index)
    End If
    
    If KeyCode = vbKeyRight And Shift = vbCtrlMask Then
        txtFields(Index).Text = flx.Text
    End If
    
    If KeyCode = vbKeyReturn Then
        flx_DblClick
    End If
End Sub

Private Sub txtProducerID_DblClick()
    SQ = "SELECT ID, NAME,CODE FROM PERSONAL WHERE ID LIKE " & QT("P%") & " ORDER BY 2"
    frmShow.Init SQ
    If sArray(0) <> "" Then
        txtProducerID.Text = sArray(0)
        txtProducerName.Text = sArray(1)
    End If
End Sub

Private Sub txtProducerID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CN(6).dbOpen "SELECT ID, NAME,CODE FROM PERSONAL WHERE ID=" & QT(txtProducerID.Text)
        If Not CN(6).recs.EOF Then
            txtProducerID.Text = CN(6).recs!id
            txtProducerName.Text = CN(6).recs!Name
        End If
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Static SearchedRow As Long
    If KeyCode = 13 Then
        SearchText = Trim(IIf(Val(txtSearchCol.Text) = ISBNCol, ParseISBN(txtSearch.Text), txtSearch.Text))
        If SearchedRow = -1 Then
            SearchedRow = flxTable.FindRow(SearchText, , Val(txtSearchCol.Text), False, False)
        Else
            SearchedRow = flxTable.FindRow(SearchText, SearchedRow + 1, Val(txtSearchCol.Text), False, False)
        End If
        If SearchedRow = -1 Then
            txtSearch.BackColor = vbRed
        Else
            flxTable.Row = SearchedRow
            flxTable.ShowCell SearchedRow, Val(txtSearchCol.Text)
            txtSearch.BackColor = vbYellow
        End If
    End If
    txtSearch.SetFocus
End Sub

Private Sub txtSearchCol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtSearch.SetFocus
End Sub

Private Sub ColShowHide(ByVal opt As Integer)
    If opt = 0 Then     'Show
        For i = 0 To flxTable.COLS - 1
            flxTable.ColHidden(i) = False
            txtFields(i).Visible = True
        Next
    End If
    If opt = 1 Then     'Hide
        For Each i In Split(mdiOne.sckGo.GReadINI("[Items-HiddenCols]"), ",")
            flxTable.ColHidden(i) = True
            txtFields(i).Visible = False
        Next
    End If
End Sub

Private Function CHECKSIDEFIELDS(ByVal i As Integer) As Boolean
    Select Case i
        Case 5:
            CHECKSIDEFIELDS = True
        Case Else
            CHECKSIDEFIELDS = False
    End Select
End Function
