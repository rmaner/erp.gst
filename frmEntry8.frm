VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.4#0"; "UniBox10.ocx"
Begin VB.Form frmEntry8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personal Accounts Master..."
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "frmEntry8.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VSFlex8UCtl.VSFlexGrid flxTable 
      Align           =   3  'Align Left
      Height          =   3990
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9180
      _cx             =   16192
      _cy             =   7038
      Appearance      =   2
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
      BackColorSel    =   16711680
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   4
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   4
      SelectionMode   =   1
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
      AutoSearch      =   2
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   3060
      Left            =   0
      ScaleHeight     =   3000
      ScaleWidth      =   9120
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3990
      Width           =   9180
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   0
         Top             =   0
         Width           =   900
         _Version        =   65540
         _ExtentX        =   1587
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Top             =   0
         Width           =   4665
         _Version        =   65540
         _ExtentX        =   8229
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   2
         Left            =   1230
         TabIndex        =   3
         Top             =   330
         Width           =   5595
         _Version        =   65540
         _ExtentX        =   9869
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   3
         Left            =   1230
         TabIndex        =   4
         Top             =   660
         Width           =   2805
         _Version        =   65540
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   9
         Left            =   1230
         TabIndex        =   10
         Top             =   1650
         Width           =   5595
         _Version        =   65540
         _ExtentX        =   9869
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   4
         Left            =   4050
         TabIndex        =   5
         Top             =   660
         Width           =   2775
         _Version        =   65540
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   5
         Left            =   1215
         TabIndex        =   6
         Top             =   990
         Width           =   2820
         _Version        =   65540
         _ExtentX        =   4974
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   10
         Left            =   1230
         TabIndex        =   11
         Top             =   1980
         Width           =   1395
         _Version        =   65540
         _ExtentX        =   2466
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   11
         Left            =   2640
         TabIndex        =   12
         Top             =   1980
         Width           =   1365
         _Version        =   65540
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   12
         Left            =   4035
         TabIndex        =   13
         Top             =   1980
         Width           =   1365
         _Version        =   65540
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   21
         Left            =   5700
         TabIndex        =   14
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
         _Version        =   65540
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   17
         Left            =   1230
         TabIndex        =   19
         Top             =   2640
         Width           =   1395
         _Version        =   65540
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   18
         Left            =   2640
         TabIndex        =   21
         Top             =   2640
         Width           =   1365
         _Version        =   65540
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   13
         Left            =   1230
         TabIndex        =   15
         Top             =   2310
         Width           =   1395
         _Version        =   65540
         _ExtentX        =   2466
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   14
         Left            =   2640
         TabIndex        =   16
         Top             =   2310
         Width           =   1365
         _Version        =   65540
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   15
         Left            =   4035
         TabIndex        =   17
         Top             =   2310
         Width           =   1365
         _Version        =   65540
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   16
         Left            =   5430
         TabIndex        =   18
         Top             =   2310
         Width           =   1395
         _Version        =   65540
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin VSFlex8UCtl.VSFlexGrid flx 
         Height          =   2970
         Left            =   6840
         TabIndex        =   37
         Top             =   -15
         Width           =   2280
         _cx             =   4022
         _cy             =   5239
         Appearance      =   3
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
         BackColorSel    =   255
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   32768
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
         AutoSearchDelay =   2
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
      Begin UniToolbox.UniText txtSearch 
         Height          =   315
         Left            =   5430
         TabIndex        =   20
         Top             =   2640
         Width           =   1395
         _Version        =   65540
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   19
         Left            =   4035
         TabIndex        =   22
         Top             =   2640
         Width           =   1365
         _Version        =   65540
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         BackColor       =   12648384
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   6
         Left            =   4050
         TabIndex        =   7
         Top             =   990
         Width           =   2775
         _Version        =   65540
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   7
         Left            =   1215
         TabIndex        =   8
         Top             =   1320
         Width           =   2790
         _Version        =   65540
         _ExtentX        =   4921
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   8
         Left            =   4020
         TabIndex        =   40
         Top             =   1320
         Width           =   2820
         _Version        =   65540
         _ExtentX        =   4974
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   22
         Left            =   5985
         TabIndex        =   41
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
         _Version        =   65540
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   23
         Left            =   6255
         TabIndex        =   42
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
         _Version        =   65540
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   24
         Left            =   6540
         TabIndex        =   43
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
         _Version        =   65540
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
      End
      Begin UniToolbox.UniText txtFields 
         Height          =   315
         Index           =   20
         Left            =   5430
         TabIndex        =   44
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
         _Version        =   65540
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   109
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "PAN/GSTIN:"
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
         Index           =   7
         Left            =   60
         TabIndex        =   39
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Ac/Dt/OB/CL:"
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
         Index           =   1
         Left            =   75
         TabIndex        =   36
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "GRP/TYPE:"
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
         Index           =   15
         Left            =   75
         TabIndex        =   35
         Top             =   2655
         Width           =   1140
      End
      Begin VB.Label lblLabels 
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
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   34
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "City/State: "
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
         Index           =   3
         Left            =   75
         TabIndex        =   33
         Top             =   705
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "ShipAddr:"
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
         Index           =   4
         Left            =   75
         TabIndex        =   32
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone/Email:"
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
         Index           =   5
         Left            =   60
         TabIndex        =   31
         Top             =   1020
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "ID/Name:"
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
         Index           =   0
         Left            =   75
         TabIndex        =   30
         Top             =   45
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "T/K/P:"
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
         Index           =   12
         Left            =   75
         TabIndex        =   29
         Top             =   2010
         Width           =   1140
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   9120
      TabIndex        =   27
      Top             =   7050
      Width           =   9180
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Top             =   15
         Width           =   1590
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3810
         TabIndex        =   24
         Top             =   15
         Width           =   1590
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5595
         TabIndex        =   25
         Top             =   15
         Width           =   1590
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7380
         TabIndex        =   26
         Top             =   15
         Width           =   1590
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2025
         TabIndex        =   23
         Top             =   15
         Width           =   1590
      End
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone/Email:"
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
      Index           =   6
      Left            =   0
      TabIndex        =   38
      Top             =   30
      Width           =   1140
   End
End
Attribute VB_Name = "frmEntry8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(5) As New clsData
Dim flds, SQ, FStr, Grp, GrpLetter As String

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    flx.Editable = flexEDKbdMouse
    flx.DataMode = flexDMBound
    CN(0).dbOpen "SELECT GRP, Initial FROM GRP ORDER BY 1"
    Set flx.DataSource = CN(0).recs
    flds = "ID, Name, Address, City, State, Phones, Email, PAN, GSTIN, ShipAddress, TID, KID, PID, ACCOUNT, OBDATE, OB, CreditLimit, GRP, Type, Code, DiscGrp, DiscTplt, TEMP1, TEMP2, TEMP3"
    flx_EnterCell
    flx.Row = 1
End Sub

Private Sub Form_Resize()
    flx.Width = Me.ScaleWidth
    flxTable.Width = Me.ScaleWidth
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyN And Shift = vbCtrlMask Then
        cmdUpdate_Click
        cmdAdd_Click
    End If
    If KeyCode = vbKeyF12 Then cmdUpdate_Click
    KeyCode = 0
End Sub

Private Sub flxTable_EnterCell()
    On Error Resume Next
    For i = 0 To flxTable.COLS - 1
        txtFields(i).Text = flxTable.TextMatrix(flxTable.Row, i)
    Next
End Sub

Private Sub flxTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = vbCtrlMask Then
        SaveGrid Me.flxTable
    End If
End Sub

Private Sub flx_EnterCell()
    If flx.ROWS > 1 Then
        Grp = flx.TextMatrix(flx.Row, 0): GrpLetter = flx.TextMatrix(flx.Row, 1)
        CN(1).dbOpen "SELECT " & flds & " from Personal Where Grp=" & QT(Grp) & " ORDER BY 1 DESC", 1
        Set flxTable.DataSource = CN(1).recs
        flxTable.DataMode = flexDMBoundImmediate
        flxTable_EnterCell
    End If
    flxTable.AutoSize 0, flxTable.COLS - 1
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    txtFields(Index).SelStart = 0: txtFields(Index).SelLength = Len(txtFields(Index).Text)
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Static SearchedRow As Long
    If KeyCode = 13 Then
        If SearchedRow = -1 Then
            SearchedRow = flxTable.FindRow(Trim(txtFields(Index).Text), , Index, False, False)
        Else
            SearchedRow = flxTable.FindRow(Trim(txtFields(Index).Text), SearchedRow + 1, Index, False, False)
        End If
        If SearchedRow <> -1 Then
            flxTable.Row = SearchedRow: flxTable.ShowCell SearchedRow, Index
            txtFields(Index).BackColor = vbCyan
        End If
    Else
        txtFields(Index).BackColor = vbWhite
    End If
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And Shift = vbCtrlMask Then
        If flxTable.Row > 1 Then flxTable.Row = flxTable.Row - 1
    End If
    If KeyCode = vbKeyDown And Shift = vbCtrlMask Then
        If flxTable.Row < flxTable.ROWS - 1 Then flxTable.Row = flxTable.Row + 1
    End If
    flxTable.ShowCell flxTable.Row, 0
End Sub

Private Sub cmdAdd_Click()
    Dim MaxID As String
    MaxID = ""
    CN(2).dbOpen "SELECT MAX(ID) AS MaxID FROM PERSONAL WHERE GRP=" & QT(Grp), 1
    If Not IsNull(CN(2).recs!MaxID) Then MaxID = CN(2).recs!MaxID
    NewID = Format(Val(Right(MaxID, 4)) + 1, "\" & GrpLetter & "0000")
    
    CN(3).dbOpen "INSERT INTO PERSONAL (ID,GRP) VALUES (" & QT(NewID) & ", " & QT(UCase(Grp)) & ")", 1
    CN(1).recs.Requery: flxTable.Refresh
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    rw = flxTable.Row
    For i = 1 To flxTable.COLS - 1
        If flxTable.Row <> 0 Then flxTable.TextMatrix(flxTable.Row, i) = txtFields(i).Text
    Next
    CN(1).recs.Requery: flxTable.Refresh
    flxTable.Row = rw
End Sub

Private Sub cmdDelete_Click()
    Dim rw As Integer
    rw = flxTable.Row
    If MsgBox("Confirm deletion?", vbYesNo + vbQuestion) = vbYes Then
        CN(4).dbOpen "DELETE PERSONAL WHERE ID=" & QT(flxTable.TextMatrix(rw, 0)), 1
        CN(1).recs.Requery: flxTable.Refresh
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Static SearchedRow As Long
    If KeyCode = 13 Then
        SearchText = Trim(txtSearch.Text)
        If SearchedRow = -1 Then
            SearchedRow = flxTable.FindRow(SearchText, , flxTable.Col, False, False)
        Else
            SearchedRow = flxTable.FindRow(SearchText, SearchedRow + 1, flxTable.Col, False, False)
        End If
        If SearchedRow = -1 Then
            txtSearch.BackColor = vbRed
        Else
            flxTable.Row = SearchedRow
            flxTable.ShowCell SearchedRow, flxTable.Col
            txtSearch.BackColor = vbYellow
        End If
    End If
    txtSearch.SetFocus
End Sub

