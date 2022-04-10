VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTRN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Transactions..."
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTRN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4170
      Left            =   15
      TabIndex        =   28
      Top             =   0
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   7355
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "New Transaction..."
      TabPicture(0)   =   "frmTRN.frx":4E0E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Edit Transaction..."
      TabPicture(1)   =   "frmTRN.frx":4E2A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Print..."
      TabPicture(2)   =   "frmTRN.frx":4E46
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   3705
         Left            =   105
         TabIndex        =   37
         Top             =   345
         Width           =   7035
         Begin VB.TextBox txtCrName 
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
            Height          =   315
            Left            =   1830
            TabIndex        =   6
            ToolTipText     =   "Name "
            Top             =   1410
            Width           =   4710
         End
         Begin VB.TextBox txtCrID 
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
            Left            =   900
            TabIndex        =   5
            ToolTipText     =   "Payee's ID"
            Top             =   1410
            Width           =   915
         End
         Begin VB.CommandButton cmdSelectCrID 
            DownPicture     =   "frmTRN.frx":4E62
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6570
            Picture         =   "frmTRN.frx":51A5
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Open Payee"
            Top             =   1410
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   390
            Left            =   3555
            TabIndex        =   13
            ToolTipText     =   "Adds New Record"
            Top             =   3165
            Width           =   1560
         End
         Begin VB.CommandButton cmdQuit 
            Caption         =   "&Quit"
            Height          =   390
            Left            =   5280
            TabIndex        =   14
            Top             =   3165
            Width           =   1560
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   390
            Left            =   90
            TabIndex        =   11
            ToolTipText     =   "Adds New Record"
            Top             =   3165
            Width           =   1560
         End
         Begin VB.TextBox txtAmount 
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
            Left            =   900
            TabIndex        =   8
            ToolTipText     =   "Amount"
            Top             =   1815
            Width           =   1350
         End
         Begin VB.TextBox txtNarration 
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
            Left            =   900
            TabIndex        =   10
            ToolTipText     =   "Narration"
            Top             =   2625
            Width           =   6045
         End
         Begin VB.TextBox txtSerial 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   900
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Serial No."
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox txtDrName 
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
            Height          =   315
            Left            =   1830
            TabIndex        =   3
            ToolTipText     =   "Name "
            Top             =   1005
            Width           =   4710
         End
         Begin VB.TextBox txtDrID 
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
            Left            =   900
            TabIndex        =   2
            ToolTipText     =   "Payee's ID"
            Top             =   1005
            Width           =   915
         End
         Begin VB.CommandButton cmdSelectDrID 
            DownPicture     =   "frmTRN.frx":54E8
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6570
            Picture         =   "frmTRN.frx":582B
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Open Payee"
            Top             =   1005
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.TextBox txtAmountWords 
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
            Height          =   315
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Amount in words"
            Top             =   1815
            Width           =   4665
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            Height          =   390
            Left            =   1815
            TabIndex        =   12
            ToolTipText     =   "Adds New Record"
            Top             =   3165
            Width           =   1560
         End
         Begin MSComCtl2.DTPicker txtDate 
            Height          =   330
            Left            =   900
            TabIndex        =   1
            Top             =   585
            Width           =   1185
            _ExtentX        =   2090
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
            CustomFormat    =   "dd-MMM-yy "
            Format          =   43974659
            CurrentDate     =   38023
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cr Ac ID:"
            Height          =   270
            Index           =   1
            Left            =   60
            TabIndex        =   43
            Top             =   1410
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Narration:"
            Height          =   270
            Index           =   4
            Left            =   60
            TabIndex        =   42
            Top             =   2640
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   270
            Index           =   2
            Left            =   60
            TabIndex        =   41
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Serial:"
            Height          =   270
            Index           =   1
            Left            =   60
            TabIndex        =   40
            Top             =   195
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount:"
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   39
            Top             =   1830
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Dr Ac ID:"
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   38
            Top             =   1005
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3735
         Left            =   -74895
         TabIndex        =   36
         Top             =   360
         Width           =   7065
         Begin VSPrinter8LibCtl.VSPrinter vp 
            Height          =   3555
            Left            =   30
            TabIndex        =   27
            Top             =   135
            Width           =   6945
            _cx             =   12250
            _cy             =   6271
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            MousePointer    =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoRTF         =   -1  'True
            Preview         =   -1  'True
            DefaultDevice   =   0   'False
            PhysicalPage    =   -1  'True
            AbortWindow     =   -1  'True
            AbortWindowPos  =   0
            AbortCaption    =   "Printing..."
            AbortTextButton =   "Cancel"
            AbortTextDevice =   "on the %s on %s"
            AbortTextPage   =   "Now printing Page %d of"
            FileName        =   ""
            MarginLeft      =   1440
            MarginTop       =   1440
            MarginRight     =   1440
            MarginBottom    =   1440
            MarginHeader    =   0
            MarginFooter    =   0
            IndentLeft      =   0
            IndentRight     =   0
            IndentFirst     =   0
            IndentTab       =   720
            SpaceBefore     =   0
            SpaceAfter      =   0
            LineSpacing     =   100
            Columns         =   1
            ColumnSpacing   =   180
            ShowGuides      =   2
            LargeChangeHorz =   300
            LargeChangeVert =   300
            SmallChangeHorz =   30
            SmallChangeVert =   30
            Track           =   0   'False
            ProportionalBars=   -1  'True
            Zoom            =   17.3295454545455
            ZoomMode        =   3
            ZoomMax         =   400
            ZoomMin         =   10
            ZoomStep        =   25
            EmptyColor      =   -2147483636
            TextColor       =   0
            HdrColor        =   0
            BrushColor      =   0
            BrushStyle      =   0
            PenColor        =   0
            PenStyle        =   0
            PenWidth        =   0
            PageBorder      =   0
            Header          =   ""
            Footer          =   ""
            TableSep        =   "|;"
            TableBorder     =   7
            TablePen        =   0
            TablePenLR      =   0
            TablePenTB      =   0
            NavBar          =   3
            NavBarColor     =   -2147483633
            ExportFormat    =   0
            URL             =   ""
            Navigation      =   3
            NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
            AutoLinkNavigate=   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3705
         Left            =   -74895
         TabIndex        =   29
         Top             =   360
         Width           =   7035
         Begin VB.TextBox txtEditCrName 
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
            Height          =   315
            Left            =   1830
            TabIndex        =   19
            ToolTipText     =   "Name "
            Top             =   1410
            Width           =   4710
         End
         Begin VB.TextBox txtEditCrID 
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
            Height          =   315
            Left            =   900
            TabIndex        =   18
            ToolTipText     =   "Payee's ID"
            Top             =   1410
            Width           =   915
         End
         Begin VB.TextBox txtEditDrName 
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
            Height          =   315
            Left            =   1830
            TabIndex        =   17
            ToolTipText     =   "Name "
            Top             =   1005
            Width           =   4710
         End
         Begin VB.TextBox txtEditDrID 
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
            Height          =   315
            Left            =   900
            TabIndex        =   16
            ToolTipText     =   "Payee's ID"
            Top             =   1005
            Width           =   915
         End
         Begin VB.TextBox txtEditAmountWords 
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
            Height          =   315
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   21
            ToolTipText     =   "Amount in words"
            Top             =   1815
            Width           =   4665
         End
         Begin VB.TextBox txtEditSerial 
            Alignment       =   1  'Right Justify
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
            Left            =   900
            TabIndex        =   15
            ToolTipText     =   "Serial No."
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox txtEditNarration 
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
            Left            =   900
            TabIndex        =   22
            ToolTipText     =   "Narration"
            Top             =   2625
            Width           =   6045
         End
         Begin VB.TextBox txtEditAmount 
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
            Left            =   900
            TabIndex        =   20
            ToolTipText     =   "Amount"
            Top             =   1815
            Width           =   1350
         End
         Begin VB.CommandButton cmdEditUpdate 
            Caption         =   "&Update"
            Height          =   390
            Left            =   90
            TabIndex        =   23
            ToolTipText     =   "Adds New Record"
            Top             =   3165
            Width           =   1560
         End
         Begin VB.CommandButton cmdEditQuit 
            Caption         =   "&Quit"
            Height          =   390
            Left            =   5280
            TabIndex        =   26
            Top             =   3165
            Width           =   1560
         End
         Begin VB.CommandButton cmdSelectEditSerial 
            DownPicture     =   "frmTRN.frx":5B6E
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1815
            Picture         =   "frmTRN.frx":5EB1
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Open Serial"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdEditDelete 
            Caption         =   "&Delete"
            Height          =   390
            Left            =   1815
            TabIndex        =   24
            ToolTipText     =   "Adds New Record"
            Top             =   3165
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.CommandButton cmdEditPrint 
            Caption         =   "&Print"
            Height          =   390
            Left            =   3555
            TabIndex        =   25
            ToolTipText     =   "Adds New Record"
            Top             =   3165
            Width           =   1560
         End
         Begin MSComCtl2.DTPicker txtEditDate 
            Height          =   330
            Left            =   900
            TabIndex        =   31
            Top             =   585
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MMM-yy "
            Format          =   43974659
            CurrentDate     =   38023
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cr Ac ID:"
            Height          =   270
            Index           =   3
            Left            =   60
            TabIndex        =   45
            Top             =   1410
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Dr Ac ID:"
            Height          =   270
            Index           =   2
            Left            =   60
            TabIndex        =   44
            Top             =   1005
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount:"
            Height          =   270
            Index           =   5
            Left            =   60
            TabIndex        =   35
            Top             =   1830
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Serial:"
            Height          =   270
            Index           =   6
            Left            =   60
            TabIndex        =   34
            Top             =   195
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   270
            Index           =   7
            Left            =   60
            TabIndex        =   33
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Narration:"
            Height          =   270
            Index           =   9
            Left            =   60
            TabIndex        =   32
            Top             =   2640
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmTRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    txtDate.Value = Now
    txtEditSerial.Text = 1
    SSTab1.Tab = 0
    SetFont ("[frmPrintReceipt-Font]")
End Sub

'==== TAB ADD

Private Sub txtDrID_Change()
    sSQL(0) = "SELECT Name FROM appview_AcHeads WHERE ID=" & Chr(39) & txtDrID.Text & Chr(39)
    dbOpen (0)
    If recs(0).RecordCount = 1 Then
        txtDrName.Text = recs(0)!Name
    Else
        txtDrName.Text = ""
    End If
    dbClose (0)
End Sub

Private Sub txtDrID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCrID.SetFocus
End Sub

Private Sub txtCrID_Change()
    sSQL(0) = "SELECT Name FROM appview_AcHeads WHERE ID=" & Chr(39) & txtCrID.Text & Chr(39)
    dbOpen (0)
    If recs(0).RecordCount = 1 Then
        txtCrName.Text = recs(0)!Name
    Else
        txtCrName.Text = ""
    End If
    dbClose (0)
End Sub
Private Sub txtCrID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtAmount.SetFocus
End Sub

Private Sub cmdSelectDrID_Click()
    sSQL(0) = "SELECT * FROM appview_AcHeads"
    frmShow.Show vbModal: sSQL(0) = ""
    If sArray(0) <> "" Then
        txtDrID.Text = sArray(0): txtDrName.Text = sArray(1)
    End If
End Sub

Private Sub cmdSelectCrID_Click()
    sSQL(0) = "SELECT * FROM appview_AcHeads"
    frmShow.Show vbModal: sSQL(0) = ""
    If sArray(0) <> "" Then
        txtCrID.Text = sArray(0): txtCrName.Text = sArray(1)
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtNarration.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.Text = Format(Val(txtAmount.Text), "##0.00")
    txtAmountWords.Text = ConvertCurrencyToEnglish(Val(txtAmount.Text))
End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdAdd.SetFocus
End Sub

Private Sub cmdAdd_Click()
    Dim jRef As Long
    Dim DrAc, CrAc, Narr As String
    If MsgBox("Are you sure to add new transaction?", vbYesNo) = vbYes Then
        If txtDrID.Text <> "" And txtCrID.Text <> "" Then
            sSQL(0) = "INSERT INTO TRN (DATE, DrID, DrNAME, CrID, CrName, AMOUNT, NARRATION) VALUES ("
            sSQL(0) = sSQL(0) & Chr(39) & Format(txtDate.Value, "MM-dd-yyyy") & Chr(39) & ", "
            sSQL(0) = sSQL(0) & Chr(39) & txtDrID.Text & Chr(39) & ", "
            sSQL(0) = sSQL(0) & Chr(39) & txtDrName.Text & Chr(39) & ", "
            sSQL(0) = sSQL(0) & Chr(39) & txtCrID.Text & Chr(39) & ", "
            sSQL(0) = sSQL(0) & Chr(39) & txtCrName.Text & Chr(39) & ", "
            sSQL(0) = sSQL(0) & Val(txtAmount.Text) & ", "
            sSQL(0) = sSQL(0) & Chr(39) & txtNarration.Text & Chr(39) & ")"
            dbOpen (0): dbClose (0)
            sSQL(0) = "SELECT TOP 1 * FROM TRN ORDER BY 1 DESC"
            dbOpen (0)
                txtSerial.Text = recs(0)!Serial
                txtDate.Value = recs(0)!Date
                txtDrID.Text = recs(0)!DrID
                txtDrName.Text = recs(0)!DrName
                txtCrID.Text = recs(0)!CrID
                txtCrName.Text = recs(0)!CrName
                txtAmount.Text = recs(0)!Amount
                txtAmountWords.Text = ConvertCurrencyToEnglish(recs(0)!Amount)
                txtNarration.Text = recs(0)!Narration
            dbClose (0)
            jRef = MakeJournalEntry(txtDate.Value, txtDrID.Text, txtCrID.Text, Val(txtAmount.Text), txtNarration.Text & " vide TRN#" & txtSerial.Text, 0, Format(Val(txtSerial.Text), "\T\R\N0"), True)
            sSQL(0) = "UPDATE TRN SET JREF=" & jRef & " WHERE SERIAL=" & Val(txtSerial.Text): dbOpen (0): dbClose (0)
        Else
            MsgBox "ID error!", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub cmdClear_Click()
    txtSerial.Text = "": txtDrID.Text = "": txtDrName.Text = "": txtCrID.Text = "": txtCrName.Text = "": txtAmount.Text = 0: txtNarration.Text = ""
    txtAmount_LostFocus
End Sub

Private Sub cmdPrint_Click()
    If Val(txtSerial.Text) <> 0 Then
        msg = "   " & vbCrLf
        msg = msg & CompanyName & ": TRANSACTION SLIP" & vbCrLf & vbCrLf
        msg = msg & "TRN#" & txtSerial.Text & "              " & "Date:" & txtDate.Value & vbCrLf
        msg = msg & "Dr Account: " & txtDrID.Text & " - " & txtDrName.Text & vbCrLf
        msg = msg & "Cr Account: " & txtCrID.Text & " - " & txtCrName.Text & vbCrLf
        msg = msg & "Transacted Amount Rs." & Format(Val(txtAmount.Text), "#,##0.00") & " (" & txtAmountWords.Text & " )" & vbCrLf
        msg = msg & vbCrLf
        msg = msg & "Current A/c Bal of " & txtDrID.Text & " - " & txtDrName.Text & " is Rs." & Format(WhatIsLedgerBalance(txtDrID.Text), "#,##0.00") & vbCrLf
        msg = msg & "Current A/c Bal of " & txtCrID.Text & " - " & txtCrName.Text & " is Rs." & Format(WhatIsLedgerBalance(txtCrID.Text), "#,##0.00") & vbCrLf
    End If
    
    
    With vp
        .StartDoc
        .StartTable
        .TableBorder = tbAll
        .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
        .TableCell(tcColWidth, , 1) = "7.50in"
        .TableCell(tcColAlign, , 1) = taLeftTop
        .TableCell(tcText, 1, 1) = msg
        .EndTable
        .EndDoc
    End With
    SSTab1.Tab = 2
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub


'=== TAB EDIT

Private Sub txtEditSerial_LostFocus()
    txtEditSerial.Text = Val(txtEditSerial.Text)
    sSQL(0) = "SELECT * FROM TRN WHERE SERIAL=" & Val(txtEditSerial.Text)
    dbOpen (0)
    If recs(0).RecordCount = 1 Then
            txtEditSerial.Text = recs(0)!Serial
            txtEditDate.Value = recs(0)!Date
            txtEditDrID.Text = recs(0)!DrID
            txtEditDrName.Text = recs(0)!DrName
            txtEditCrID.Text = recs(0)!CrID
            txtEditCrName.Text = recs(0)!CrName
            txtEditAmount.Text = recs(0)!Amount
            txtEditAmountWords.Text = ConvertCurrencyToEnglish(recs(0)!Amount)
            txtEditNarration.Text = recs(0)!Narration
    Else
        MsgBox "No such transaction! Check again.", vbOKOnly + vbCritical
        txtEditSerial.Text = "1": txtEditSerial.SetFocus
    End If
End Sub

Private Sub cmdSelectEditSerial_Click()
    sSQL(0) = "SELECT * FROM PMT ORDER BY 1 DESC"
    frmShow.Show vbModal: sSQL(0) = ""
    If sArray(0) <> "" Then
        txtEditSerial.Text = sArray(0)
    End If
    txtEditSerial_LostFocus
End Sub

Private Sub txtEditAmount_LostFocus()
    txtEditAmount.Text = Format(Val(txtEditAmount.Text), "##0.00")
    txtEditAmountWords.Text = ConvertCurrencyToEnglish(Val(txtEditAmount.Text))
End Sub

Private Sub cmdEditUpdate_Click()
    Dim jRef As Long
    Dim CrAc As String
    If MsgBox("Only changes to the amount & narration is permissible, Continue?", vbYesNo) = vbYes Then
        If txtEditDrID.Text <> "" And txtEditCrID.Text <> "" Then
            sSQL(0) = "UPDATE TRN SET "
            sSQL(0) = sSQL(0) & " AMOUNT=" & Val(txtEditAmount.Text) & ", "
            sSQL(0) = sSQL(0) & " NARRATION=" & Chr(39) & txtEditNarration.Text & Chr(39)
            sSQL(0) = sSQL(0) & " WHERE SERIAL=" & Val(txtEditSerial.Text)
            dbOpen (0): dbClose (0)
            sSQL(0) = "SELECT * FROM TRN WHERE SERIAL=" & Val(txtEditSerial.Text)
            dbOpen (0)
                txtEditSerial.Text = recs(0)!Serial
                txtEditDate.Value = recs(0)!Date
                txtEditDrID.Text = recs(0)!DrID
                txtEditDrName.Text = recs(0)!DrName
                txtEditCrID.Text = recs(0)!CrID
                txtEditCrName.Text = recs(0)!CrName
                txtEditAmount.Text = recs(0)!Amount
                txtEditAmountWords.Text = ConvertCurrencyToEnglish(recs(0)!Amount)
                txtEditNarration.Text = recs(0)!Narration
                jRef = recs(0)!jRef
            dbClose (0)
            MakeJournalEntry txtEditDate.Value, txtEditDrID.Text, txtEditCrID.Text, Val(txtEditAmount.Text), txtEditNarration.Text & " VIDE TRN#" & txtSerial.Text, jRef, Format(Val(txtSerial.Text), "\T\R\N0"), True
        Else
            MsgBox "ID error!", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub cmdEditPrint_Click()
    If Val(txtEditSerial.Text) <> 0 Then
        msg = "   " & vbCrLf
        msg = msg & CompanyName & ": TRANSACTION SLIP" & vbCrLf & vbCrLf
        msg = msg & "TRN#" & txtEditSerial.Text & "              " & "Date:" & txtEditDate.Value & vbCrLf
        msg = msg & "Dr Account: " & txtEditDrID.Text & " - " & txtEditDrName.Text & vbCrLf
        msg = msg & "Cr Account: " & txtEditCrID.Text & " - " & txtEditCrName.Text & vbCrLf
        msg = msg & "Transacted Amount Rs." & Format(Val(txtEditAmount.Text), "#,##0.00") & " (" & txtEditAmountWords.Text & " )" & vbCrLf
        msg = msg & vbCrLf
        msg = msg & "Current A/c Bal of " & txtEditDrID.Text & " - " & txtEditDrName.Text & " is Rs." & Format(WhatIsLedgerBalance(txtEditDrID.Text), "#,##0.00") & vbCrLf
        msg = msg & "Current A/c Bal of " & txtEditCrID.Text & " - " & txtEditCrName.Text & " is Rs." & Format(WhatIsLedgerBalance(txtEditCrID.Text), "#,##0.00") & vbCrLf
    End If
    
    
    With vp
        .StartDoc
        .StartTable
        .TableBorder = tbAll
        .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
        .TableCell(tcColWidth, , 1) = "7.50in"
        .TableCell(tcColAlign, , 1) = taLeftTop
        .TableCell(tcText, 1, 1) = msg
        .EndTable
        .EndDoc
    End With
    SSTab1.Tab = 2
End Sub

Private Sub cmdEditQuit_Click()
    Unload Me
End Sub



'=== SET FONT ROUTINES

Private Sub SetFont(S As String)
    vp.PaperSize = pprEnvB4
    vp.MarginLeft = GReadINI("[MarginLeft]"): vp.MarginRight = GReadINI("[MarginRight]"): vp.MarginTop = GReadINI("[MarginTop]"): vp.MarginBottom = GReadINI("[MarginBottom]")
    vp.PenStyle = psSolid: vp.TrueType = ttBitmap
    vp.FontName = ReadFont(S, 0)
    vp.FontSize = ReadFont(S, 1)
    vp.FontBold = ReadFont(S, 2)
    vp.FontItalic = ReadFont(S, 3)
End Sub
