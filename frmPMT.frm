VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPMT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payments..."
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPMT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   7408
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Add/Update..."
      TabPicture(0)   =   "frmPMT.frx":114DA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Print..."
      TabPicture(1)   =   "frmPMT.frx":114F6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Height          =   3765
         Left            =   -74925
         TabIndex        =   29
         Top             =   360
         Width           =   7515
         Begin VSPrinter8LibCtl.VSPrinter vp 
            Height          =   3555
            Left            =   90
            TabIndex        =   30
            Top             =   165
            Width           =   7335
            _cx             =   12938
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Height          =   3765
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   7500
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
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
            Left            =   1515
            TabIndex        =   20
            ToolTipText     =   "Adds New Record"
            Top             =   3270
            Width           =   1470
         End
         Begin VB.CommandButton cmdSelectSerial 
            DownPicture     =   "frmPMT.frx":11512
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
            Picture         =   "frmPMT.frx":11855
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Open Serial"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
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
            Left            =   2985
            TabIndex        =   18
            ToolTipText     =   "Adds New Record"
            Top             =   3270
            Width           =   1470
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
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Amount in words"
            Top             =   1815
            Width           =   5130
         End
         Begin VB.CommandButton cmdSelectID 
            DownPicture     =   "frmPMT.frx":11B98
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
            Left            =   1830
            Picture         =   "frmPMT.frx":11EDB
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Open Payee"
            Top             =   1005
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.TextBox txtID 
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
            ToolTipText     =   "Payee's ID"
            Top             =   1005
            Width           =   915
         End
         Begin VB.TextBox txtName 
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
            TabIndex        =   14
            ToolTipText     =   "Name "
            Top             =   1410
            Width           =   6510
         End
         Begin VB.TextBox txtSerial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
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
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Serial No."
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox txtMode 
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
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Mode of Payment"
            Top             =   2220
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
            TabIndex        =   11
            ToolTipText     =   "Narration"
            Top             =   2625
            Width           =   6510
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
            TabIndex        =   10
            ToolTipText     =   "Amount"
            Top             =   1815
            Width           =   1350
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "&New"
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
            Left            =   45
            TabIndex        =   9
            ToolTipText     =   "Adds New Record"
            Top             =   3270
            Width           =   1470
         End
         Begin VB.CommandButton cmdQuit 
            Caption         =   "&Quit"
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
            Left            =   5925
            TabIndex        =   8
            Top             =   3270
            Width           =   1470
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
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
            Left            =   4455
            TabIndex        =   7
            ToolTipText     =   "Adds New Record"
            Top             =   3270
            Width           =   1470
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2310
            TabIndex        =   2
            Top             =   2055
            Width           =   5070
            Begin VB.OptionButton opMode 
               Caption         =   "&OTHER"
               Height          =   195
               Index           =   3
               Left            =   3855
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   240
               Width           =   870
            End
            Begin VB.OptionButton opMode 
               Caption         =   "&DD"
               Height          =   195
               Index           =   2
               Left            =   2700
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   240
               Width           =   915
            End
            Begin VB.OptionButton opMode 
               Caption         =   "CHE&QUE"
               Height          =   195
               Index           =   1
               Left            =   1335
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   240
               Width           =   1020
            End
            Begin VB.OptionButton opMode 
               Caption         =   "&CASH"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   3
               Top             =   240
               Width           =   1020
            End
         End
         Begin MSComCtl2.DTPicker txtDate 
            Height          =   330
            Left            =   900
            TabIndex        =   21
            Top             =   585
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   19267587
            CurrentDate     =   38023
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ID:"
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   28
            Top             =   1005
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount:"
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   1830
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Serial:"
            Height          =   270
            Index           =   1
            Left            =   60
            TabIndex        =   26
            Top             =   195
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   270
            Index           =   2
            Left            =   60
            TabIndex        =   25
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Name:"
            Height          =   270
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   1410
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Mode:"
            Height          =   270
            Index           =   3
            Left            =   60
            TabIndex        =   23
            Top             =   2235
            Width           =   810
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Narration:"
            Height          =   270
            Index           =   4
            Left            =   60
            TabIndex        =   22
            Top             =   2640
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmPMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MemoFormat = "\P\T0"
Private Const Table = "PMT"
Private Const NewProc = "ADDNEW_PMT"
Private Const SaveProc = "UPDATE_PMT"

Private Sub Form_Load()
    Me.Move 0, 0
    txtDate.Value = Now
    txtSerial.Text = 1
    opMode(0).Value = True
    SSTab1.Tab = 0
    SetFont ("[frmPrintPayment-Font]")
End Sub


'==== TAB ADD
Private Sub txtSerial_LostFocus()
    txtSerial.Text = Val(txtSerial.Text)
    sSQL(0) = "SELECT * FROM " & Table & " WHERE SERIAL=" & Val(txtSerial.Text)
    dbOpen (0)
    If recs(0).RecordCount = 1 Then
        txtDate.Value = recs(0)!Date
        txtID.Text = recs(0)!ID
        txtName.Text = recs(0)!Name
        txtAmount.Text = recs(0)!Amount
        txtAmountWords.Text = ConvertCurrencyToEnglish(recs(0)!Amount)
        txtMode.Text = recs(0)!Mode
        txtNarration.Text = recs(0)!Narration
        txtAmount.SetFocus
    Else
        MsgBox "No such record! Check again.", vbOKOnly + vbCritical
        cmdClear_Click
    End If
End Sub

Private Sub cmdSelectSerial_Click()
    sSQL(0) = "SELECT * FROM " & Table & " ORDER BY 1 DESC"
    frmShow.Init sSQL(0): sSQL(0) = ""
    If sArray(0) <> "" Then
        txtSerial.Text = sArray(0)
    End If
    txtSerial_LostFocus
End Sub

Private Sub txtID_Change()
    sSQL(5) = "SELECT Name FROM appview_AllAccounts WHERE ID=" & Chr(39) & txtID.Text & Chr(39)
    dbOpen (5)
    If recs(5).RecordCount = 1 Then
        txtName.Text = recs(5)!Name
    Else
        txtName.Text = ""
    End If
    dbClose (5)
End Sub

Private Sub txtID_LostFocus()
    sSQL(0) = "SELECT Name FROM appview_AllAccounts WHERE ID=" & Chr(39) & txtID.Text & Chr(39)
    dbOpen (0)
    If recs(0).RecordCount = 1 Then
        txtName.Text = recs(0)!Name
    Else
        txtName.Text = "": txtID.Text = ""
    End If
    dbClose (0)
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtAmount.SetFocus
End Sub

Private Sub cmdSelectID_Click()
    sSQL(0) = "SELECT * FROM appview_AllAccounts"
    frmShow.Init sSQL(0): sSQL(0) = ""
    If sArray(0) <> "" Then
        txtID.Text = sArray(0): txtName.Text = sArray(1)
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

Private Sub opMode_Click(Index As Integer)
    txtMode.Text = Replace(opMode(Index).Caption, "&", "")
End Sub

Private Sub cmdNew_Click()
    Dim jRef As Long
    Dim CrAc, Narr As String
    If MsgBox("Are you sure to add new?", vbYesNo) = vbYes Then
            sSQL(0) = NewProc
            dbOpen (0)
            Set recs(0) = recs(0).NextRecordset
            txtSerial.Text = recs(0)!MaxSerial
            dbClose (0)
            txtSerial_LostFocus
    End If
    txtID.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim jRef As Long
    Dim CrAc As String
    If HasAccount(txtID.Text) Then
        sSQL(0) = SaveProc & " " & Val(txtSerial.Text) & ", " & QT(Format(txtDate.Value, "dd-MMM-yy")) & ", " & QT(txtID.Text) & ", " & QT(txtName.Text) & ", " & Val(txtAmount.Text) & ", " & QT(UCase(txtMode.Text)) & ", " & QT(txtNarration.Text)
        dbOpen (0): dbClose (0)
        MakeSingleJournalEntry Table, Val(txtSerial.Text), MemoFormat
    Else
        MsgBox "ID Error!", vbOKOnly + vbCritical
    End If
    txtSerial_LostFocus
End Sub

Private Sub cmdClear_Click()
    txtSerial.Text = "": txtID.Text = "": txtName.Text = "": txtAmount.Text = 0: txtMode.Text = "CASH": txtNarration.Text = ""
    txtAmount_LostFocus
End Sub

Private Sub cmdPrint_Click()
    If Val(txtSerial.Text) <> 0 And Table = "RCT" Then
    With vp
        If UCase(txtMode.Text) = "CASH" Then
            transactionMode = "CASH"
        Else
            transactionMode = txtMode.Text
        End If
        headString = "{\f0\fs16 " & txtMode.Text & " RECEIPT\par \b\fs22 " & CompanyName & " " & AboutCompany & " \par \b0\fs20 " & CompanyAddress & " \par }"
        bodyString = "{\b0\f0\fs18 No.: " & Format(Val(txtSerial.Text), "0000") & " (" & txtID.Text & ")\tab\tab DATE: " & Format(txtDate.Value, "dd-MM-yy HH:MM") & "\par\par Received with thanks from \b " & txtName.Text & " \b0 \i0\fs20 the sum of \b Rs." & Format(Val(txtAmount.Text), "#,#00.00") & " (" & txtAmountWords.Text & ") \b0 \fs18 by " & txtMode.Text & " on account of _________________________. \par\par \i Your ledger balance after this transaction is \b Rs." & Format(Val(WhatIsLedgerBalance(txtID.Text)), "#,#00.00(Dr); #,#00.00(Cr)") & " \b0\i0\par Carried by: Self \par Comments: \par \pard\qr\i \par\par Authorised Signatory\i0\par \f1}"
        
        .PaperSize = pprA4
        .TextAlign = taCenterTop
        .MarginLeft = 400
        .FontName = "Tahoma"
        .StartDoc
        .PenWidth = 40
        '.DrawPicture mdiOne.ImgList.ListImages(1).Picture, 2000, 450
        .StartTable
        .TableBorder = tbAll
        .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
        .TableCell(tcColWidth, , 1) = "6.00in"
        .TableCell(tcAlign, 1, 1, 1, 1) = taCenterMiddle
        .TableCell(tcText, 1, 1) = headString
        .TableCell(tcAlign, 2, 1, 2, 1) = taJustMiddle
        .TableCell(tcText, 2, 1) = bodyString
        .EndTable
        .EndDoc
        .Zoom = 76
    End With
    SSTab1.Tab = 1
    End If
End Sub

Private Sub cmdQuit_Click()
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





