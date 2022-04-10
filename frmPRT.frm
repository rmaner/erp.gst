VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPRT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payments/Receipts/Transfers..."
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPRT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   4395
      Left            =   6285
      TabIndex        =   29
      Top             =   0
      Width           =   6225
      Begin VSPrinter8LibCtl.VSPrinter vp 
         Height          =   4170
         Left            =   90
         TabIndex        =   30
         Top             =   165
         Width           =   6015
         _cx             =   10610
         _cy             =   7355
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
         Zoom            =   21.2121212121212
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
      Height          =   4395
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6225
      Begin VB.TextBox txtCity 
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
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Name "
         Top             =   1500
         Width           =   1740
      End
      Begin VB.CommandButton cmdPaymentForwarding 
         Caption         =   "&Forwarding"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4530
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   150
         Width           =   1560
      End
      Begin VB.CommandButton cmdSelectMode 
         DownPicture     =   "frmPRT.frx":114DA
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
         Left            =   5685
         Picture         =   "frmPRT.frx":1181D
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Open Payee"
         Top             =   2280
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox txtModeName 
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Amount in words"
         Top             =   2295
         Width           =   3705
      End
      Begin VB.TextBox txtMode 
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
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Amount"
         Top             =   2295
         Width           =   1035
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
         Left            =   4875
         TabIndex        =   17
         Top             =   3870
         Width           =   1215
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
         Left            =   3675
         TabIndex        =   16
         ToolTipText     =   "Adds New Record"
         Top             =   3870
         Width           =   1215
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
         Left            =   2475
         TabIndex        =   15
         ToolTipText     =   "Adds New Record"
         Top             =   3870
         Width           =   1215
      End
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
         Left            =   1275
         TabIndex        =   14
         ToolTipText     =   "Adds New Record"
         Top             =   3870
         Width           =   1215
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
         Left            =   75
         TabIndex        =   13
         ToolTipText     =   "Adds New Record"
         Top             =   3870
         Width           =   1215
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
         TabIndex        =   7
         ToolTipText     =   "Amount"
         Top             =   1890
         Width           =   1035
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
         TabIndex        =   12
         ToolTipText     =   "Narration"
         Top             =   2715
         Width           =   5205
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Serial No."
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox txtName 
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
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Name "
         Top             =   1500
         Width           =   3450
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
         TabIndex        =   3
         ToolTipText     =   "Payee's ID"
         Top             =   1095
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectID 
         DownPicture     =   "frmPRT.frx":11B60
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
         Left            =   2130
         Picture         =   "frmPRT.frx":11EA3
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Open Payee"
         Top             =   1095
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Amount in words"
         Top             =   1890
         Width           =   4155
      End
      Begin VB.CommandButton cmdSelectSerial 
         DownPicture     =   "frmPRT.frx":121E6
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
         Picture         =   "frmPRT.frx":12529
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Open Serial"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.Frame Frame4 
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
         Left            =   3495
         TabIndex        =   20
         Top             =   750
         Visible         =   0   'False
         Width           =   2610
         Begin MSMask.MaskEdBox txtAccountBalance 
            Height          =   300
            Left            =   75
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   210
            Width           =   2445
            _ExtentX        =   4313
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
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   330
         Left            =   900
         TabIndex        =   2
         Top             =   705
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   80019459
         CurrentDate     =   38023
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Narration:"
         Height          =   270
         Index           =   4
         Left            =   60
         TabIndex        =   28
         Top             =   2730
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Mode:"
         Height          =   270
         Index           =   3
         Left            =   60
         TabIndex        =   27
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   270
         Index           =   1
         Left            =   60
         TabIndex        =   26
         Top             =   1530
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   270
         Index           =   2
         Left            =   60
         TabIndex        =   25
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial:"
         Height          =   270
         Index           =   1
         Left            =   60
         TabIndex        =   24
         Top             =   195
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   23
         Top             =   1905
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   1095
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmPRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyCAPTION As String
Public FRAMECOLOR As Double
Public MemoFormat As String
Public ADDPROC As String
Public UPDATEPROC As String
Public Support As String
Private CN(5) As New clsData
Private cJr As New clsJournals

Private Sub Form_DblClick()
    For i = 400 To 433
'        txtSerial.Text = i: txtSerial_LostFocus: cmdSave_Click
    Next
End Sub

Private Sub Form_Load()
    Me.Move 0, 0
    Me.Caption = MyCAPTION
    txtDate.Value = Now
    mdiOne.SetFormFont Me
    SetVPFont ("[PRT-VP-FONT]")
    Frame2.BackColor = FRAMECOLOR
    vp.Zoom = 88
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageUp Then cmdSelectID_Click
    If KeyCode = vbKeyPageDown Then cmdSelectMode_Click
    If KeyCode = vbKeyF12 Then cmdSave_Click
    If KeyCode = vbKeyEscape Then
        cmdSave.Caption = "&Save"
        txtSerial.Visible = True
    End If
    If (KeyCode = 98 Or KeyCode = 50) And Shift = vbCtrlMask Then DUPLICATETHIS     'TWO
End Sub

'==== TAB ADD
Private Sub txtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtDate.SetFocus
    If KeyCode = vbKeyUp Then
        txtSerial.Text = Val(txtSerial.Text) + 1: txtSerial_LostFocus
    End If
    If KeyCode = vbKeyDown Then
        txtSerial.Text = Val(txtSerial.Text) - 1: txtSerial_LostFocus
    End If
End Sub

Private Sub txtSerial_LostFocus()
    CN(0).dbOpen "SELECT ID, NAME FROM PERSONAL WHERE TYPE=" & QT("CASH") & " OR TYPE=" & QT("BANK")
    txtSerial.Text = Val(txtSerial.Text)
    CN(0).dbOpen "SELECT * FROM " & Support & " WHERE SERIAL=" & Val(txtSerial.Text)
    If CN(0).recs.RecordCount = 1 Then
        txtDate.Value = CN(0).recs!Date
        txtID.Text = CN(0).recs!id: txtName.Text = CN(0).recs!Name: txtCity.Text = CN(0).recs!City
        txtAmount.Text = CN(0).recs!Amount
        txtAmountWords.Text = ConvertCurrencyToEnglish(CN(0).recs!Amount)
        txtMode.Text = CN(0).recs!Mode: txtModeName.Text = CN(0).recs!ModeName
        txtNarration.Text = CN(0).recs!Narration
    Else
        MsgBox "No such record! Check again.", vbOKOnly + vbCritical
        cmdClear_Click
    End If
    Preview
End Sub

Private Sub cmdSelectSerial_Click()
    frmShow.Init "SELECT * FROM  " & Support & " ORDER BY 1 DESC"
    If sArray(0) <> "" Then
        txtSerial.Text = sArray(0)
    End If
    txtSerial_LostFocus
    txtAccountBalance.Text = Val(WhatIsLedgerBalance(txtID.Text))
End Sub

Private Sub txtID_Change()
    CN(1).dbOpen "SELECT * FROM appview_ALLACCOUNTS WHERE ID=" & Chr(39) & txtID.Text & Chr(39)
    If CN(1).recs.RecordCount = 1 Then
        txtName.Text = CN(1).recs!Name: txtCity.Text = CN(1).recs!City
    Else
       txtName.Text = "_": txtCity.Text = "_"
    End If
    txtAccountBalance.Text = Val(WhatIsLedgerBalance(txtID.Text))
End Sub

Private Sub cmdSelectID_Click()
    frmShow.Init "SELECT * FROM appview_ALLACCOUNTS ORDER BY NAME"
    If sArray(0) <> "" Then
        txtID.Text = sArray(0)
    End If
    txtID.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.Text = Format(Val(txtAmount.Text), "##0.00")
    txtAmountWords.Text = ConvertCurrencyToEnglish(Val(txtAmount.Text))
End Sub

Private Sub cmdSelectMode_Click()
    frmShow.Init "SELECT ID, NAME FROM appview_ALLACCOUNTS WHERE TYPE=" & QT("CASH") & " OR TYPE=" & QT("BANK") & " OR TYPE=" & QT("BRANCH")
    If sArray(0) <> "" Then
        txtMode.Text = sArray(0)
        txtModeName.Text = sArray(1)
    End If
    txtNarration.SetFocus
End Sub

Private Sub cmdNew_Click()
    On Error Resume Next
    If MsgBox("Are you sure to add?", vbYesNo) = vbYes Then
        CN(0).dbOpen ADDPROC
        Set CN(0).recs = CN(0).recs.NextRecordset
        txtSerial.Text = CN(0).recs!MaxSerial
    End If
    txtSerial_LostFocus
    txtDate.SetFocus
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next

    If MsgBox("Are you sure to update?", vbYesNo) = vbYes Then
        If UCase(cmdSave.Caption) = "DUPLICATE" Then
            CN(0).dbOpen ADDPROC: Set CN(0).recs = CN(0).recs.NextRecordset
            txtSerial.Visible = True
            txtSerial.Text = CN(0).recs!MaxSerial: cmdSave.Caption = "&Save"
        End If
        If Val(txtSerial.Text) <> 0 And txtID.Text <> "" Then
            CN(2).dbOpen UPDATEPROC & Val(txtSerial.Text) & ", " & QT(Format(txtDate.Value, "dd-MMM-yy HH:MM")) & ", " & QT(UCase(txtID.Text)) & ", " & QT(txtName.Text) & ", " & QT(txtCity.Text) & ", " & Val(txtAmount.Text) & ", " & QT(txtMode.Text) & ", " & QT(txtModeName.Text) & ", " & QT(txtNarration.Text)
            cJr.PRTJournalEntries Support, Val(txtSerial.Text), MemoFormat
        Else
            MsgBox "ID or MODE error!", vbOKOnly + vbCritical
        End If
    End If
    txtSerial.SetFocus
    txtSerial_LostFocus
    cmdAdd.SetFocus
End Sub

Private Sub cmdClear_Click()
    txtSerial.Text = ""
    txtID.Text = ""
    txtName.Text = ""
    txtAmount.Text = 0
    txtMode.Text = "R0002": txtModeName.Text = "CASH"
    txtNarration.Text = ""
    txtAmount_LostFocus
End Sub

Private Sub cmdPrint_Click()
    Preview
    vp.PrintDoc
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

'=== SET FONT ROUTINES
Private Sub SetVPFont(S As String)
    PprDim = Split(mdiOne.sckGo.GReadINI("[PaperHeight/Width/Left/Right/Top/Bottom]"), ":")
    vp.PaperSize = pprUser
    vp.PaperHeight = PprDim(0): vp.PaperWidth = PprDim(1)
    vp.MarginLeft = PprDim(2): vp.MarginRight = PprDim(3)
    vp.MarginTop = PprDim(4): vp.MarginBottom = PprDim(5)
    
    vp.PenStyle = psSolid: vp.TrueType = ttBitmap
    vp.FontName = ReadFont(S, 0)
    vp.FontSize = ReadFont(S, 1)
    vp.FontBold = ReadFont(S, 2)
    vp.FontItalic = ReadFont(S, 3)
End Sub

Private Sub Preview()
    If Support = "RCT" Then
        RCTPrint
    Else
        PMTPrint
    End If
End Sub

Private Sub RCTPrint()
    If Val(txtSerial.Text) <> 0 Then
    With vp
        .Zoom = 90
        headString = "{\f0\fs16 RECEIPT\par \b\fs22 " & CompanyName & "\par \b0\fs20 " & CompanyAddr0 & " \par " & CompanyPhone & "; " & CompanyMobile & "\par}"
        bodyString = "{\b0\f0\fs18 No.: " & Format(Val(txtSerial.Text), "0000") & " (" & txtID.Text & "/" & ")\tab\tab DATE: " & Format(txtDate.Value, "dd-MM-yy HH:MM:SS") & "\par\par Received with thanks from \b " & txtName.Text & " \b0 \i0\fs20 the sum of \b Rs." & Format(Val(txtAmount.Text), "#,#00.00") & " (" & txtAmountWords.Text & ") \b0 \fs18 by " & txtNarration.Text & "\par\par \i Your ledger balance after this transaction is \b Rs." & Format(Val(WhatIsLedgerBalance(txtID.Text)), "#,#00.00") & " \par \pard\qr\i \par\par Authorised Signatory\i0\par \f1}"
        
        .PaperSize = pprA4
        .MarginLeft = 260
        .FontName = "Tahoma"
        .StartDoc
        .PenWidth = 40
'        .DrawPicture mdiOne.ImgList.ListImages(1).Picture, 450, 500
        .StartTable
        .TableBorder = tbAll
        .TableCell(tcCols) = 1: .TableCell(tcRows) = 2
        .TableCell(tcColWidth, , 1) = "3.70in"
        .TableCell(tcAlign, 1, 1, 1, 1) = taCenterTop
        .TableCell(tcText, 1, 1) = headString
        .TableCell(tcAlign, 2, 1, 2, 1) = taJustMiddle
        .TableCell(tcText, 2, 1) = bodyString
        .EndTable
        .EndDoc
    End With
    End If
End Sub

Private Sub PMTPrint()
    If Val(txtSerial.Text) <> 0 Then
    With vp
        headString = "{\f0\fs16 PAYMENT VOUCHER\par \b\fs22 " & CompanyName & "\par \b0\fs20 " & CompanyAddr0 & " \par " & CompanyPhone & "; " & CompanyMobile & "\par}"
        bodyString = "{\b0\f0\fs20 No.:" & Format(Val(txtSerial.Text), "####") & " (" & txtID.Text & ")\tab\tab DATE: " & Format(txtDate.Value, "dd-MM-yy HH:MM:SS") & "\par\par Paid to \b " & txtName.Text & " \b0 \i \i0\fs22 the sum of \b\i Rs." & Format(Val(txtAmount.Text), "#,#00.00") & "(" & txtAmountWords.Text & ") \b0\i0 \fs20 by " & txtNarration.Text & " \par \i Your ledger balance after this transaction is \b Rs." & Format(Val(WhatIsLedgerBalance(txtID.Text)), "#,#00.00") & " \i\par\par Authorised Signatory\i0\par \pard\f1}"
        
        .PaperSize = pprA4
        .MarginLeft = 260
        .FontName = "Tahoma"
        .StartDoc
        .PenWidth = 40
 '       .DrawPicture mdiOne.ImgList.ListImages(1).Picture, 450, 500
 '       .DrawPicture mdiOne.ImgList.ListImages(1).Picture, 6200, 500
        .StartTable
        .TableBorder = tbAll
        .TableCell(tcCols) = 3: .TableCell(tcRows) = 2
        .TableCell(tcFontName, 1, 1, 1, 3) = "Tahoma"
        .TableCell(tcFontName, 2, 1, 2, 3) = "Times New Roman"
        
        .TableCell(tcColWidth, , 1) = "3.70in": .TableCell(tcColWidth, , 2) = "0.30in": .TableCell(tcColWidth, , 3) = "3.70in"
        .TableCell(tcAlign, 1, 1, 1, 1) = taCenterTop
        .TableCell(tcText, 1, 1) = headString
        .TableCell(tcAlign, 2, 1, 2, 1) = taJustMiddle
        .TableCell(tcText, 2, 1) = bodyString
        
        .TableCell(tcAlign, 1, 3, 1, 3) = taCenterTop
        .TableCell(tcText, 1, 3) = headString
        .TableCell(tcAlign, 2, 3, 2, 3) = taJustMiddle
        .TableCell(tcText, 2, 3) = bodyString
        .EndTable
        .EndDoc
    End With
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    txtID_Change
    If KeyAscii = vbKeyReturn Then txtAmount.SetFocus
End Sub

Private Sub txtAmount_GotFocus()
    txtAmount.SelStart = 0: txtAmount.SelLength = Len(txtAmount.Text)
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtNarration.SetFocus
End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdSave.SetFocus
End Sub

Public Sub LinkOpen()
    Me.SetFocus
    txtSerial.Text = XREF: txtSerial_LostFocus
End Sub


Private Sub txtID_DblClick()
    XID = txtID.Text
    frmLedger.LinkOpen
End Sub

Private Sub txtMode_DblClick()
    XID = txtMode.Text
    frmLedger.LinkOpen
End Sub

Private Sub cmdPaymentForwarding_Click()
    CN(3).dbOpen "SELECT ID, NAME, ADDRESS, CITY, PHONES FROM APPVIEW_ALLACCOUNTS WHERE ID=" & QT(txtID.Text)
    With vp
        .StartDoc
        SetFont "[Render_Section_A]"
        .CurrentY = 200
        .TextAlign = taCenterTop
        .Text = "PAYMENT FORWARDING NOTE"
        .TextAlign = taLeftTop
        Y = .CurrentY + 200
        vp.DrawPicture mdiOne.ImgList.ListImages(1).Picture, .MarginLeft, Y + 300, 837, 1100
        .CurrentY = Y + 250
        SetFont "[Render_Section_B]"
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "0.60in"
            .TableCell(tcColWidth, , 2) = "5.00in"
            .TableCell(tcColAlign, , 2) = taLeftTop
            .TableCell(tcText, 1, 2) = CompanyName
        .EndTable
        SetFont "[Render_Section_C]"
        .StartTable
            .TableBorder = tbBottom
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "0.60in"
            .TableCell(tcColWidth, , 2) = "6.90in"
            .TableCell(tcColAlign, , 2) = taLeftTop
            .TableCell(tcText, 1, 2) = AboutCompany & vbCrLf & CompanyAddr0 & vbCrLf & CompanyPhone & ", " & CompanyFax
        .EndTable
        .CurrentY = .CurrentY + 150
        .FontName = "Bookman Old Style": .FontSize = 12
        .StartTable
            .TableBorder = tbNone: .TableCell(tcRowSpaceAfter) = 20
            .TableCell(tcCols) = 3: .TableCell(tcRows) = 3: .TableCell(tcColBorder, , 2, , 2) = 0
            .TableCell(tcRowSpan, 1, 1) = 3
            .TableCell(tcColWidth, , 1) = "4.0in": .TableCell(tcColWidth, , 2) = "2.3in": .TableCell(tcColWidth, , 3) = "1.4in"
            .TableCell(tcText, 1, 1) = "{\PAR\PAR\PAR To, " & " \par\b " & CN(3).recs!Name & " (" & CN(3).recs!id & ")" & " \par " & CN(3).recs!Address & " \par " & CN(3).recs!City & " \par Phone: " & CN(3).recs!Phones & " \PAR } "
        .EndTable
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 1: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "7.50in"
            .TableCell(tcColAlign, , 1) = taJustTop
            .TableCell(tcText, 1, 1) = vbCrLf & "Dear Sir," & vbCrLf & "      Remit Cheque/DD No. " & txtNarration.Text & " of  " & txtModeName.Text & " for Rs." & Format(Val(txtAmount.Text), "#0.00") & " (" & ConvertCurrencyToEnglish(txtAmount.Text) & ") against your outstanding dues. " & vbCrLf & vbCrLf & "Kindly acknowledge the same and send the upto date statement of accounts." & vbCrLf & vbCrLf & "Thanking you." & vbCrLf & vbCrLf
        .EndTable
        .StartTable
            .TableBorder = tbNone
            .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
            .TableCell(tcColWidth, , 1) = "3.75in"
            .TableCell(tcColWidth, , 2) = "3.75in"
            .TableCell(tcColAlign, , 1) = taJustTop
            .TableCell(tcColAlign, , 2) = taCenterTop
            .TableCell(tcText, 1, 1) = ""
            .TableCell(tcText, 1, 2) = "Yours truly" & vbCrLf & vbCrLf & "For " & CompanyName
        .EndTable
        .EndDoc
    End With
End Sub

Private Sub SetFont(S As String)
    vp.FontName = ReadFont(S, 0)
    vp.FontSize = ReadFont(S, 1)
    vp.FontBold = ReadFont(S, 2)
    vp.FontItalic = ReadFont(S, 3)
End Sub

Private Sub DUPLICATETHIS()
    cmdSave.Caption = "DUPLICATE"
    txtSerial.Visible = False
    txtDate.SetFocus
End Sub
