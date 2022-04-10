VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmVouchers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vouchers..."
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6585
      Begin VB.TextBox txtVType 
         Alignment       =   2  'Center
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
         Left            =   5415
         Locked          =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "Name "
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdSelectCrID 
         DownPicture     =   "frmVouchers.frx":0000
         Height          =   300
         Left            =   1830
         Picture         =   "frmVouchers.frx":0343
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Open Payee"
         Top             =   1725
         UseMaskColor    =   -1  'True
         Width           =   375
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
         Left            =   1065
         TabIndex        =   23
         ToolTipText     =   "Payee's ID"
         Top             =   1725
         Width           =   765
      End
      Begin VB.TextBox txtCrName 
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
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "Name "
         Top             =   1725
         Width           =   3180
      End
      Begin VB.TextBox txtCrCity 
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
         Left            =   5415
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Name "
         Top             =   1725
         Width           =   1125
      End
      Begin VB.CommandButton cmdSelectSerial 
         DownPicture     =   "frmVouchers.frx":0686
         Height          =   330
         Left            =   1800
         Picture         =   "frmVouchers.frx":09C9
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Open Serial"
         Top             =   165
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Amount in words"
         Top             =   2415
         Width           =   4110
      End
      Begin VB.CommandButton cmdSelectDrID 
         DownPicture     =   "frmVouchers.frx":0D0C
         Height          =   300
         Left            =   1830
         Picture         =   "frmVouchers.frx":104F
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Open Payee"
         Top             =   1005
         UseMaskColor    =   -1  'True
         Width           =   375
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
         Left            =   1065
         TabIndex        =   11
         ToolTipText     =   "Payee's ID"
         Top             =   1005
         Width           =   765
      End
      Begin VB.TextBox txtDrName 
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
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Name "
         Top             =   1005
         Width           =   3180
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
         Left            =   1065
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Serial No."
         Top             =   180
         Width           =   735
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
         Left            =   1065
         TabIndex        =   8
         ToolTipText     =   "Narration"
         Top             =   2790
         Width           =   5460
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
         Left            =   1065
         TabIndex        =   7
         ToolTipText     =   "Amount"
         Top             =   2415
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
         Left            =   30
         TabIndex        =   6
         ToolTipText     =   "Adds New Record"
         Top             =   3660
         Width           =   1305
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
         Left            =   1335
         TabIndex        =   5
         ToolTipText     =   "Adds New Record"
         Top             =   3660
         Width           =   1305
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
         Left            =   2640
         TabIndex        =   4
         ToolTipText     =   "Adds New Record"
         Top             =   3660
         Width           =   1305
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
         Left            =   3945
         TabIndex        =   3
         ToolTipText     =   "Adds New Record"
         Top             =   3660
         Width           =   1305
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
         Left            =   5250
         TabIndex        =   2
         Top             =   3660
         Width           =   1305
      End
      Begin VB.TextBox txtDrCity 
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
         Left            =   5415
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Name "
         Top             =   1005
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   330
         Left            =   1065
         TabIndex        =   15
         Top             =   585
         Width           =   1125
         _ExtentX        =   1984
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
      Begin MSMask.MaskEdBox txtDrAccountBalance 
         Height          =   255
         Left            =   4515
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1335
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   450
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
      Begin MSMask.MaskEdBox txtCrAccountBalance 
         Height          =   255
         Left            =   4515
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   450
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
      Begin MSForms.TextBox txtUserNo 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   3240
         Width           =   6375
         VariousPropertyBits=   746604567
         ForeColor       =   32768
         BorderStyle     =   1
         Size            =   "11245;556"
         Value           =   "- by "
         BorderColor     =   -2147483640
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.SpinButton spinVType 
         Height          =   255
         Left            =   5415
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   465
         Width           =   1125
         Size            =   "1984;450"
         Max             =   4
         Orientation     =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit A/c:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   25
         Top             =   1725
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Debit A/c:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   20
         Top             =   1020
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   19
         Top             =   2430
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   18
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   30
         TabIndex        =   17
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Narration:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   30
         TabIndex        =   16
         Top             =   2805
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmVouchers"
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

Public DrAcSQL As String
Public CrAcSQL As String

Private CN(5) As New clsData
Private cJr As New clsJournals

Private Sub Form_Load()
    Me.Move 0, 0
    
    Me.MyCAPTION = "VOUCHERS..."
    Me.FRAMECOLOR = vbGreen
    Me.MemoFormat = "\V\R0"
    Me.ADDPROC = "ADDNEW_VOUCHERS "
    Me.UPDATEPROC = "UPDATE_VOUCHERS "
    Me.Support = "VOUCHERS"
    
    Me.Caption = MyCAPTION
    txtDate.Value = Now
    mdiOne.SetFormFont Me
    Frame2.BackColor = FRAMECOLOR
End Sub

Private Sub spinVType_Change()
    Select Case spinVType.Value
        Case 0: txtVType.Text = "PYMT"
        Case 1: txtVType.Text = "RCPT"
        Case 2: txtVType.Text = "CTRA"
        Case 3: txtVType.Text = "JRNL"
        Case 4: txtVType.Text = "MEMO"
        Case Else: txtVType.Text = "MEMO"
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageUp Then cmdSelectDrID_Click
    If KeyCode = vbKeyPageDown Then cmdSelectCrID_Click
    If KeyCode = vbKeyF12 Then cmdSave_Click
    If KeyCode = vbKeyEscape Then
        cmdSave.Caption = "&Save"
        txtSerial.Visible = True
    End If
    If (KeyCode = 98 Or KeyCode = 50) And Shift = vbCtrlMask Then DUPLICATETHIS     'TWO
End Sub

'==== TAB ADD
Private Sub txtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtSerial.Text = Val(txtSerial.Text) + 1: txtSerial_LostFocus
    End If
    If KeyCode = vbKeyDown Then
        txtSerial.Text = Val(txtSerial.Text) - 1: txtSerial_LostFocus
    End If
    If KeyCode = vbKeyReturn Then txtDate.SetFocus
End Sub

Private Sub txtSerial_LostFocus()
    CN(0).dbOpen "SELECT ID, NAME FROM PERSONAL WHERE TYPE=" & QT("CASH") & " OR TYPE=" & QT("BANK")
    txtSerial.Text = Val(txtSerial.Text)
    CN(0).dbOpen "SELECT * FROM " & Support & " WHERE SERIAL=" & Val(txtSerial.Text)
    If CN(0).recs.RecordCount = 1 Then
        txtDate.Value = CN(0).recs!Date
        txtDrID.Text = CN(0).recs!DrID: txtDrName.Text = CN(0).recs!DrName: txtDrCity.Text = CN(0).recs!DrCity
        txtCrID.Text = CN(0).recs!CrID: txtCrName.Text = CN(0).recs!CrName: txtCrCity.Text = CN(0).recs!CrCity
        txtAmount.Text = CN(0).recs!Amount
        txtAmountWords.Text = ConvertCurrencyToEnglish(CN(0).recs!Amount)
        txtNarration.Text = CN(0).recs!Narration
        txtVType.Text = CN(0).recs!VType
    Else
        MsgBox "No such record! Check again.", vbOKOnly + vbCritical
    End If
    txtDrAccountBalance.Text = Val(WhatIsLedgerBalance(txtDrID.Text))
    txtCrAccountBalance.Text = Val(WhatIsLedgerBalance(txtCrID.Text))
End Sub

Private Sub cmdSelectSerial_Click()
    frmShow.Init "SELECT * FROM  " & Support & " ORDER BY 1 DESC"
    If sArray(0) <> "" Then
        txtSerial.Text = sArray(0)
    End If
    txtSerial_LostFocus
End Sub

Private Sub txtDrID_Change()
    CN(1).dbOpen "SELECT * FROM appview_ALLACCOUNTS WHERE ID=" & QT(txtDrID.Text)
    If CN(1).recs.RecordCount = 1 Then
        txtDrName.Text = CN(1).recs!Name: txtDrCity.Text = CN(1).recs!City
    Else
       txtDrName.Text = "_": txtDrCity.Text = "_"
    End If
    txtDrAccountBalance.Text = Val(WhatIsLedgerBalance(txtDrID.Text))
End Sub

Private Sub cmdSelectDrID_Click()
    frmShow.Init "SELECT * FROM appview_ALLACCOUNTS ORDER BY NAME"
    If sArray(0) <> "" Then txtDrID.Text = sArray(0)
    txtDrID.SetFocus
End Sub


Private Sub txtCrID_Change()
    CN(1).dbOpen "SELECT * FROM appview_ALLACCOUNTS WHERE ID=" & QT(txtCrID.Text)
    If CN(1).recs.RecordCount = 1 Then
        txtCrName.Text = CN(1).recs!Name: txtCrCity.Text = CN(1).recs!City
    Else
       txtCrName.Text = "_": txtCrCity.Text = "_"
    End If
    txtCrAccountBalance.Text = Val(WhatIsLedgerBalance(txtCrID.Text))
End Sub

Private Sub cmdSelectCrID_Click()
    frmShow.Init "SELECT * FROM appview_ALLACCOUNTS ORDER BY NAME"
    If sArray(0) <> "" Then txtCrID.Text = sArray(0)
    txtCrID.SetFocus
End Sub


Private Sub txtAmount_LostFocus()
    txtAmount.Text = Format(Val(txtAmount.Text), "##0.00")
    txtAmountWords.Text = ConvertCurrencyToEnglish(Val(txtAmount.Text))
End Sub

Private Sub cmdNew_Click()
    On Error Resume Next
    If MsgBox("Are you sure to add?", vbYesNo) = vbYes Then
        CN(0).dbOpen ADDPROC & " " & QT(txtVType.Text)
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
        If Val(txtSerial.Text) <> 0 And txtDrID.Text <> "" And txtCrID.Text <> "" Then
            CN(2).dbOpen UPDATEPROC & Val(txtSerial.Text) & ", " & QT(Format(txtDate.Value, "dd-MMM-yy HH:MM")) _
            & ", " & QT(UCase(txtDrID.Text)) & ", " & QT(txtDrName.Text) & ", " & QT(txtDrCity.Text) _
            & ", " & QT(UCase(txtCrID.Text)) & ", " & QT(txtCrName.Text) & ", " & QT(txtCrCity.Text) _
            & ", " & Val(txtAmount.Text) & ", " & QT(txtNarration.Text) & ", " & QT(txtVType.Text)
            cJr.VoucherToJournalEntries Val(txtSerial.Text), MemoFormat
        Else
            MsgBox "ID or MODE error!", vbOKOnly + vbCritical
        End If
    End If
    txtSerial.SetFocus
    txtSerial_LostFocus
    cmdAdd.SetFocus
End Sub

Private Sub cmdPrint_Click()
'    Preview
'    vp.PrintDoc
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

Private Sub txtDrID_KeyPress(KeyAscii As Integer)
    txtDrID_Change
    If KeyAscii = vbKeyReturn Then txtCrID.SetFocus
End Sub

Private Sub txtCrID_KeyPress(KeyAscii As Integer)
    txtCrID_Change
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

Private Sub txtDrID_DblClick()
    XID = txtDrID.Text
    frmLedger.LinkOpen
End Sub

Private Sub txtCrID_DblClick()
    XID = txtCrID.Text
    frmLedger.LinkOpen
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

