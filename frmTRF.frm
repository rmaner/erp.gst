VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTRF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts Transactions..."
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTRF.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6165
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
         Left            =   4455
         Locked          =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "Name "
         Top             =   2160
         Width           =   1590
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
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Name "
         Top             =   1380
         Width           =   1590
      End
      Begin VB.CommandButton cmdSelectCrID 
         DownPicture     =   "frmTRF.frx":4E0E
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
         Left            =   1815
         Picture         =   "frmTRF.frx":5151
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Open Payee"
         Top             =   1785
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
         Left            =   885
         TabIndex        =   24
         ToolTipText     =   "Payee's ID"
         Top             =   1785
         Width           =   915
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
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Name "
         Top             =   2160
         Width           =   3570
      End
      Begin VB.CommandButton cmdSelectSerial 
         DownPicture     =   "frmTRF.frx":5494
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
         Picture         =   "frmTRF.frx":57D7
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Amount in words"
         Top             =   2610
         Width           =   4275
      End
      Begin VB.CommandButton cmdSelectDrID 
         DownPicture     =   "frmTRF.frx":5B1A
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
         Picture         =   "frmTRF.frx":5E5D
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   900
         TabIndex        =   12
         ToolTipText     =   "Payee's ID"
         Top             =   1005
         Width           =   915
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
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Name "
         Top             =   1380
         Width           =   3570
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
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Serial No."
         Top             =   180
         Width           =   915
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
         TabIndex        =   9
         ToolTipText     =   "Narration"
         Top             =   2940
         Width           =   5205
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
         Top             =   2610
         Width           =   915
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
         TabIndex        =   7
         ToolTipText     =   "Adds New Record"
         Top             =   3270
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
         Left            =   1245
         TabIndex        =   6
         ToolTipText     =   "Adds New Record"
         Top             =   3270
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
         Left            =   2445
         TabIndex        =   5
         ToolTipText     =   "Adds New Record"
         Top             =   3270
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
         Left            =   3645
         TabIndex        =   4
         ToolTipText     =   "Adds New Record"
         Top             =   3270
         Width           =   1215
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
         Left            =   4845
         TabIndex        =   3
         Top             =   3270
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   330
         Left            =   900
         TabIndex        =   16
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
         Format          =   21561347
         CurrentDate     =   38023
      End
      Begin MSMask.MaskEdBox txtCrAccountBalance 
         Height          =   300
         Left            =   3615
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1770
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
      Begin MSMask.MaskEdBox txtDrAccountBalance 
         Height          =   300
         Left            =   3645
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1005
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CrID:"
         Height          =   270
         Index           =   3
         Left            =   45
         TabIndex        =   27
         Top             =   1785
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   270
         Index           =   2
         Left            =   45
         TabIndex        =   26
         Top             =   2190
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DrID:"
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   1005
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   2625
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial:"
         Height          =   270
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   195
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   270
         Index           =   2
         Left            =   60
         TabIndex        =   19
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   270
         Index           =   1
         Left            =   60
         TabIndex        =   18
         Top             =   1410
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Narration:"
         Height          =   270
         Index           =   4
         Left            =   60
         TabIndex        =   17
         Top             =   2955
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   3735
      Left            =   6165
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VSPrinter8LibCtl.VSPrinter vp 
         Height          =   3510
         Left            =   90
         TabIndex        =   1
         Top             =   165
         Width           =   5565
         _cx             =   9816
         _cy             =   6191
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
         Zoom            =   17.0454545454545
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
End
Attribute VB_Name = "frmTRF"
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

Private Sub Form_Load()
    Me.Move 0, 0
    txtDate.Value = Now
    mdiOne.SetFormFont Me
    SetVPFont ("[PRT-VP-FONT]")
    Frame2.BackColor = FRAMECOLOR
    vp.Zoom = 88
End Sub

'==== TAB ADD
Private Sub txtSerial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtSerial.Text = Val(txtSerial.Text) + 1: txtSerial_LostFocus
    End If
    If KeyCode = vbKeyDown Then
        txtSerial.Text = Val(txtSerial.Text) - 1: txtSerial_LostFocus
    End If
End Sub

Private Sub txtSerial_LostFocus()
    txtSerial.Text = Val(txtSerial.Text)
    CN(0).dbOpen "SELECT * FROM " & Support & " WHERE SERIAL=" & Val(txtSerial.Text)
    If CN(0).recs.RecordCount = 1 Then
        txtDate.Value = CN(0).recs!Date
        txtDrID.Text = CN(0).recs!DrID
        txtDrName.Text = CN(0).recs!DrName: txtDrCity.Text = CN(0).recs!DrCity
        txtCrID.Text = CN(0).recs!CrID
        txtCrName.Text = CN(0).recs!CrName: txtCrCity.Text = CN(0).recs!CrCity
        txtAmount.Text = CN(0).recs!Amount
        txtAmountWords.Text = ConvertCurrencyToEnglish(CN(0).recs!Amount)
        txtNarration.Text = CN(0).recs!Narration
    Else
        MsgBox "No such record! Check again.", vbOKOnly + vbCritical
        cmdClear_Click
    End If
    Preview
    txtDrID.SetFocus
End Sub

Private Sub cmdSelectSerial_Click()
    frmShow.Init "SELECT * FROM  " & Support & " ORDER BY 1 DESC"
    If sArray(0) <> "" Then
        txtSerial.Text = sArray(0)
    End If
    txtSerial_LostFocus
    txtDrAccountBalance.Text = Val(WhatIsLedgerBalance(txtDrID.Text))
    txtCrAccountBalance.Text = Val(WhatIsLedgerBalance(txtCrID.Text))
End Sub

Private Sub txtDrID_Change()
    CN(1).dbOpen "SELECT * FROM appview_ALLACCOUNTS WHERE ID=" & QT(txtDrID.Text)
    If CN(1).recs.RecordCount = 1 Then
        txtDrName.Text = CN(1).recs!Name: txtDrCity.Text = CN(1).recs!City
    Else
       txtDrName.Text = "": txtDrCity.Text = ""
    End If
    txtDrAccountBalance.Text = Val(WhatIsLedgerBalance(txtDrID.Text))
End Sub

Private Sub cmdSelectDrID_Click()
    frmShow.Init "SELECT * FROM appview_ALLACCOUNTS"
    If sArray(0) <> "" Then
        txtDrID.Text = sArray(0)
    End If
    txtDrID.SetFocus
End Sub

Private Sub txtCrID_Change()
    CN(1).dbOpen "SELECT * FROM appview_ALLACCOUNTS WHERE ID=" & QT(txtCrID.Text)
    If CN(1).recs.RecordCount = 1 Then
        txtCrName.Text = CN(1).recs!Name: txtCrCity.Text = CN(1).recs!City
    Else
       txtCrName.Text = "": txtCrCity.Text = ""
    End If
    txtCrAccountBalance.Text = Val(WhatIsLedgerBalance(txtCrID.Text))
End Sub

Private Sub cmdSelectCrID_Click()
    frmShow.Init "SELECT * FROM appview_ALLACCOUNTS"
    If sArray(0) <> "" Then
        txtCrID.Text = sArray(0)
    End If
    txtCrID.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.Text = Format(Val(txtAmount.Text), "##0.00")
    txtAmountWords.Text = ConvertCurrencyToEnglish(Val(txtAmount.Text))
End Sub

Private Sub cmdNew_Click()
    On Error Resume Next
    If MsgBox("Are you sure to add?", vbYesNo) = vbYes Then
        CN(0).dbOpen ADDPROC
        Set CN(0).recs = CN(0).recs.NextRecordset
        txtSerial.Text = CN(0).recs!MaxSerial
    End If
    txtSerial_LostFocus
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    If MsgBox("Are you sure to update?", vbYesNo) = vbYes Then
        If Val(txtSerial.Text) <> 0 And txtDrID.Text <> "" And txtCrID.Text <> "" Then
            CN(2).dbOpen UPDATEPROC & Val(txtSerial.Text) & ", " & QT(Format(txtDate.Value, "dd-MMM-yy HH:MM")) & ", " & QT(UCase(txtDrID.Text)) & ", " & QT(txtDrName.Text) & ", " & QT(UCase(txtCrID.Text)) & ", " & QT(txtCrName.Text) & ", " & Val(txtAmount.Text) & ", " & QT(txtNarration.Text)
            cJr.TRFJournalEntries Support, Val(txtSerial.Text), MemoFormat
        Else
            MsgBox "ID or MODE error!", vbOKOnly + vbCritical
        End If
    End If
    txtSerial_LostFocus
    cmdAdd.SetFocus
End Sub

Private Sub cmdClear_Click()
    txtSerial.Text = ""
    txtDrID.Text = ""
    txtDrName.Text = "": txtDrCity = ""
    txtCrID.Text = ""
    txtCrName.Text = "": txtCrCity = ""
    txtAmount.Text = 0
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
    PprDim = Split(GReadINI("[PaperHeight/Width/Left/Right/Top/Bottom]"), ":")
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
    If Support = "TRF" Then
        TRFPrint
    End If
End Sub

Private Sub TRFPrint()
    If Val(txtSerial.Text) <> 0 Then
    With vp
        .Zoom = 90
        headString = "{\f0\fs16 TRANSFER NOTE\par \b\fs22 " & CompanyName & "\par \b0\fs20 " & CompanyAddress & " \par " & CompanyPhone & "; " & CompanyMobile & "\par}"
        bodyString = "{\b0\f0\fs18 No.: " & Format(Val(txtSerial.Text), "0000") & " (" & txtDrID.Text & "/" & ")\tab\tab DATE: " & Format(txtDate.Value, "dd-MM-yy HH:MM:SS") & "\par\par Debit Account: " & txtDrName.Text & " \par Credit Account: " & txtCrName.Text & " \par  Amount: Rs." & txtAmount.Text & "\par Narration: " & txtNarration.Text & "\par \pard\qr\i \par\par Authorised Signatory\i0\par \f1}"
        
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



Private Sub txtDrID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtCrID.SetFocus
End Sub

Private Sub txtCrID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtAmount.SetFocus
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

