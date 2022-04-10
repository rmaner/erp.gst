VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmVCH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VCH"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtVSerial 
      Height          =   330
      Left            =   1815
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      ToolTipText     =   "Ctrl+0"
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox txtVType 
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
      Left            =   5850
      TabIndex        =   3
      Top             =   0
      Width           =   1785
   End
   Begin VB.TextBox txtVRef 
      Height          =   330
      Left            =   465
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Ctrl+0"
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox pboxA 
      Align           =   2  'Align Bottom
      Height          =   1515
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   7590
      TabIndex        =   11
      Top             =   3075
      Width           =   7650
      Begin VB.TextBox txtNarration 
         Height          =   1215
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   7605
      End
      Begin VB.Label Label9 
         Caption         =   "Narration:"
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
         Left            =   0
         TabIndex        =   13
         Top             =   30
         Width           =   1695
      End
   End
   Begin VB.PictureBox pboxB 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7590
      TabIndex        =   9
      Top             =   4590
      Width           =   7650
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   4215
         TabIndex        =   6
         Top             =   0
         Width           =   1125
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
         Left            =   5340
         TabIndex        =   7
         Top             =   0
         Width           =   1125
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
         Left            =   6465
         TabIndex        =   8
         Top             =   0
         Width           =   1125
      End
      Begin MSForms.TextBox txtUserNo 
         Height          =   300
         Left            =   0
         TabIndex        =   10
         Top             =   15
         Width           =   3060
         VariousPropertyBits=   746604567
         ForeColor       =   32768
         BorderStyle     =   1
         Size            =   "5397;529"
         Value           =   "- by "
         BorderColor     =   -2147483640
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComCtl2.DTPicker txtVDate 
      Height          =   330
      Left            =   3270
      TabIndex        =   2
      Top             =   0
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
      Format          =   20381699
      CurrentDate     =   38023
   End
   Begin VSFlex8UCtl.VSFlexGrid flxVCH 
      Height          =   2670
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   7650
      _cx             =   13494
      _cy             =   4710
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
      BackColorBkg    =   8421504
      BackColorAlternate=   14737632
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
   Begin VB.Label Label3 
      Caption         =   "VID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   12
      Top             =   45
      Width           =   810
   End
End
Attribute VB_Name = "frmVCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(25) As New clsData
Private cJr As New clsJournals

Public HiddenCols As String
Public Main As String
Public Support As String
Public MainAddNew As String
Public SerialCol, VRefXCol, DrCrCol, AcIDCol, AcNameCol, AmountCol, DrCol, CrCol As Integer
Private VRef As Integer

Private Sub Form_Load()
    Me.Move 0, 0
    Me.Caption = MyCAPTION
    mdiOne.SetFormFont Me
    
    Main = "VMain": Support = "Vouchers": MainAddNew = "ADDNEW_VMAIN"
    SerialCol = 0: VRefXCol = 1: DrCrCol = 2: AcIDCol = 3: AcNameCol = 4: AmountCol = 5: DrCol = 6: CrCol = 7
    txtVDate.Value = Now
    txtVType.Clear: txtVType.AddItem "CONTRA": txtVType.AddItem "PAYMENTS": txtVType.AddItem "RECEIPTS": txtVType.AddItem "JOURNAL": txtVType.AddItem "SALES": txtVType.AddItem "CREDITNOTE": txtVType.AddItem "PURCHASE": txtVType.AddItem "DEBITNOTE": txtVType.AddItem "REVJOURNAL": txtVType.AddItem "MEMOS": txtVType.ListIndex = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyN And Shift = vbCtrlMask Then
        txtVRef.Text = "NEW"
        CN(0).dbOpen "appproc_ReturnNextVSerial " & QT(txtVType.Text), 1
        If Not CN(0).recs.EOF Then txtVSerial.Text = CN(0).recs!NextVSerial
        txtVDate.SetFocus
    End If
    If KeyCode = 188 And Shift = vbCtrlMask Then Call ColShowHide(1)             '< Hide
    If KeyCode = 190 And Shift = vbCtrlMask Then Call ColShowHide(0)             '> Show
    If KeyCode = 191 And Shift = vbCtrlMask Then                                 '/ hide one
        flxVCH.ColHidden(flxVCH.Col) = True
        If flxVCH.Col < flxVCH.COLS - 2 Then flxVCH.Col = flxVCH.Col + 1
    End If
    If KeyCode = vbKeyF12 Then cmdSave_Click
End Sub

Private Sub txtVRef_Change()
    If txtVRef.Text = "NEW" Then
        txtVRef.BackColor = vbRed
    Else
        txtVRef.BackColor = vbWhite
    End If
End Sub

Private Sub txtVRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtVRef.Text = Val(txtVRef.Text) + 1
        txtVRef_LostFocus
    End If
    If KeyCode = vbKeyDown Then
        txtVRef.Text = Val(txtVRef.Text) - 1
        txtVRef_LostFocus
    End If
End Sub

Private Sub txtVRef_LostFocus()
    If txtVRef.Text = "NEW" Then
        PictureBoxStatus (True)
    Else
        CN(1).dbOpen "Select * from " & Main & " where VRef=" & Val(txtVRef.Text)
        VRef = Val(sArray(0))
        If Val(sArray(0)) <> 0 Then
            FillFormText 'Filling of Forms's text boxes.
            CN(2).dbOpen "Select * from " & Support & " where VRefX=" & Val(txtVRef.Text) & " Order by Serial", 1
            Set flxVCH.DataSource = CN(2).recs
            ColShowHide (1)
            PictureBoxStatus (True)
        Else
            txtVRef.Text = "0"
            PictureBoxStatus (False)
        End If
        flxVCH.ColWidth(SerialCol) = 350
        flxVCH.ColWidth(VRefXCol) = 350
        flxVCH.ColWidth(DrCrCol) = 350
        flxVCH.ColWidth(AcIDCol) = 900
        flxVCH.ColWidth(AcNameCol) = 2000
        flxVCH.ColWidth(AmountCol) = 1500
        flxVCH.ColWidth(DrCol) = 700
        flxVCH.ColWidth(CrCol) = 700
    End If
End Sub

Private Sub txtVSerial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtVSerial.Text = Val(txtVSerial.Text) + 1
        txtVSerial_LostFocus
    End If
    If KeyCode = vbKeyDown Then
        txtVSerial.Text = Val(txtVSerial.Text) - 1
        txtVSerial_LostFocus
    End If
End Sub

Private Sub txtVSerial_LostFocus()
    CN(3).dbOpen "Select * from " & Main & " where VSerial=" & Val(txtVSerial.Text) & " AND VType=" & QT(txtVType.Text)
    If Not CN(3).recs.EOF Then
        txtVRef.Text = CN(3).recs!VRef
        txtVRef_LostFocus
    Else
        txtVSerial.Text = "0"
    End If
End Sub

Private Sub flxVch_EnterCell()
    With flxVCH
        c = .Col
        Select Case c
            Case 2, 3, 4, 5, 6, 7
                .Editable = flexEDKbdMouse
            Case Else
                .Editable = flexEDNone
        End Select
    End With
End Sub

Private Sub flxVch_LeaveCell()
    With flxVCH
    For R = 1 To .ROWS - 1
        For c = 0 To .COLS - 1
        Select Case c
            Case 5, 6, 7
                .TextMatrix(R, c) = Val(.TextMatrix(R, c))
        End Select
        Next
    Next
    End With
    flxVCH.Editable = flexEDNone
End Sub

Private Sub flxVch_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = AcNameCol Then
        'asdf
    End If
    
    If KeyAscii = vbKeyReturn And (Col = AcNameCol) Then
        'asdf
    End If
    Calculate
End Sub

Private Sub flxVch_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
End Sub

Private Sub flxVch_AfterSort(ByVal Col As Long, Order As Integer)
    EnumerateGrid
End Sub

Private Sub Calculate()
    'asdfasdf
End Sub

Private Sub cmdDeleteBill_Click()
    If MsgBox("Confirm deletion?", vbYesNo + vbQuestion) = vbYes Then
        'Delete Memo and its related journal entries; but leaves the order intact.
        txtVRef_LostFocus
    End If
End Sub

Public Sub cmdSave_Click()
    If txtVRef.Text = "NEW" Then
        CN(4).dbOpen "ADDNEW_VMAIN " & Val(txtVSerial.Text) & ", " & QT(Format(txtVDate.Value, "dd-MMM-yyyy")) & ", " & QT(txtVType.Text) & ", " & QT(txtNarration.Text) & ", " & QT(GUID)
        Set CN(4).recs = CN(4).recs.NextRecordset: txtVRef.Text = CN(4).recs!VRef
        DeleteString = "Delete " & Support & " where VRefX=" & Val(txtVRef.Text)
        SaveString = "Select * from " & Support & " where VRefX=" & Val(txtVRef.Text) & " order by Serial"
        toSupport DeleteString, SaveString, flxVCH
    Else
        DeleteString = "Delete " & Support & " where VRefX=" & Val(txtVRef.Text)
        SaveString = "Select * from " & Support & " where VRefX=" & Val(txtVRef.Text) & " order by Serial"
        toMain
        toSupport DeleteString, SaveString, flxVCH
    End If
    frmTimeOne.Show
    msgUITS Support & " Data Saved!"
    Call txtVRef_LostFocus
End Sub

Private Sub cmdQuit_Click()
    Erase sArray
    Unload Me
End Sub

'================================================
' SUB ROUTINES
'================================================

Private Sub FillFormText()
    i = 0
'   txtVRef.Text = sArray(i): i = i + 1
    i = i + 1
    txtVSerial.Text = sArray(i): i = i + 1
    txtVDate.Value = sArray(i): i = i + 1
    txtVType.Text = sArray(i): i = i + 1
    txtNarration.Text = sArray(i): i = i + 1
    txtUserNo.Text = sArray(i): i = i + 1
End Sub

Public Sub toMain()
    CN(5).dbOpen "Select * from " & Main & " where VRef=" & VRef, 1
'       cn(5).recs!VRef = txtVRef.Text
        CN(5).recs!vserial = txtVSerial.Text
        CN(5).recs!VDate = txtVDate.Value
        CN(5).recs!VType = txtVType.Text
        CN(5).recs!Narration = txtNarration.Text
        CN(5).recs!UserNo = Left(Left(Status, 2) & " " & Format(flxVCH.ROWS - 1, "00") & " " & Format(Now, "DDMMM hhmm") & " " & Left(GUID, 4) & "|" & CN(5).recs!UserNo, 198)
    CN(5).recs.Update
End Sub

Public Sub toSupport(ByVal DeleteStr As String, ByVal SaveStr As String, MyFlex As Control)
    Dim i As Integer
    CN(6).dbOpen DeleteStr, 1
    CN(6).dbOpen SaveStr, 1
    With MyFlex
        For i = 1 To .ROWS - 1
            If ValidateAcID((.TextMatrix(i, 3))) = True Then
                CN(6).recs.AddNew
                CN(6).recs!Serial = i              '.TextMatrix(i, 0)
                CN(6).recs!VRefX = Val(txtVRef.Text)
                CN(6).recs!DRCR = Val(.TextMatrix(i, DrCrCol))
                CN(6).recs!AcID = Trim(.TextMatrix(i, AcIDCol))
                CN(6).recs!AcName = Left(.TextMatrix(i, AcNameCol), 200)
                CN(6).recs!Amount = Val(.TextMatrix(i, AmountCol))
                CN(6).recs!Dr = Val(.TextMatrix(i, DrCol))
                CN(6).recs!Cr = Val(.TextMatrix(i, CrCol))
            End If
        Next
    End With
    CN(6).recs.UpdateBatch
End Sub

Private Sub EnumerateGrid()
    For i = 1 To flxVCH.ROWS - 1
        flxVCH.TextMatrix(i, 0) = i
    Next
End Sub

Private Function ValidateAcID(ByVal AcID As Long) As Boolean
    ValidateAcID = HasAccount(AcID)
End Function

Private Sub ColShowHide(ByVal Opt As Integer)
    If Opt = 0 Then     'Show
        For i = 0 To flxVCH.COLS - 1
            flxVCH.ColHidden(i) = False
        Next
    End If
    If Opt = 1 Then     'Hide
        For Each i In Split(GReadINI(HiddenCols), ",")
            flxVCH.ColHidden(i) = True
        Next
    End If
End Sub

Private Sub PictureBoxStatus(S As Boolean)
    pboxA.Enabled = S
    pboxB.Enabled = S
End Sub
