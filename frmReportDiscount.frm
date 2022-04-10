VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmReportDiscount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Discounts..."
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmReportDiscount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VSFlex8UCtl.VSFlexGrid flxInDiscounts 
      Align           =   1  'Align Top
      Height          =   1245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _cx             =   8255
      _cy             =   2196
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
      BackColorSel    =   12632319
      ForeColorSel    =   64
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
      AutoSearch      =   1
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
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
      Editable        =   2
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
   Begin VSFlex8UCtl.VSFlexGrid flxOutDiscounts 
      Align           =   1  'Align Top
      Height          =   1500
      Left            =   0
      TabIndex        =   1
      Top             =   1245
      Width           =   4680
      _cx             =   8255
      _cy             =   2646
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
      BackColorSel    =   12632319
      ForeColorSel    =   64
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
      AutoSearch      =   1
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
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
      Editable        =   2
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
Attribute VB_Name = "frmReportDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(2) As New clsData
Private Const MyCAPTION = "In/Out Discounts~~~"

Private Sub Form_Load()
    On Error Resume Next
    Me.Move 0, 0
    mdiOne.SetFormFont Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub flxInDiscounts_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW
        For i = 0 To flxInDiscounts.COLS - 1
            flxInDiscounts.ColHidden(i) = False
        Next
    End If
    If KeyCode = 191 And Shift = vbCtrlMask Then            '/ hide one
        flxInDiscounts.ColHidden(flxInDiscounts.Col) = True
        If flxInDiscounts.Col < flxInDiscounts.COLS - 2 Then flxInDiscounts.Col = flxInDiscounts.Col + 1
    End If
End Sub

Private Sub flxOutDiscounts_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 190 And Shift = vbCtrlMask Then  '> SHOW
        For i = 0 To flxOutDiscounts.COLS - 1
            flxOutDiscounts.ColHidden(i) = False
        Next
    End If
    If KeyCode = 191 And Shift = vbCtrlMask Then            '/ hide one
        flxOutDiscounts.ColHidden(flxOutDiscounts.Col) = True
        If flxOutDiscounts.Col < flxOutDiscounts.COLS - 2 Then flxOutDiscounts.Col = flxOutDiscounts.Col + 1
    End If
End Sub

Private Sub flxindiscounts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = MyCAPTION & "Sum on " & str(flxInDiscounts.Col) & " = " & Format(FlxSum(Me.flxInDiscounts), "##,##0.00")
    If Button = vbKeyRButton And Shift = vbCtrlMask Then SaveGrid Me.flxInDiscounts
End Sub

Private Sub flxoutdiscounts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Caption = MyCAPTION & "Sum on " & str(flxOutDiscounts.Col) & " = " & Format(FlxSum(Me.flxOutDiscounts), "##,##0.00")
    If Button = vbKeyRButton And Shift = vbCtrlMask Then SaveGrid Me.flxOutDiscounts
End Sub

Public Sub LoadReport(ByVal SQ0 As String, ByVal SQ1 As String)
    On Error Resume Next
    CN(0).dbOpen SQ0
    Set flxInDiscounts.DataSource = Nothing: Set flxInDiscounts.DataSource = CN(0)
    
    CN(1).dbOpen SQ1
    Set flxOutDiscounts.DataSource = Nothing: Set flxOutDiscounts.DataSource = CN(1)
    Me.Show: Me.ZOrder 0
End Sub

