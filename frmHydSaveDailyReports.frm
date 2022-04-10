VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmHydSaveDailyReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HydSaveDailyReports"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBackupDaysReports 
      Caption         =   "SAVE"
      Height          =   975
      Left            =   1350
      TabIndex        =   0
      Top             =   2865
      Width           =   4080
   End
   Begin VSFlex8UCtl.VSFlexGrid flxItem 
      Align           =   1  'Align Top
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7035
      _cx             =   12409
      _cy             =   4921
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
      ExtendLastCol   =   0   'False
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
End
Attribute VB_Name = "frmHydSaveDailyReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(10) As New clsData

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    Me.Move 0, 0
End Sub

Private Sub cmdBackupDaysReports_Click()
    If MsgBox("Do you wish to save all transaction daily reports?", vbYesNo + vbExclamation) = vbYes Then
        mdiOne.CDlg.ShowSave
        
        SaveTables "SMAIN", Sale
        SaveTables "SRETURNMAIN", SaleReturn
        SaveTables "PMAIN", Purchase
        SaveTables "PRETURNMAIN", PurchaseReturn
        SaveTables "TINMAIN", StockTransferIN
        SaveTables "TOUTMAIN", StockTransferOUT
        
        SavePRT "PMT", PYMT
        SavePRT "RCT", RCPT
    End If
End Sub

Private Sub SaveTables(Main As String, F As Object)
    SQ = "SELECT * FROM " & Main & " Where DateDiff(D, DBDate," & QT(Format(Now, "DD-MMM-YY")) & ")=0 ORDER BY 1 Desc"
    CN(0).dbOpen SQ, 1
    flxItem.Clear
    Set flxItem.DataSource = CN(0)
    
    mdiOne.CDlg.FileName = Format(Main & "-" & Format(Now, "DD-MMM-YY HHMM")) & ".xls"
    flxItem.SaveGrid mdiOne.CDlg.FileName, flexFileExcel, flexXLSaveFixedCells
    
    While Not CN(0).recs.EOF
        F.frm.txtDBRef = CN(0).recs!DBRef
        F.frm.txtDBRef_LostFocus
        
        mdiOne.CDlg.FileName = Format(F.frm.txtDBRef.Text, F.frm.MemoFormat) & "-" & F.frm.txtID.Text & "-" & Format(Now, "DD-MMM-YY HHMM") & ".xls"
        F.frm.flxOrder.SaveGrid mdiOne.CDlg.FileName, flexFileExcel

        CN(0).recs.MoveNext
    Wend
End Sub

Private Sub SavePRT(Main As String, F As Object)
    SQ = "SELECT * FROM " & Main & " Where DateDiff(D, Date," & QT(Format(Now, "DD-MMM-YY")) & ")=0 ORDER BY 1 Desc"
    CN(0).dbOpen SQ, 1
    flxItem.Clear
    Set flxItem.DataSource = CN(0)
    
    mdiOne.CDlg.FileName = Format(Main & "-" & Format(Now, "DD-MMM-YY HHMM")) & ".xls"
    flxItem.SaveGrid mdiOne.CDlg.FileName, flexFileExcel, flexXLSaveFixedCells
End Sub

