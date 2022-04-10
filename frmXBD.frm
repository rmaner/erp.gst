VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Begin VB.Form frmXBD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DiscountMap..."
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12645
   Icon            =   "frmXBD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   12585
      TabIndex        =   2
      Top             =   0
      Width           =   12645
      Begin VB.CommandButton cmdReturnDiscountTest 
         Caption         =   "ReturnDiscountTest"
         Height          =   330
         Left            =   10695
         TabIndex        =   8
         Top             =   0
         Width           =   1905
      End
      Begin VB.CommandButton cmdSelectID 
         DownPicture     =   "frmXBD.frx":114DA
         Height          =   315
         Left            =   1260
         Picture         =   "frmXBD.frx":1181D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Open Payee"
         Top             =   0
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
         Left            =   330
         TabIndex        =   4
         ToolTipText     =   "Payee's ID"
         Top             =   0
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
         Left            =   1635
         TabIndex        =   3
         ToolTipText     =   "Name "
         Top             =   0
         Width           =   9015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   15
         Width           =   315
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxPub 
      Align           =   3  'Align Left
      Height          =   6825
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   6675
      _cx             =   11774
      _cy             =   12039
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
   Begin VSFlex8UCtl.VSFlexGrid flxPersonal 
      Align           =   4  'Align Right
      Height          =   6825
      Left            =   6720
      TabIndex        =   1
      Top             =   375
      Width           =   5925
      _cx             =   10451
      _cy             =   12039
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
   Begin VSFlex8UCtl.VSFlexGrid flxItems 
      Align           =   2  'Align Bottom
      Height          =   2205
      Left            =   0
      TabIndex        =   7
      Top             =   7200
      Width           =   12645
      _cx             =   22304
      _cy             =   3889
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
Attribute VB_Name = "frmXBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    sSQL(0) = "SELECT DISTINCT PUBLISHERID AS PUBID, PUBLISHERNAME, PDISCOUNT AS PDISC, SDISCOUNT AS SDISC, SDISCOUNT AS DISC FROM ITEMS ORDER BY PUBLISHERID, SDISCOUNT"
    dbOpen (0): Set flxPub.DataSource = recs(0): dbClose (0)
    flxPub.ColWidth(1) = 2500
    flxPub.AutoSearch = flexSearchNone
End Sub

Private Sub txtID_Change()
    sSQL(5) = "SELECT * FROM appview_ALLACCOUNTS WHERE ID=" & QT(txtID.Text)
    dbOpen (5)
    If recs(5).RecordCount = 1 Then
        txtName.Text = recs(5)!Name & ", " & recs(5)!City
        sSQL(0) = "Select * from XBD WHERE ID=" & QT(txtID.Text)
        dbOpen (0): Set flxPersonal.DataSource = recs(0): dbClose (0)
    Else
        txtName.Text = ""
    End If
    dbClose (5)
End Sub

Private Sub cmdSelectID_Click()
    frmShow.Init "SELECT * FROM appview_ALLACCOUNTS"
    
    If sArray(0) <> "" Then
        txtID.Text = sArray(0): txtName.Text = sArray(1)
        sSQL(0) = "Select * from XBD WHERE ID=" & QT(txtID.Text)
        dbOpen (0): Set flxPersonal.DataSource = recs(0): dbClose (0)
    End If
End Sub

Private Sub cmdReturnDiscountTest_Click()
        For i = 1 To flxPub.ROWS - 1
            flxPub.TextMatrix(i, 4) = "0.00"
            DoEvents
        Next
        
        For i = 1 To flxPub.ROWS - 1
            'PARTY
            Party = txtID.Text
            PublisherID = flxPub.TextMatrix(i, 0)
            PDiscount = flxPub.TextMatrix(i, 2)
            SDiscount = flxPub.TextMatrix(i, 3)
            
            sSQL(6) = "appproc_ReturnDiscount " & QT(Party) & ", " & QT(PublisherID) & ", " & QT(PDiscount) & ", " & QT(SDiscount) & ", " & QT(Opt)
            dbOpen (6)
            If Not recs(6).EOF Then flxPub.TextMatrix(i, 4) = recs(6)!Discount
            dbClose (6)
            flxPub.ShowCell i, 4
            DoEvents
        Next
End Sub

Private Sub flxPub_EnterCell()
    If flxPub.Col = 4 Then
        flxPub.Editable = flexEDKbdMouse
    Else
        flxPub.Editable = flexEDNone
    End If
    sSQL(20) = "SELECT ITEMID, ISBN, ITEMNAME, PUBLISHERID AS PUBID, PUBLISHERNAME, PDISCOUNT, SDISCOUNT, PRICE FROM ITEMS WHERE PUBLISHERID=" & QT(flxPub.TextMatrix(flxPub.Row, 0)) & " AND PDISCOUNT=" & QT(flxPub.TextMatrix(flxPub.Row, 2)) & " AND SDISCOUNT=" & QT(flxPub.TextMatrix(flxPub.Row, 3))
    dbOpen (20): Set flxItems.DataSource = recs(20): dbClose (20)
End Sub

Private Sub flxPub_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If HasAccount(txtID.Text) = True Then
        sSQL(0) = "appproc_SaveDiscount " & QT(txtID.Text)
        sSQL(0) = sSQL(0) & " ," & QT(flxPub.TextMatrix(Row, 0))
        sSQL(0) = sSQL(0) & " ," & QT(flxPub.TextMatrix(Row, 1))
        sSQL(0) = sSQL(0) & " ," & QT(flxPub.TextMatrix(Row, 2))
        sSQL(0) = sSQL(0) & " ," & QT(flxPub.TextMatrix(Row, 3))
        sSQL(0) = sSQL(0) & " ," & QT(flxPub.TextMatrix(Row, 4))
        dbOpen (0): dbClose (0)
        sSQL(0) = "Select * from XBD WHERE ID=" & QT(txtID.Text)
        dbOpen (0): Set flxPersonal.DataSource = recs(0): dbClose (0)
    End If
End Sub
