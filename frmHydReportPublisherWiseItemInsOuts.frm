VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHydReportPublisherWiseItemInsOuts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report PublisherWise PeriodWise Item Ins / Outs "
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   12465
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   12405
      TabIndex        =   0
      Top             =   0
      Width           =   12465
      Begin VB.ComboBox cmbPersonal 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmHydReportPublisherWiseItemInsOuts.frx":0000
         Left            =   6150
         List            =   "frmHydReportPublisherWiseItemInsOuts.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "CUSTOMER"
         Top             =   15
         Width           =   1920
      End
      Begin VB.ComboBox cmbCustomer 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmHydReportPublisherWiseItemInsOuts.frx":0004
         Left            =   8085
         List            =   "frmHydReportPublisherWiseItemInsOuts.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   10
         Text            =   "cmbCustomer"
         Top             =   15
         Width           =   4320
      End
      Begin VB.CheckBox chkQtyAmountWise 
         Caption         =   "Qty_Wise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10290
         TabIndex        =   9
         Top             =   510
         Width           =   2100
      End
      Begin VB.Frame Frame1 
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   15
         TabIndex        =   4
         Top             =   345
         Width           =   5520
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   330
            Left            =   885
            TabIndex        =   7
            Top             =   165
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "ddd, dd-MMM-yy"
            Format          =   82378755
            CurrentDate     =   39407
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   330
            Left            =   3180
            TabIndex        =   8
            Top             =   165
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "ddd, dd-MMM-yy"
            Format          =   82378755
            CurrentDate     =   39407
         End
         Begin VB.Label Label3 
            Caption         =   "to :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2850
            TabIndex        =   6
            Top             =   225
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "From: "
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
            Left            =   255
            TabIndex        =   5
            Top             =   225
            Width           =   705
         End
      End
      Begin VB.ComboBox cmbPublisher 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmHydReportPublisherWiseItemInsOuts.frx":0008
         Left            =   885
         List            =   "frmHydReportPublisherWiseItemInsOuts.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cmbPublisher"
         Top             =   15
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Publisher:"
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
         Left            =   30
         TabIndex        =   3
         Top             =   45
         Width           =   1545
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid flxReport 
      Align           =   2  'Align Bottom
      Height          =   8310
      Left            =   0
      TabIndex        =   2
      Top             =   1050
      Width           =   12465
      _cx             =   21987
      _cy             =   14658
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      AutoSearch      =   2
      AutoSearchDelay =   3
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmHydReportPublisherWiseItemInsOuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(5) As New clsData
Dim opt As Integer

Private Sub Form_Load()
    mdiOne.SetFormFont Me
    Me.Move 0, 0
    
    opt = 0
    dtFrom.Value = "01-APR-2007"
    dtTo.Value = Date
    
    CN(0).dbOpen "SELECT DISTINCT PUBLISHERNAME FROM items ORDER BY 1", 1
    If Not CN(0).recs.EOF Then CN(0).recs.MoveFirst
    cmbPublisher.Clear: cmbPublisher.AddItem "  ALL"
    Do Until CN(0).recs.EOF
        cmbPublisher.AddItem CN(0).recs.FIELDS(0)
        CN(0).recs.MoveNext
    Loop
    cmbPublisher.ListIndex = 0
    
    cmbPersonal.Clear: cmbPersonal.AddItem "CUSTOMER": cmbPersonal.AddItem "DISTRIBUTOR": cmbPersonal.AddItem "WAREHOUSE"
    LoadPersonalData
    cmbPersonal.ListIndex = 0
    
    opt = 1
End Sub

Private Sub cmbpublisher_Click()
    LoadReport
End Sub

Private Sub cmbPersonal_Click()
    LoadPersonalData
    LoadReport
End Sub

Private Sub cmbCustomer_Click()
    LoadReport
End Sub

Private Sub dtFrom_Change()
    LoadReport
End Sub

Private Sub dtTo_Change()
    LoadReport
End Sub

Private Sub chkQtyAmountWise_Click()
    On Error Resume Next
    If chkQtyAmountWise.Value = 0 Then
        chkQtyAmountWise.Caption = "QTY_WISE"
    Else
        chkQtyAmountWise.Caption = "AMOUNT_WISE"
    End If
    LoadReport
End Sub

Private Sub LoadReport()
    On Error Resume Next
    If opt = 1 Then
        If chkQtyAmountWise.Value = 0 Then
            If Trim(cmbPublisher.Text) = "ALL" Then
                If Trim(cmbCustomer.Text) = "ALL" Then
                    SQ = "SELECT itemID, itemName, PublisherName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Qty_InOut_Main_HYD WHERE ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) GROUP BY itemID, itemNAME, PublisherName, PRICE ORDER BY 1,2,3"
                Else
                    SQ = "SELECT itemID, itemName, PublisherName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Qty_InOut_Main_HYD WHERE ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) AND NAME=" & QT(cmbCustomer.Text) & " GROUP BY itemID, itemNAME, PublisherName, PRICE ORDER BY 1,2,3"
                End If
            Else
                If Trim(cmbCustomer.Text) = "ALL" Then
                    SQ = "SELECT itemID, ISBN, itemName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Qty_InOut_Main_HYD WHERE PUBLISHERNAME= " & QT(cmbPublisher.Text) & " AND ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) GROUP BY itemID, ISBN, itemNAME, PRICE ORDER BY 1,2,3"
                Else
                    SQ = "SELECT itemID, ISBN, itemName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Qty_InOut_Main_HYD WHERE PUBLISHERNAME= " & QT(cmbPublisher.Text) & " AND ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) AND NAME=" & QT(cmbCustomer.Text) & " GROUP BY itemID, ISBN, itemNAME, PRICE ORDER BY 1,2,3"
                End If
            End If
        Else
            If Trim(cmbPublisher.Text) = "ALL" Then
                If Trim(cmbCustomer.Text) = "ALL" Then
                    SQ = "SELECT itemID, itemName, PublisherName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Amount_InOut_Main_HYD WHERE ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) GROUP BY itemID, itemNAME, PublisherName, PRICE ORDER BY 1,2,3"
                Else
                    SQ = "SELECT itemID, itemName, PublisherName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Amount_InOut_Main_HYD WHERE ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) AND NAME=" & QT(cmbCustomer.Text) & " GROUP BY itemID, itemNAME, PublisherName, PRICE ORDER BY 1,2,3"
                End If
            Else
                If Trim(cmbCustomer.Text) = "ALL" Then
                    SQ = "SELECT itemID, ISBN, itemName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Amount_InOut_Main_HYD WHERE PUBLISHERNAME= " & QT(cmbPublisher.Text) & " AND ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) GROUP BY itemID, ISBN, itemNAME, PRICE ORDER BY 1,2,3"
                Else
                    SQ = "SELECT itemID, ISBN, itemName, Price, SUM(PURC) as PURC, SUM(PRTN) as PRTN, SUM(T_IN) as T_IN, SUM(TOUT) as TOUT, SUM(SALE) as SALE, SUM(SRTN) as SRTN FROM appview_Report_PublisherWise_Amount_InOut_Main_HYD WHERE PUBLISHERNAME= " & QT(cmbPublisher.Text) & " AND ( DT BETWEEN " & QT(dtFrom.Value) & " AND " & QT(dtTo.Value) & " ) AND NAME=" & QT(cmbCustomer.Text) & " GROUP BY itemID, ISBN, itemNAME, PRICE ORDER BY 1,2,3"
                End If
            End If
        End If
        flxReport.Clear
        CN(1).dbOpen SQ, 1
        Set flxReport.DataSource = CN(1)
        flxReport.SubTotal flexSTSum, -1, 2, "#,##0", , vbBlue, True, "Total"
        flxReport.SubTotal flexSTSum, -1, 3, "#,##0", , vbBlue, True, "Total"
        flxReport.SubTotal flexSTSum, -1, 4, "#,##0", , vbBlue, True, "Total"
        flxReport.SubTotal flexSTSum, -1, 5, "#,##0", , vbBlue, True, "Total"
        flxReport.SubTotal flexSTSum, -1, 6, "#,##0", , vbBlue, True, "Total"
        flxReport.SubTotal flexSTSum, -1, 7, "#,##0", , vbBlue, True, "Total"
        flxReport.SubTotal flexSTSum, -1, 8, "#,##0", , vbBlue, True, "Total"
        flxReport.SubTotal flexSTSum, -1, 9, "#,##0", , vbBlue, True, "Total"
        
        flxReport.AutoSize 0, flxReport.COLS - 1
        flxReport.ShowCell flxReport.ROWS - 1, 0
    End If
End Sub

Private Sub flxReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 And Shift = vbCtrlMask Then SaveGrid Me.flxReport
    Me.Caption = FlxSum(Me.flxReport)
End Sub

Private Sub LoadPersonalData()
    On Error Resume Next
    Select Case cmbPersonal.Text
        Case "CUSTOMER": CN(0).dbOpen "SELECT NAME FROM PERSONAL WHERE ID like " & QT("[C]%") & " ORDER BY 1", 1
        Case "DISTRIBUTOR": CN(0).dbOpen "SELECT NAME FROM PERSONAL WHERE ID like " & QT("[D]%") & " ORDER BY 1", 1
        Case "WAREHOUSE": CN(0).dbOpen "SELECT NAME FROM PERSONAL WHERE ID like " & QT("[W]%") & " ORDER BY 1", 1
    End Select
        
    If Not CN(0).recs.EOF Then CN(0).recs.MoveFirst
    cmbCustomer.Clear: cmbCustomer.AddItem "  ALL"
    Do Until CN(0).recs.EOF
        cmbCustomer.AddItem CN(0).recs.FIELDS(0)
        CN(0).recs.MoveNext
    Loop
    cmbCustomer.ListIndex = 0
End Sub

