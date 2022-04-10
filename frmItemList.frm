VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmItemList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report: Items subject wise..."
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "frmItemList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   9510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   16775
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Grid..."
      TabPicture(0)   =   "frmItemList.frx":114DA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxItem"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "PicB"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Preview..."
      TabPicture(1)   =   "frmItemList.frx":114F6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox PicB 
         Height          =   1395
         Left            =   60
         ScaleHeight     =   1335
         ScaleWidth      =   8625
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   8040
         Width           =   8685
         Begin VB.TextBox txtListHeading 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1110
            TabIndex        =   19
            Top             =   660
            Width           =   7515
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
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
            Left            =   0
            TabIndex        =   18
            Top             =   1020
            Width           =   1725
         End
         Begin VB.CheckBox chkStockFilter 
            Alignment       =   1  'Right Justify
            Caption         =   "Exclude Zero Stock Items"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   7125
            TabIndex        =   17
            Top             =   30
            Width           =   1380
         End
         Begin VB.TextBox txtSubjectHeading 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1110
            TabIndex        =   13
            Top             =   330
            Width           =   3495
         End
         Begin VB.TextBox txtSubject 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1110
            TabIndex        =   12
            ToolTipText     =   "Sub1, Sub2,... & YoP1, YoP2, ..."
            Top             =   0
            Width           =   5730
         End
         Begin VB.CommandButton cmdSaveGrid 
            Caption         =   "SaveGrid"
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
            Left            =   1725
            TabIndex        =   11
            Top             =   1020
            Width           =   1725
         End
         Begin VB.CommandButton cmdRender 
            Caption         =   "&Render"
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
            Left            =   5175
            TabIndex        =   10
            Top             =   1020
            Width           =   1725
         End
         Begin VB.CommandButton cmdDeleteItem 
            Caption         =   "&DeleteRow"
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
            Left            =   3450
            TabIndex        =   9
            Top             =   1020
            Width           =   1725
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
            Left            =   6900
            TabIndex        =   8
            Top             =   1020
            Width           =   1725
         End
         Begin VB.ComboBox cmbList 
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
            Left            =   4620
            TabIndex        =   7
            Text            =   "Combo1"
            Top             =   330
            Width           =   2220
         End
         Begin VB.Label lblSubject 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "List Heading:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   30
            TabIndex        =   20
            Top             =   690
            Width           =   1080
         End
         Begin VB.Label lblSubject 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "List Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   30
            TabIndex        =   16
            Top             =   345
            Width           =   1080
         End
         Begin VB.Label lblSubject 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subs && YoPs: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   30
            TabIndex        =   14
            Top             =   30
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Height          =   9075
         Left            =   -74940
         TabIndex        =   1
         Top             =   360
         Width           =   8670
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   5670
            ScaleHeight     =   270
            ScaleWidth      =   2940
            TabIndex        =   2
            Top             =   120
            Width           =   2940
            Begin VB.CommandButton cmdChangePageOrientation 
               Caption         =   "Orientation"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1515
               TabIndex        =   3
               Top             =   0
               Width           =   1425
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
               Height          =   270
               Left            =   105
               TabIndex        =   4
               Top             =   0
               Width           =   1425
            End
         End
         Begin VSPrinter8LibCtl.VSPrinter vp 
            Height          =   8880
            Left            =   60
            TabIndex        =   5
            Top             =   135
            Width           =   8550
            _cx             =   15081
            _cy             =   15663
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            MousePointer    =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
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
            PhysicalPage    =   0   'False
            PalettePicture  =   "frmItemList.frx":11512
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
            Zoom            =   64
            ZoomMode        =   0
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
            NavBar          =   1
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
      Begin VSFlex8UCtl.VSFlexGrid flxItem 
         Height          =   7620
         Left            =   60
         TabIndex        =   15
         Top             =   390
         Width           =   8685
         _cx             =   15319
         _cy             =   13441
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
         BackColorFixed  =   8454143
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
         AllowUserResizing=   3
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
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
End
Attribute VB_Name = "frmItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New clsData

Dim IsPageOrientationPortrait As Boolean
Dim itemListSQL, Filter As String

Private Sub Form_Load()
    Me.Move 0, 0
    mdiOne.SetFormFont Me
    flxItem.FontName = "Arial Unicode MS": flxItem.FontSize = 8
    SSTab1.Tab = 0
    IsPageOrientationPortrait = True
    cmbList.AddItem "item List [PubID]"
    cmbList.AddItem "item List [PubName]"
    cmbList.AddItem "item List [PubID] YoP"
    cmbList.AddItem "item List [PubName] YoP"
    cmbList.AddItem "Extended item List"
    cmbList.ListIndex = 0
End Sub

Private Sub flxItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton And Shift = vbCtrlMask Then
        SaveGrid Me.flxItem
    End If
End Sub

Private Sub flxItem_AfterSort(ByVal Col As Long, Order As Integer)
    Enumerate
End Sub

Private Sub flxItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    flxItem.Visible = False
    flxItem.AutoSize 0, flxItem.COLS - 1, False, 40
    flxItem.Visible = True
End Sub

Private Sub txtSubject_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRefresh_Click
    Else
        CreateFilter
    End If
End Sub

Private Sub cmbList_Click()
    Select Case cmbList.ListIndex
        Case 0: IsPageOrientationPortrait = True: itemListSQL = "SELECT * FROM appview_itemListShortPubID"
        Case 1: IsPageOrientationPortrait = True: itemListSQL = "SELECT * FROM appview_itemListShortPubName"
        Case 2: IsPageOrientationPortrait = True: itemListSQL = "SELECT * FROM appview_itemListShortPubIDYoP"
        Case 3: IsPageOrientationPortrait = True: itemListSQL = "SELECT * FROM appview_itemListShortPubNameYoP"
        Case 4: IsPageOrientationPortrait = False: itemListSQL = "SELECT * FROM appview_itemList"
    End Select
    cmdRefresh_Click
End Sub

Private Sub chkStockFilter_Click()
    CreateFilter
End Sub

Private Sub cmdRefresh_Click()
    CreateFilter
    flxItem.Visible = False
    Set flxItem.DataSource = Nothing: flxItem.FixedCols = 0: flxItem.FixedRows = 1
    CN.dbOpen itemListSQL & Filter, 1
    Set flxItem.DataSource = CN.recs
    Enumerate
    flxItem.Visible = True
End Sub

Private Sub cmdSaveGrid_Click()
    If MsgBox("Do you wish to save the grid?", vbYesNo + vbQuestion) = vbYes Then
        mdiOne.CDlg.FileName = txtSubjectHeading.Text
        mdiOne.CDlg.Filter = CompanyName & " Excel Report |*.xls"
        mdiOne.CDlg.ShowSave
        If mdiOne.CDlg.CancelError = False Then flxItem.SaveGrid mdiOne.CDlg.FileName, flexFileExcel, flexXLSaveFixedRows
    End If
End Sub

Private Sub cmdDeleteitem_Click()
    If MsgBox("Confirm deletion?", vbYesNo + vbQuestion) = vbYes Then
        For i = flxItem.SelectedRows - 1 To 0 Step -1
            R = flxItem.SelectedRow(i)
            If flxItem.ROWS > 1 Then flxItem.RemoveItem R
        Next
        Enumerate
    End If
End Sub

Private Sub cmdRender_Click()
    ApplyColumnSettings
    RenderReport
    SSTab1.Tab = 1
End Sub

Private Sub cmdChangePageOrientation_Click()
    IsPageOrientationPortrait = Not IsPageOrientationPortrait
    RenderReport
End Sub

Private Sub cmdPrint_Click()
    vp.Visible = False: vp.PrintDoc True: vp.Visible = True
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub


Private Sub ApplyColumnSettings()
    flxItem.Visible = False: flxItem.BackColorFixed = vbWhite
    flxItem.Clear flexClearEverywhere, flexClearFormatting
    Select Case cmbList.ListIndex
        Case 0, 1:               ' GENERAL LIST Portrait
            flxItem.ColWidth(0) = 600: flxItem.ColAlignment(0) = flexAlignRightCenter
            flxItem.ColWidth(1) = 2000
            flxItem.ColWidth(2) = 2000
            flxItem.ColWidth(3) = 5000
            flxItem.ColWidth(4) = 2000
            flxItem.ColWidth(5) = 500: flxItem.ColAlignment(5) = flexAlignRightCenter
            flxItem.ColWidth(6) = 800: flxItem.ColFormat(5) = "#.00"
            flxItem.ColWidth(7) = 10000
            flxItem.AutoSize 0, flxItem.COLS - 1, False, 40
        Case 2, 3:               ' GENERAL LIST YoP Portrait
            flxItem.ColWidth(0) = 600: flxItem.ColAlignment(0) = flexAlignRightCenter
            flxItem.ColWidth(1) = 2000
            flxItem.ColWidth(2) = 2000
            flxItem.ColWidth(3) = 4700
            flxItem.ColWidth(4) = 600
            flxItem.ColWidth(5) = 1700: flxItem.ColAlignment(5) = flexAlignLeftCenter
            flxItem.ColWidth(6) = 500: flxItem.ColAlignment(6) = flexAlignRightCenter
            flxItem.ColWidth(7) = 800: flxItem.ColFormat(7) = "#.00"
            flxItem.ColWidth(8) = 10000
            flxItem.AutoSize 0, flxItem.COLS - 1, False, 40
        Case Else
            flxItem.ColWidth(0) = 600
            flxItem.ColWidth(1) = 1250
            flxItem.ColWidth(2) = 5250
            flxItem.ColWidth(3) = 2000
            flxItem.ColWidth(4) = 600
            flxItem.ColWidth(5) = 600
            flxItem.ColWidth(6) = 1800
            flxItem.ColWidth(7) = 500
            flxItem.ColWidth(8) = 500
            flxItem.ColWidth(9) = 800
            flxItem.ColWidth(10) = 1600
            flxItem.ColWidth(11) = 1600
            flxItem.ColWidth(12) = 500
            flxItem.ColWidth(13) = 500
            flxItem.AutoSize 0, flxItem.COLS - 5, False, 40
    End Select
    flxItem.Visible = True
End Sub

Private Sub RenderReport()
    vp.PaperSize = pprFanfoldStdGerman
    vp.MarginLeft = 400: vp.MarginRight = 50
    vp.MarginTop = 350: vp.MarginBottom = 100
    RenderBody
    RenderOverlay
End Sub

Private Sub RenderBody()
    With vp
        If IsPageOrientationPortrait Then
            .Orientation = orPortrait
        Else
            .Orientation = orLandscape
        End If
        .StartDoc
            .TextAlign = taCenterTop
            SetFont ("[Render_Section_A]")
            .Text = "item LIST"
            .CurrentY = 700
            .TextAlign = taLeftTop
            .DrawPicture mdiOne.ImgList.ListImages(1).Picture, 500, .CurrentY, 837, 1100
            SetFont ("[Render_Section_B]")
            .StartTable
                .TableBorder = tbNone
                .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = "0.7in"
                .TableCell(tcColWidth, , 2) = "5.0in"
                .TableCell(tcColAlign, , 2) = taLeftTop
                .TableCell(tcText, 1, 2) = CompanyName
            .EndTable
            SetFont ("[Render_Section_C]")
            .StartTable
                .TableBorder = tbNone
                .TableCell(tcCols) = 2: .TableCell(tcRows) = 1
                .TableCell(tcColWidth, , 1) = "0.7in"
                .TableCell(tcColWidth, , 2) = "5.0in"
                .TableCell(tcColAlign, , 2) = taLeftTop
                .TableCell(tcText, 1, 2) = AboutCompany & vbCrLf & CompanyAddr0 & vbCrLf
            .EndTable
            SetFont ("[Render_Section_D]")
            msg = txtListHeading.Text
            .TextAlign = taCenterTop
            .Text = msg
            .CurrentY = .CurrentY + 300
            .TextAlign = taLeftTop
            
            flxItem.ColHidden(1) = True
            flxItem.ColHidden(flxItem.COLS - 3) = True
            flxItem.ColHidden(flxItem.COLS - 2) = True
            flxItem.ColHidden(flxItem.COLS - 1) = True
            .RenderControl = flxItem.hWnd
            flxItem.ColHidden(1) = False
            flxItem.ColHidden(flxItem.COLS - 3) = False
            flxItem.ColHidden(flxItem.COLS - 2) = False
            flxItem.ColHidden(flxItem.COLS - 1) = False
        .EndDoc
    End With
End Sub
    
Public Sub RenderOverlay()
    With vp
        SetFont ("[Render_Section_Z]")
        For i = 1 To vp.PageCount
            vp.StartOverlay i: vp.CurrentX = vp.PageWidth - 2000: vp.CurrentY = 100
            vp.Text = "Page " & i & " of " & vp.PageCount
            vp.EndOverlay
        Next
    End With
End Sub

Private Sub vp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 188 And Shift = vbCtrlMask Then  '<
        vp.MarginTop = vp.MarginTop + 150
        vp.MarginBottom = vp.MarginBottom + 150
        RenderReport
    End If
    If KeyCode = 190 And Shift = vbCtrlMask Then  '>
        vp.MarginTop = vp.MarginTop - 150
        vp.MarginBottom = vp.MarginBottom - 150
        RenderReport
    End If
    vp.SetFocus
End Sub

Private Sub CreateFilter()
    txtSubject.ToolTipText = "Sub1, Sub2,... & YoP1, YoP2, ..."
    Filter = ""
    If txtSubject <> "" And InStr(1, txtSubject.Text, "&", vbTextCompare) Then
        sy = Split(txtSubject.Text, "&"): Subjects = Split(sy(0), ","): YoPs = Split(sy(1), ",")
        For i = 0 To UBound(Subjects)
            Subjects(i) = " Subjects Like " & QT("%" & Trim(Subjects(i)) & "%") & "  "
        Next
        For i = 0 To UBound(YoPs)
            YoPs(i) = " YoP=" & QT(Trim(YoPs(i))) & "  "
        Next
        
        subfilter = Join(Subjects, " OR "): If InStr(1, subfilter, "*", vbTextCompare) Then subfilter = ""
        yopfilter = Join(YoPs, " OR "): If InStr(1, yopfilter, "*", vbTextCompare) Then yopfilter = ""
        
        Fiter = ""
        If subfilter <> "" And yopfilter <> "" Then Filter = " WHERE (" & subfilter & ") AND (" & yopfilter & ")"
        If subfilter = "" Xor yopfilter = "" Then Filter = " WHERE (" & subfilter & yopfilter & ")"
        
        If subfilter = "" Then
            txtSubjectHeading = "All Subjects"
        Else
            txtSubjectHeading = UCase(sy(0))
        End If
    
        If chkStockFilter.Value <> 0 Then
            If Filter = "" Then
                Filter = " WHERE (STOCK > 0)"
            Else
                Filter = Filter & " AND (STOCK > 0)"
            End If
        End If
    End If
    Me.Caption = "item List: With filter as " & Filter
    txtListHeading.Text = "New Arrivals " & txtSubjectHeading.Text & " item List from " & Year(Now)
End Sub

Private Sub Enumerate()
    For i = 1 To flxItem.ROWS - (flxItem.FixedRows)
        flxItem.TextMatrix(i + flxItem.FixedRows - 1, 0) = i
    Next
End Sub

Private Sub SetFont(S As String)
    vp.FontName = ReadFont(S, 0)
    vp.FontSize = ReadFont(S, 1)
    vp.FontBold = ReadFont(S, 2)
    vp.FontItalic = ReadFont(S, 3)
End Sub

