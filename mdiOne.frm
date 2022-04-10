VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{99A6C39B-454A-45C9-979A-E51215BC1B64}#1.1#0"; "uitsBase2.ocx"
Begin VB.MDIForm mdiOne 
   BackColor       =   &H8000000C&
   Caption         =   "UITS..."
   ClientHeight    =   8250
   ClientLeft      =   1740
   ClientTop       =   1860
   ClientWidth     =   16335
   Icon            =   "mdiOne.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picLogo 
      Align           =   3  'Align Left
      AutoSize        =   -1  'True
      Height          =   7995
      Left            =   0
      ScaleHeight     =   7935
      ScaleWidth      =   990
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
      Begin uitsBase.sckGo sckGo 
         Height          =   675
         Left            =   60
         Top             =   3270
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1191
      End
      Begin VB.Timer timerMDI 
         Interval        =   1000
         Left            =   0
         Top             =   2280
      End
      Begin MSMAPI.MAPIMessages MAPIMsg 
         Left            =   0
         Top             =   1710
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession MAPISess 
         Left            =   0
         Top             =   1140
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin MSCommLib.MSComm MyMSComm1 
         Left            =   0
         Top             =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSComctlLib.ImageList ImgList 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiOne.frx":114DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiOne.frx":118D0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CDlg 
         Left            =   0
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7995
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18018
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "5/19/2018"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "1:25 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnutest 
      Caption         =   "test"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLogin 
         Caption         =   "Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuFileDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileChangePassword 
         Caption         =   "ChangePassword"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Admin"
      Begin VB.Menu mnuAdminDBConnection 
         Caption         =   "DBConnection"
      End
      Begin VB.Menu mnuAdminUserInformations 
         Caption         =   "UserInformations"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAdminDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminEditTables 
         Caption         =   "&EditTables"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAdminDiscountMap 
         Caption         =   "DiscountMap"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "T&ransaction"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuTransactionItemS 
         Caption         =   "Items"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnuTransactionQuickItem 
         Caption         =   "QuickItem"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTransactionDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransactionSale 
         Caption         =   "Sale"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuTransactionSaleReturn 
         Caption         =   "SaleReturn"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuTransactionDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransactionPurchase 
         Caption         =   "Purchase"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuTransactionPurchaseReturn 
         Caption         =   "PurchaseReturn"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuTransactionDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransactionTOUT 
         Caption         =   "StockTransferOUT"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuTransactionTIN 
         Caption         =   "StockTransferIN"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuTransactionDash04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminUnifiedEntry 
         Caption         =   "UnifiedEntry"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuAdminCurrencyRates 
         Caption         =   "Currency Rates"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "Accounts"
      Begin VB.Menu mnuAccountsPymt 
         Caption         =   "P&ymt"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAccountsRcpt 
         Caption         =   "&Rcpt"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuAccountsVouchers 
         Caption         =   "&Vouchers"
      End
      Begin VB.Menu mnuAccountSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountsDebitNote 
         Caption         =   "DebitNote"
      End
      Begin VB.Menu mnuAccountsCreditNote 
         Caption         =   "CreditNote"
      End
      Begin VB.Menu mnuAccountSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountsLedger 
         Caption         =   "&Ledger"
      End
      Begin VB.Menu mnuAccountSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountsAccountSummary 
         Caption         =   "AccountSummary"
      End
   End
   Begin VB.Menu mnuReportS 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsDayBook 
         Caption         =   "DayBook"
      End
      Begin VB.Menu mnuHydReports 
         Caption         =   "HydReports"
         Begin VB.Menu mnuHydReportPublisherWiseItemInsOuts 
            Caption         =   "HydReportPublisherWiseItemInsOuts"
         End
         Begin VB.Menu mnuHydReportSearchBills 
            Caption         =   "HydReportSearchBills"
         End
         Begin VB.Menu mnuHydSaveDailyReports 
            Caption         =   "HydSaveDailyReports"
         End
      End
      Begin VB.Menu mnuReportsGeneralReports 
         Caption         =   "General Reports"
      End
      Begin VB.Menu mnuReportsDateReports 
         Caption         =   "Date Reports"
      End
      Begin VB.Menu mnuReportsPeriodReports 
         Caption         =   "Period Reports"
      End
      Begin VB.Menu mnuReportsCustomReports 
         Caption         =   "Custom Reports"
      End
      Begin VB.Menu mnuReportsStock 
         Caption         =   "Stock Report"
      End
      Begin VB.Menu mnuReportSaleAndStockHolding 
         Caption         =   "SaleAndStockHolding"
      End
      Begin VB.Menu mnuProfitReport 
         Caption         =   "Profit Report"
      End
      Begin VB.Menu mnuReportsDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportsDuplicateItems 
         Caption         =   "DuplicateItems"
      End
      Begin VB.Menu mnuReportsItemsList 
         Caption         =   "ItemsList"
      End
      Begin VB.Menu mnuReportsISBNList 
         Caption         =   "ISBNList"
      End
   End
   Begin VB.Menu mnuStaff 
      Caption         =   "St&aff"
      Visible         =   0   'False
      Begin VB.Menu mnuStaffAttendance 
         Caption         =   "Attendance"
      End
      Begin VB.Menu mnuStaffPaySlip 
         Caption         =   "PaySlip"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      WindowList      =   -1  'True
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuHelpCalculator 
         Caption         =   "Calculator"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu mnuHelpPhoneDialer 
         Caption         =   "PhoneDialer"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpUDPConnect 
         Caption         =   "UDPConnect"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuReCreateJournal 
      Caption         =   "ReCreateJournal"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdiOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SuccessfulLoad As Integer
Dim hCalculator As Double
Dim SecondsElapsed As Double

Private Sub MDIForm_Load()
    On Error Resume Next
    mdiOne.sckGo.SetParentPath App.Path
    
    Call menuCLOSE
    DateChanged = False
    GetWorkingDay
    startTime = Now
    mdiOne.Caption = CompanyName & ": Items Inventory System " & Format(Date, "Long Date") & ")"
    mdiOne.Picture = LoadPicture("uits.gif")
    Me.CDlg.FontName = ReadFont("[APPLICATION_FONT]", 0)
    Me.CDlg.FontSize = ReadFont("[APPLICATION_FONT]", 1)
    Me.CDlg.FontBold = ReadFont("[APPLICATION_FONT]", 2)
    Me.CDlg.FontItalic = ReadFont("[APPLICATION_FONT]", 3)
    Me.SetFormFont Me.ActiveForm
    
    Set MSComm1 = MyMSComm1
    For Each i In Split(mdiOne.sckGo.GReadINI("[StartUp]", "[END]"), "(:-)")
        sSQL(0) = i: dbOpen (0): dbClose (0)
    Next
    hCalculator = 0
    OutBoundRule
    'mnuFileLogin_Click
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
        frmTest.Show
End Sub

Private Sub timerMDI_Timer()
    mdiOne.Caption = CompanyName & "; " & CompanyState & " [ " & initCatalog & "@" & DatabaseServer & " ] " & " [ Items Inventory System ] " & " [ " & Format(Now, "ddd, DD-MMM-YYYY HH:MM:SS") & " ]"
End Sub

Private Sub MDIForm_DblClick()
    CDlg.ShowFont
    Me.SetFormFont Me.ActiveForm
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'Close & unset connection
    Unload frmShow
    If IsEmpty(conn) Then conn.Close
    Set conn = Nothing
End Sub

Private Sub mnuProfitReport_Click()
    frmReportProfit.Show
End Sub

Private Sub mnuReportsDayBook_Click()
    frmReportsDayBook.Show
End Sub

Private Sub mnutest_Click()
    frmTest.Show
    'frmGeneralLedgerAccounts.Show
End Sub


'======================= FILE MENUS
Private Sub mnuFileLogin_Click()
    frmLoginVista.Show vbModal
    frmTime.Show
End Sub

Private Sub mnuFileLogout_Click()
    ConnectToDatabase "master", DatabaseServer
    Call menuCLOSE
    frmTime.Show
End Sub

Private Sub mnuFileChangePassword_Click()
    frmChangePassword.Show
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'======================= ADMIN MENUS
Private Sub mnuAdminDBConnection_Click()
    'frmDatabaseConnection.Show vbModal
End Sub

Private Sub mnuAdminUserInformations_Click()
    frmUserInformations.Show
End Sub

Private Sub mnuAdminUnifiedEntry_Click()
    UnifiedEntry = "GENERAL"
    frmEntry8.Show
End Sub

Private Sub mnuAdminEditTables_Click()
    frmEditTable.Show
End Sub

Private Sub mnuAdminCurrencyRates_Click()
    frmCurrency.Show
End Sub

Private Sub mnuAdminDiscountMap_Click()
    frmDiscountMap.Show
End Sub

'======================= TRANSACTION MENUS
Private Sub mnuTransactionitems_Click()
    frmItems.Show
End Sub

Private Sub mnuTransactionQuickitem_Click()
    frmQuickItem.Show
End Sub


Private Sub mnuTransactionSale_Click()
    Sale.ShowForm
End Sub

Private Sub mnuTransactionSaleReturn_Click()
    SaleReturn.ShowForm
End Sub

Private Sub mnuTransactionPurchase_Click()
    Purchase.ShowForm
End Sub

Private Sub mnuTransactionPurchaseReturn_Click()
    PurchaseReturn.ShowForm
End Sub

Private Sub mnuTransactionTIN_Click()
    StockTransferIN.ShowForm
End Sub

Private Sub mnuTransactionTOUT_Click()
    StockTransferOUT.ShowForm
End Sub

'======================= ACCOUNTS MENUS
Private Sub mnuAccountsPYMT_Click()
    PYMT.ShowForm
End Sub

Private Sub mnuAccountsRCPT_Click()
    RCPT.ShowForm
End Sub

Private Sub mnuAccountsVouchers_Click()
    frmVouchers.Show
End Sub

Private Sub mnuAccountsCreditNote_Click()
    'TRF.ShowForm
End Sub

Private Sub mnuAccountsDebitNote_Click()
    'TRF.ShowForm
End Sub

Private Sub mnuAccountsLedger_Click()
    frmLedger.Show
End Sub

Private Sub mnuAccountsAccountSummary_Click()
    frmAccountSummary.Show
End Sub

'======================= REPORTS MENUS
Private Sub mnuHydReportPublisherWiseitemInsOuts_Click()
    frmHydReportPublisherWiseItemInsOuts.Show
End Sub

Private Sub mnuHydReportSearchBills_Click()
    frmHydSearchBills.Show
End Sub

Private Sub mnuHydSaveDailyReports_Click()
    frmHydSaveDailyReports.Show
End Sub


Private Sub mnuReportsGeneralReports_Click()
    frmReportsGen.Show
End Sub

Private Sub mnuReportsDateReports_Click()
    frmReportsDate.Show
End Sub

Private Sub mnuReportsPeriodReports_Click()
    frmReportsPeriod.Show
End Sub

Private Sub mnuReportsCustomReports_Click()
    frmReportsCustom.Show
End Sub

Private Sub mnuReportsStock_Click()
    frmReportStock.Show
End Sub

Private Sub mnuReportSaleAndStockHolding_Click()
    frmSaleAndStockHolding.Show
End Sub

Private Sub mnuReportsDuplicateitems_Click()
    frmReportDuplicateItems.Show
End Sub

Private Sub mnuReportsitemsList_Click()
    frmItemList.Show
End Sub

Private Sub mnuReportsISBNList_Click()
    frmISBNList.Show
End Sub

'======================= STAFF MENUS
Private Sub mnuStaffAttendance_Click()

End Sub

'======================= HELP MENUS
Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpCalculator_Click()
    hCalculator = Shell("Calc.exe")
End Sub

Private Sub mnuHelpUDPConnect_Click()
    MSN.ShowForm
End Sub

Private Sub mnuHelpContents_Click()
    frmHelpContents.Show
End Sub

'======================= CREATE JOURNALENTRY MENU
Private Sub mnuReCreateJournal_Click()
    ReCreateJournalEntries
End Sub


'==============================================
'=============== OPEN/CLOSE ===================
'==============================================

Public Sub menuOPEN(ByVal Rights As Integer)
    Select Case Rights
        Case 0  'Administrator
            mnuAdmin.Enabled = True: mnuAdminDBConnection.Enabled = True: mnuAdminUserInformations.Enabled = True: mnuAdminEditTables.Enabled = True
            mnuTransaction.Enabled = True
            mnuAccounts.Enabled = True
            mnuReportS.Enabled = True
        Case 1  'SuperUser
            mnuAdmin.Enabled = True: mnuAdminDBConnection.Enabled = False: mnuAdminUserInformations.Enabled = False: mnuAdminEditTables.Enabled = False
            mnuTransaction.Enabled = True
            mnuAccounts.Enabled = True
            mnuReportS.Enabled = True
        Case 2  'General User
            mnuAdmin.Enabled = False
            mnuTransaction.Enabled = True
            mnuReportS.Enabled = True
            mnuAccounts.Enabled = True
        Case 3  'Report Viewer
            mnuAdmin.Enabled = False
            mnuTransaction.Enabled = False
            mnuReportS.Enabled = True
            mnuAccounts.Enabled = False
        Case 4  'Accounts
            mnuAdmin.Enabled = False
            mnuTransaction.Enabled = False
            mnuReportS.Enabled = True
            mnuAccounts.Enabled = True
    End Select
End Sub

Public Sub menuCLOSE()
    mnuAdmin.Enabled = False
    mnuTransaction.Enabled = False
    mnuAccounts.Enabled = False
    mnuReportS.Enabled = False
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Do you wish to quit?", vbOKCancel + vbQuestion) = vbOK Then
        If DateChanged = True Then
            MsgBox "Date is reverting back to " & RealDate, vbOKOnly + vbExclamation
            Date = RealDate
            Date = DateAdd("s", SecondsElapsed, Date)
        End If
    Else
        Cancel = 1
    End If
End Sub

Public Sub SetFormFont(F As Form, Optional ByVal FntName As Variant, Optional ByVal FntSize As Variant)
    On Error Resume Next
    Dim A As Control
    For Each A In F.Controls
        If IsMissing(FntName) Then
            A.FontName = CDlg.FontName
            A.FontSize = CDlg.FontSize
        Else
            A.FontName = FntName
            If Not IsMissing(FntSize) Then A.FontSize = FntSize
        End If
    Next
End Sub
