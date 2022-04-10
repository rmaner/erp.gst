VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchaseReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public frm As New frmPPSS

Private Sub Class_Initialize()
    frm.MyCAPTION = "PurchaseReturn "
    frm.StockValidationRequired = False
    frm.HiddenCols = "[PPSS-HiddenCols]"
    frm.Main = "PRETURNMAIN"
    frm.Support = "PURCHASERETURN"
    
    frm.CounterMain = "PMAIN"
    frm.CounterSupport = "PURCHASE"
    
    frm.MainAddNew = "ADDNEW_PRETURNMain"
    frm.MainSelectView = "appview_PRETURNMain_Select_View"
    frm.PartyView = "appview_AllAccounts"
    frm.PartyInitial = "[D]%"
    frm.ItemSelectView = "appview_SelectitemPurchase"
    frm.Discs = True
    frm.AutoCalculate = True
    frm.MemoFormat = "\P\R0"
End Sub

Public Sub ShowForm()
    frm.Show
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub

