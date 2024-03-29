VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public frm As New frmPPSS

Private Sub Class_Initialize()
    frm.MyCAPTION = "Sale "
    frm.StockValidationRequired = False
    frm.HiddenCols = "[PPSS-HiddenCols]"
    frm.Main = "SMAIN"
    frm.Support = "SALE"
    
    frm.CounterMain = "SRETURNMAIN"
    frm.CounterSupport = "SALERETURN"
    
    frm.MainAddNew = "ADDNEW_SMain"
    frm.MainSelectView = "appview_SMain_Select_View"
    frm.PartyView = "appview_AllAccounts"
    frm.PartyInitial = "C%"
    frm.ItemSelectView = "appview_SelectItemSale"
    frm.Discs = True
    frm.AutoCalculate = True
    frm.MemoFormat = "\S\A0"
    
End Sub

Public Sub ShowForm()
    frm.Show
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub

