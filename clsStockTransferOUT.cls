VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStockTransferOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public frm As New frmPPSS

Private Sub Class_Initialize()
    frm.MyCAPTION = "StockTransferOUT "
    frm.lblOrder.Caption = "TIN #:"
    frm.StockValidationRequired = False
    frm.HiddenCols = "[PPSS-HiddenCols]"
    frm.Main = "TOUTMAIN"
    frm.Support = "TOUT"
    
    frm.CounterMain = "PMAIN"
    frm.CounterSupport = "PURCHASE"
    
    frm.MainAddNew = "ADDNEW_TOUTMain"
    frm.MainSelectView = "appview_TOUTMain_Select_View"
    frm.PartyView = "appview_AllAccounts"
    frm.PartyInitial = "W%"
    frm.ItemSelectView = "appview_SelectItemTOUT"
    frm.Discs = True
    frm.AutoCalculate = True
    frm.MemoFormat = "\T\O0"
End Sub

Public Sub ShowForm()
    frm.Show
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub


