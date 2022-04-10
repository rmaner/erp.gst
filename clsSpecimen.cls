VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpecimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public frm As New frmPPSS

Private Sub Class_Initialize()
    frm.MyCAPTION = "Specimen "
    frm.StockValidationRequired = False
    frm.HiddenCols = "[PPSS-HiddenCols]"
    frm.Main = "SPECIMAIN"
    frm.Support = "SPECIMEN"
    
    frm.CounterMain = "PMAIN"
    frm.CounterSupport = "PURCHASE"
    
    frm.MainAddNew = "ADDNEW_SPECIMAIN"
    frm.MainSelectView = "appview_SPECIMAIN_Select_View"
    frm.PartyView = "appview_AllAccounts"
    frm.PartyInitial = "[D]%"
    frm.ItemSelectView = "appview_SelectItemPurchaseReturn"
    frm.Discs = True
    frm.AutoCalculate = True
    frm.MemoFormat = "\C\I0"
End Sub

Public Sub ShowForm()
    frm.Show
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub


