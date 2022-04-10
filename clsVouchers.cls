VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public frm As New frmVouchers

Private Sub Class_Initialize()
    frm.MyCAPTION = "VOUCHERS..."
    frm.FRAMECOLOR = vbGreen
    frm.MemoFormat = "\V\R0"
    frm.ADDPROC = "ADDNEW_JRNL "
    frm.UPDATEPROC = "UPDATE_JRNL "
    frm.Support = "VOUCHERS"
End Sub

Public Sub ShowForm()
    frm.Show
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub




