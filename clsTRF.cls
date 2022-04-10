VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public frm As New frmTRF

Private Sub Class_Initialize()
    frm.MyCAPTION = "TRANSFER/CONTRA..."
    frm.FRAMECOLOR = vbBlue
    frm.MemoFormat = "\T\R0"
    frm.ADDPROC = "ADDNEW_TRF "
    frm.UPDATEPROC = "UPDATE_TRF "
    frm.Support = "TRF"
End Sub

Public Sub ShowForm()
    frm.Show
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub



