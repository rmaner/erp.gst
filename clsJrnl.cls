VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJrnl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public frm As New frmPRT

Private Sub Class_Initialize()
    frm.MyCAPTION = "JOURNAL VOUCHERS..."
    frm.FRAMECOLOR = vbGreen
    frm.MemoFormat = "\J\L0"
    frm.ADDPROC = "ADDNEW_JRNL "
    frm.UPDATEPROC = "UPDATE_JRNL "
    frm.Support = "JRNL"
End Sub

Public Sub ShowForm()
    frm.Show
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub



