VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public recs As ADODB.Recordset

Private Sub Class_Initialize()
    'set & open recordset
    Set recs = New ADODB.Recordset
    recs.CursorLocation = adUseClient
    recs.CursorType = adOpenDynamic
    recs.LockType = adLockOptimistic
End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    Set Data = recs
End Sub

Public Sub dbOpen(ByVal sql As String, Optional DontLoadGlobalArray As Integer)
    On Error Resume Next
    If recs.State = 1 Then recs.Close
    recs.Open sql, conn, adOpenDynamic, adLockOptimistic
    If DontLoadGlobalArray = 0 Then ToGlobalArray
End Sub

Public Sub Requery()
    recs.Requery
End Sub

Private Sub ToGlobalArray()
    Dim k As Integer
    Erase sArray
    If recs.State = 1 Then
        If recs.EOF <> True Then
            For k = 0 To recs.FIELDS.Count - 1
                If Not IsNull(recs.FIELDS(k)) Then sArray(k) = recs.FIELDS(k)
            Next
        Else
            For k = 0 To recs.FIELDS.Count - 1
                sArray(k) = "X"
            Next
        End If
    End If
End Sub

Private Sub Class_Terminate()
    If recs.State = 1 Then recs.Close
End Sub

