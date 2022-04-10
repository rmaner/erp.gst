VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAccountSummary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Summary..."
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAccountSummary 
      Height          =   5610
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmAccountSummary.frx":0000
      Top             =   405
      Width           =   8565
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   315
      Left            =   795
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "ddd, dd-MMM-yy"
      Format          =   20774915
      CurrentDate     =   38675
   End
   Begin VB.Label Label1 
      Caption         =   "Date: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   915
   End
End
Attribute VB_Name = "frmAccountSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New clsData
Private Const MyCAPTION = "ACCOUNTS SUMMARY... "

Private Sub Form_Load()
    On Error Resume Next
    Me.Move 0, 0
    mdiOne.SetFormFont Me, "Courier New", 10
    DT1.Value = Now
    DT1_Click
End Sub

Private Sub DT1_Click()
    AccountSummaryReport DT1.Value
End Sub

Private Sub DT1_Change()
    AccountSummaryReport DT1.Value
End Sub

Private Sub AccountSummaryReport(Optional ByVal dt As Variant)
    If IsMissing(dt) Then dt = Now
    txtAccountSummary.Text = "Account Summary on " & Format(dt, "DD-MMM-yyyy") & vbCrLf & String(45, ".") & vbCrLf
    txtAccountSummary.Text = txtAccountSummary.Text & getBalanceLine("N001", dt) & vbCrLf
    txtAccountSummary.Text = txtAccountSummary.Text & getBalanceLine("N002", dt) & vbCrLf & vbCrLf
    txtAccountSummary.Text = txtAccountSummary.Text & getBalanceLine("N003", dt) & vbCrLf
    txtAccountSummary.Text = txtAccountSummary.Text & getBalanceLine("N004", dt) & vbCrLf & vbCrLf
    txtAccountSummary.Text = txtAccountSummary.Text & getBalanceLine("N024", dt) & vbCrLf
    txtAccountSummary.Text = txtAccountSummary.Text & getBalanceLine("N025", dt) & vbCrLf & vbCrLf
End Sub

Private Function getBalanceLine(ByVal id As String, ByVal dt As Date) As String
    CN.dbOpen "SELECT NAME FROM PERSONAL WHERE ID = " & QT(id), 0
    If Not CN.recs.EOF Then
        partyname = CN.recs!Name
        getBalanceLine = RPad(partyname, 25) & LPad(Format(WhatIsLedgerBalance(id, dt), "##,##0.00 Dr; ##,##0.00 Cr; NIL"), 20)
    Else
        getBalanceLine = ""
    End If
End Function

Private Function LPad(ByVal str As String, ByVal strSize As Integer) As String
    LPad = Right(String(100, " ") & str, strSize)
End Function

Private Function RPad(ByVal str As String, ByVal strSize As Integer) As String
    RPad = Left(str & String(100, " "), strSize)
End Function

