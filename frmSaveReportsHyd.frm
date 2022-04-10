VERSION 5.00
Begin VB.Form frmSaveReportsHyd 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackupDaysReports 
      Caption         =   "Command1"
      Height          =   1980
      Left            =   2490
      TabIndex        =   0
      Top             =   2265
      Width           =   4080
   End
End
Attribute VB_Name = "frmSaveReportsHyd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN(8) As clsData

Private Sub cmdBackupDaysReports_Click()
    mdiOne.CDlg.ShowSave
    
    SQ = "SELECT * FROM " & MainSelectView & " Where DateDiff(D, DBDate," & QT(Format(Now, "DD-MMM-YY")) & ")=0 ORDER BY 1 Desc"
    CN(0).dbOpen SQ, 1
    While Not CN(0).recs.EOF
        Sale.frm.txtDBRef = CN(0).recs!DBRef
        Sale.frm.txtDBRef_LostFocus
        
        mdiOne.CDlg.FileName = Format(Sale.frm.txtDBRef.Text, Sale.frm.MemoFormat) & "-" & txtID.Text & "-" & Format(Now, "DD-MMM-YY HHMM")
        Sale.frm.flxOrder.SaveGrid mdiOne.CDlg.FileName, flexFileExcel

        CN(0).recs.MoveNext
    Wend
End Sub
