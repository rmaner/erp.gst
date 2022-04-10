Attribute VB_Name = "CompanyInfo"
'FOR UNICODE HINDI USE HINDI REMINGTON (GAIL) KEYBOARD

Public Const CreditDelay = 0    'Months
Public Const ChallanDelay = 0  'Days
Public Const BankDelay = 0      'Days


Public CompanyName As String
Public CompanyDivision As String
Public AboutCompany As String
Public CompanyAddr0 As String
Public CompanyAddr1 As String

Public CompanyCity As String
Public CompanyState As String

Public CompanyPhone As String
Public CompanyFax As String
Public CompanyEmail As String
Public CompanyBillInitial As String
Public CompanyPAN As String
Public CompanyGSTIN As String

Public Sub SetCompanyInfo()
    CompanyName = mdiOne.sckGo.GReadINI("[CompanyName]")
    AboutCompany = mdiOne.sckGo.GReadINI("[AboutCompany]")
    CompanyAddr0 = mdiOne.sckGo.GReadINI("[CompanyAddr0]")
    CompanyAddr1 = mdiOne.sckGo.GReadINI("[CompanyAddr1]")
    
    CompanyCity = mdiOne.sckGo.GReadINI("[CompanyCity]")
    CompanyState = mdiOne.sckGo.GReadINI("[CompanyState]")
    
    CompanyPhone = mdiOne.sckGo.GReadINI("[CompanyPhone]")
    CompanyFax = mdiOne.sckGo.GReadINI("[CompanyFax]")
    CompanyEmail = mdiOne.sckGo.GReadINI("[CompanyEmail]")
    CompanyBillInitial = mdiOne.sckGo.GReadINI("[CompanyBillInitial]")
    CompanyPAN = mdiOne.sckGo.GReadINI("[CompanyPAN]")
    CompanyGSTIN = mdiOne.sckGo.GReadINI("[CompanyGSTIN]")
End Sub
