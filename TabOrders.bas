Attribute VB_Name = "TabOrders"
Public Sub SetTabIndex()
    Dim a, b As Control
    For Each a In frmSale.Controls
        For Each b In frmSaleReturn.Controls
            If a.Name = b.Name Then b.TabIndex = a.TabIndex
        Next
    Next
End Sub
