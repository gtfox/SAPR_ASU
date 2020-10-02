Sub Run(vsoShape As Object)
    lblName.Caption = vsoShape.Name
    lblNameU.Caption = vsoShape.NameU
    lblID.Caption = vsoShape.ID
    lblIndex.Caption = vsoShape.Index
    On Error Resume Next
    lblNameID.Caption = vsoShape.NameID
    frmFormatSpecialNameU.Show
End Sub

