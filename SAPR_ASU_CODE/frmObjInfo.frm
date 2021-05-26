
Dim vsoObject As Object

Sub run(vsoShape As Object)
    Set vsoObject = vsoShape
    lblName.Caption = vsoShape.Name
    lblNameU.Caption = vsoShape.NameU
    lblID.Caption = vsoShape.ID
    lblIndex.Caption = vsoShape.Index
    On Error Resume Next
    lblNameID.Caption = vsoShape.NameID
    frmObjInfo.Show
End Sub


Private Sub CommandButton1_Click()
    vsoObject.NameU = vsoObject.Name
    lblNameU.Caption = vsoObject.NameU
End Sub

Private Sub CommandButton2_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub