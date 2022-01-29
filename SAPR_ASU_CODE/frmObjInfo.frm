
Dim vsoObject As Object

Sub run(vsoShape As Object)
    Set vsoObject = vsoShape
    frmObjInfo.Height = 90
'    tbCopyRight.Height = 0
    frameCopyRight.Visible = False
    lblName.Caption = vsoShape.Name
    lblNameU.Caption = vsoShape.NameU
    lblID.Caption = vsoShape.ID
    lblIndex.Caption = vsoShape.Index
    If vsoShape.Cells("Copyright").FormulaU <> """""" Then
        frameCopyRight.Visible = True
        frmObjInfo.Height = 160
'        tbCopyRight.Height = 55
        tbCopyRight.Value = vsoShape.Cells("Copyright").ResultStr(0)
    End If
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