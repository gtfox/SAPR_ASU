
Dim vsoShape As Shape
Option Explicit

Private Sub UserForm_Initialize()
    cmbxFormat.AddItem "А0"
    cmbxFormat.AddItem "А1"
    cmbxFormat.AddItem "А2"
    cmbxFormat.AddItem "А3"
    cmbxFormat.AddItem "А4"
    cmbxFormat.ListIndex = 3
End Sub

Sub run(vsoSh As Shape)
    Set vsoShape = vsoSh
    frmMenuVpisatVList.Show
End Sub

Private Sub btnOk_Click()
    VpisatVListExec vsoShape, cmbxFormat.ListIndex
    Unload Me
End Sub
