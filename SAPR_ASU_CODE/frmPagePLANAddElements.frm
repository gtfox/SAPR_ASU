Option Explicit

Private Sub UserForm_Initialize()
    cmbxNazvanieFSA.style = fmStyleDropDownList
    Fill_cmbxNazvanieFSA
End Sub

Private Sub btnAddElements_Click()
    AddSensorsFSAOnPlan cmbxNazvanieFSA.Text
    Unload Me
End Sub

Sub Fill_cmbxNazvanieFSA()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameFSA
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieFSA.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieFSA.AddItem mstrPropPageSheet(i)
    Next
    If UBound(mstrPropPageSheet) <> -1 Then
        cmbxNazvanieFSA.ListIndex = 0
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub