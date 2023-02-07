

Option Explicit

Private Sub UserForm_Initialize()
    cmbxNazvanieShkafa.style = fmStyleDropDownList
    Fill_cmbxNazvanieShkafa
End Sub

Private Sub btnAddElements_Click()
    AddPagesSVP cmbxNazvanieShkafa.text
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Sub Fill_cmbxNazvanieShkafa()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieShkafa.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieShkafa.AddItem mstrPropPageSheet(i)
    Next
    If UBound(mstrPropPageSheet) <> -1 Then
        cmbxNazvanieShkafa.ListIndex = 0
    End If
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub