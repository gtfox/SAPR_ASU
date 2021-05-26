Option Explicit

Private Sub UserForm_Initialize()
    cmbxNazvanieShemy.style = fmStyleDropDownList
    Fill_cmbxNazvanieShemy
End Sub

Private Sub btnAddElements_Click()
    AddElementyCxemyOnVID cmbxNazvanieShemy.Text
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Sub Fill_cmbxNazvanieShemy()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieShemy.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieShemy.AddItem mstrPropPageSheet(i)
    Next
    If UBound(mstrPropPageSheet) <> -1 Then
        cmbxNazvanieShemy.ListIndex = 0
    End If
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub