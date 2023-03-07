Option Explicit

Private Sub UserForm_Initialize()
    cmbxNazvanieShkafa.style = fmStyleDropDownList
    Fill_cmbxNazvanieShkafa
End Sub

Private Sub btnAddElements_Click()
    AddElementyCxemyOnVID cmbxNazvanieShkafa.text
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Sub Fill_cmbxNazvanieShkafa()
    Dim colNameCxema As Collection
    Dim i As Integer
    
    Set colNameCxema = GetColNazvanieShkafa

    cmbxNazvanieShkafa.Clear
    For i = 1 To colNameCxema.Count
        cmbxNazvanieShkafa.AddItem colNameCxema.Item(i)
    Next
    cmbxNazvanieShkafa.text = ""
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub