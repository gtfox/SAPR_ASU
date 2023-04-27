

Option Explicit

Private Sub btnAddMaster_Click()
    AddCxemaToStencil cmbxNameStencil, tbName
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    cmbxNameStencil.style = fmStyleDropDownList
    Fill_cmbxNameStencil
End Sub


Sub Fill_cmbxNameStencil()
    Dim colNameStencil As Collection
    Dim vsoDocument As Visio.Document
    Dim i As Integer
    
    Set colNameStencil = New Collection
        
    For Each vsoDocument In Application.Documents
        If vsoDocument Like "*.vss" Then
            If Not vsoDocument Like "SAPR_ASU_*" Then
                colNameStencil.Add vsoDocument
            End If
        End If
    Next
    
    If colNameStencil.Count = 0 Then
        Set vsoDocument = Application.Documents.AddEx("vss", visMSMetric, visAddDocked + visAddStencil, 1033)
        vsoDocument.SaveAs ActiveDocument.path & vsoDocument & ".vss"
        cmbxNameStencil.AddItem vsoDocument
        cmbxNameStencil.ListIndex = 0
    Else
        cmbxNameStencil.Clear
        For i = 1 To colNameStencil.Count
            cmbxNameStencil.AddItem colNameStencil.Item(i)
        Next
        If colNameStencil.Count > 0 Then
            cmbxNameStencil.ListIndex = 0
        End If
    End If
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub