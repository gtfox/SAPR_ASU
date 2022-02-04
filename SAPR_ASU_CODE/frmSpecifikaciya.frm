

Dim NazvanieFSA As String
Dim NazvanieShemy As String

Private Sub btnRenumberCx_Click()
    ReNumberShemy
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnRenumberFSA_Click()
    ReNumberFSA
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnExportCx_Click()
FindElementShemy
End Sub

Private Sub obTekListCx_Click()
    frameOutListCx.Visible = True
End Sub

Private Sub obTekListFSA_Click()
    frameOutListFSA.Visible = True
End Sub

Private Sub obVseCx_Click()
    frameOutListCx.Visible = False
End Sub

Private Sub obVseFSA_Click()
    frameOutListFSA.Visible = False
End Sub

Private Sub obVybCx_Click()
    frameOutListCx.Visible = False
End Sub

Private Sub obVybFSA_Click()
    frameOutListFSA.Visible = False
End Sub

Private Sub UserForm_Initialize()
    
    Fill_cmbxNazvanieShemy
    Fill_cmbxNazvanieFSA
    
    cmbxNazvanieShemy.style = fmStyleDropDownList
    cmbxNazvanieFSA.style = fmStyleDropDownList
    
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieShemy", 0) Then
        NazvanieShemy = ActivePage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0)
        cmbxNazvanieShemy.Text = NazvanieShemy
    End If
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
        NazvanieFSA = ActivePage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
        cmbxNazvanieFSA.Text = NazvanieFSA
    End If

    With mpRazdel
        .Left = Me.Left
        .Top = Me.Top
        .Width = Me.Width
        .Height = Me.Height
        .Value = IIf(NazvanieFSA = "", 0, 1)
    End With

    If NazvanieShemy <> "" Then
        obVybCx.Value = True
    End If
    If NazvanieFSA <> "" Then
        obVybFSA.Value = True
    End If
    
    obNaListFSA.Value = True
    obNaListCx.Value = True
End Sub

Public Sub FindElementShemy()
'------------------------------------------------------------------------------------------------------------
' Macros        : FindElementShemy - Поиск элементов схемы и заполнение полей спецификации

                '
                '
'------------------------------------------------------------------------------------------------------------
    Dim clsStrokaSpecif As classStrokaSpecifikacii
    Dim colStrokaSpecif As Collection
    Dim colPage As Collection
    Dim colCxem As Collection
    Dim nCount As Double
    Dim strColKey As String
    
    
    Dim vsoPage As Visio.Page
    Dim ThePage As Visio.Shape
    Dim vsoShapeOnPage As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim colItems As Collection
    Dim colTermSelectNames As Collection
    Dim colElementSelectNames As Collection
    Dim ItemCol As Variant
    Dim mstrNames() As String
    Dim NumberKlemmnik As Integer
    Dim SymNameKlemmnik As String
    Dim SAType As Integer
    Dim SymName As String       'Буквенная часть нумерации
    Dim NazvanieShemy As String   'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String      'Имена листов где возможна нумерация

    Dim Cxema As classCxema
    Dim List As classListCxemy
    Dim NazvanieShemyOld As String
    Dim NextWire As Integer
    Dim NextCableSH As Integer
    Dim NextTerm As Integer
    Dim NextElement As Integer
    Dim bWireSelect As Boolean, bCableSHSelect As Boolean, bTermSelect As Boolean, bElementSelect As Boolean
    Dim i As Integer
    
    PageName = cListNameCxema  'Имена листов
    
    Set colPage = New Collection
    Set colCxem = New Collection

    For i = 1 To cmbxNazvanieShemy.ListCount
        NazvanieShemy = cmbxNazvanieShemy.List(i - 1)
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.Name Like PageName & "*" Then
                If NazvanieShemy = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) Then
                    colPage.Add vsoPage, vsoPage.Name
                End If
            End If
        Next
        If colPage.Count > 0 Then
            colCxem.Add colPage, NazvanieShemy
        End If
        Set colPage = New Collection
    Next

    Set colPage = New Collection
    If obVseCx Then
        For Each colPage In colCxem
            For Each vsoPage In colPage
                'Собираем шейпы и коллекции фильтов
                For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
                    If ShapeSAType(vsoShapeOnPage) > 1 Then   'Берем только шейпы САПР АСУ
                        UserType = ShapeSAType(vsoShapeOnPage)
                        Select Case UserType
                            Case typeCableSH 'Кабели на схеме электрической
                                
                            Case typeTerm 'Клеммы
                                
                            Case typeCoil, typeParent, typeElement, typePLCParent, typePLCModParent, typeSensor, typeActuator ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                                clsStrokaSpecif.SymName = vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                                clsStrokaSpecif.SAType = vsoShapeOnPage.Cells("User.SAType").Result(0)
                                clsStrokaSpecif.NazvanieDB = vsoShapeOnPage.Cells("Prop.NazvanieDB").ResultStr(0)
                                clsStrokaSpecif.ArtikulDB = vsoShapeOnPage.Cells("Prop.ArtikulDB").ResultStr(0)
                                clsStrokaSpecif.ProizvoditelDB = vsoShapeOnPage.Cells("Prop.ProizvoditelDB").ResultStr(0)
                                clsStrokaSpecif.CenaDB = vsoShapeOnPage.Cells("Prop.CenaDB").ResultStr(0)
                                clsStrokaSpecif.EdDB = vsoShapeOnPage.Cells("Prop.EdDB").ResultStr(0)
                                clsStrokaSpecif.KolVo = 1
                                clsStrokaSpecif.NomeraPozicij = CStr(vsoShapeOnPage.Cells("Prop.Number").Result(0))
                                strColKey = vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) & ";" & vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.ArtikulDB").ResultStr(0)
                                nCount = colStrokaSpecif.Count
                                On Error Resume Next
                                colStrokaSpecif.Add clsStrokaSpecif, strColKey
                                If colStrokaSpecif.Count = nCount Then 'Если кол-во не увеличелось, значит уже есть такой элемент - увеличиваем кол-во
                                    colStrokaSpecif(strColKey).KolVo = colStrokaSpecif(strColKey).KolVo + 1
                                    colStrokaSpecif(strColKey).NomeraPozicij = colStrokaSpecif(strColKey).NomeraPozicij & ";" & CStr(vsoShapeOnPage.Cells("Prop.Number").Result(0))
                                End If
                        End Select
                    End If
                Next
            Next
            If colPage.Count > 0 Then
                colCxem.Add colPage, NazvanieShemy
            End If
            Set colPage = New Collection
        Next
    ElseIf obVybCx Then
        NazvanieShemy = cmbxNazvanieShemy.Text
        For Each vsoPage In colCxem(NazvanieShemy)
            If vsoPage.Name Like PageName & "*" Then
                If NazvanieShemy = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) Then
                    colPage.Add vsoPage, vsoPage.Name
                End If
            End If
        Next
    ElseIf obTekListCx Then
'        NazvanieShemy = cmbxNazvanieShemy.Text
    End If

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
    cmbxNazvanieShemy.Text = ""
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
    cmbxNazvanieFSA.Text = ""
End Sub

Private Sub btnCloseCx_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnCloseFSA_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub
