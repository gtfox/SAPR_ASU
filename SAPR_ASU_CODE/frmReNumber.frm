Dim NazvanieFSA As String
Dim NazvanieShemy As String



Private Sub brnRenumberCx_Click()
    ReNumberShemy
    Unload Me
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
    
    If ActiveWindow.Selection.Count > 0 Then
        obVydNaListeCx.Value = True
        obVydNaListeFSA.Value = True
    Else
        obVybTipObCx.Value = True
        obVybTipObFSA.Value = True
    End If

End Sub

Public Sub ReNumberShemy()
'------------------------------------------------------------------------------------------------------------
' Macros        : ReNumberShemy - Перенумерация элементов схемы

                'Перенумерация происходит слева направо, сверху вниз
                'независимо от порядка появления элементов на схеме
                'и независимо от их номеров до перенумерации.
                'Если в элементе Prop.AutoNum=0 то он не участвует в перенумерации
                'Перенумерация элементов идет в пределах одной схемы или всех схем
                'Параметры перенумерации задаются в форме frmReNumber
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim ThePage As Visio.Shape
    Dim vsoShapeOnPage As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim colShp 'As Dictionary
    Dim colWires As Collection
    Dim colCableSHs As Collection
    Dim colTerms As Collection
    Dim colTermNames As Collection
    Dim colElements As Collection
    Dim colElementNames As Collection
    Dim ItemCol As Variant
    Dim mstrNames() As String
    Dim NumberKlemmnik As Integer
    Dim SymNameKlemmnik As String
    Dim SAType As Integer
    Dim ColKey As String
    Dim SymName As String       'Буквенная часть нумерации
    Dim NazvanieShemy As String   'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String      'Имена листов где возможна нумерация
    Dim NListaSxemy As Integer
    Dim i As Integer
    
    
    Set colWires = New Collection
    Set colCableSHs = New Collection
    Set colTerms = New Collection
    Set colTermNames = New Collection
    Set colElements = New Collection
    Set colElementNames = New Collection
    
    Set ThePage = ActivePage.PageSheet
    If ThePage.CellExists("Prop.SA_NazvanieShemy", 0) Then NazvanieShemy = ThePage.Cells("Prop.SA_NazvanieShemy").ResultStr(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер.
    NazvanieShemy = cmbxNazvanieShemy.Text
    PageName = cListNameCxema  'Имена листов где возможна нумерация

    If obVydNaListeCx Then 'Выделенные на листе
        If ActiveWindow.Selection.Count > 0 Then
            Set colShp = CreateObject("Scripting.Dictionary")
            'Заполняем коллекцию уникальными типами элементов
            For Each vsoShape In ActiveWindow.Selection
                ColKey = ShapeSAType(vsoShape)
                If vsoShape.CellExists("Prop.SymName", 0) Then ColKey = ColKey & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0)
                If vsoShape.CellExists("Prop.NumberKlemmnik", 0) Then ColKey = ColKey & ";" & vsoShape.Cells("Prop.NumberKlemmnik").Result(0)
                If ColKey <> "" Then
                    On Error Resume Next
                    colShp.Add vsoShape, ColKey
                End If
            Next
            'Перенумеровываем коллекцию
            For Each vsoShape In colShp
                ReNuberCxByShape vsoShape, obVseCx
            Next
        End If
    Else 'Выбранные на форме
        'Перебор всех схем
        For i = 0 To cmbxNazvanieShemy.ListCount - 1
            NListaSxemy = 0
            For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
                If Left(vsoPage.Name, Len(PageName)) = PageName Then    'Берем те, что содержат "Схема" в имени
                    If vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then    'Берем все схемы с именем
                        NListaSxemy = NListaSxemy + 1
                        For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                            If ShapeSAType(vsoShapeOnPage) > 1 Then   'Берем только шейпы САПР АСУ
                                UserType = ShapeSAType(vsoShapeOnPage)
                                If vsoShapeOnPage.CellExists("Prop.AutoNum", 0) Then
                                    If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                                        Select Case UserType
                                            Case typeWire 'Провода
                                                If cbProvCx Then
                                                    colWires.Add vsoShapeOnPage, NListaSxemy & ";" & vsoShapeOnPage.NameU
                                                End If
                                            Case typeCableSH 'Кабели на схеме электрической
                                                If cbKabCx Then
                                                    colCableSHs.Add vsoShapeOnPage, NListaSxemy & ";" & vsoShapeOnPage.NameU
                                                End If
                                            Case typeTerm 'Клеммы
                                                If cbKlemCx Then
                                                    colTerms.Add vsoShapeOnPage, NListaSxemy & ";" & vsoShapeOnPage.NameU
                                                    On Error Resume Next
                                                    colTermNames.Add vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                                                End If
                                            Case typeCoil, typeParent, typeElement, typePLCParent, typeSensor, typeActuator ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                                                If cbElCx Or cbDatCx Then
                                                    colElements.Add vsoShapeOnPage, NListaSxemy & ";" & vsoShapeOnPage.NameU
                                                    On Error Resume Next
                                                    colElementNames.Add vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                                                End If
                                        End Select
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
            
            'Перенумеровываем коллекции
            If cbProvCx Then
                ReNumber colWires
            End If
            If cbKabCx Then
                ReNumber colCableSHs
            End If
            If cbKlemCx Then
                If colTermNames.Count > 0 Then
                    For Each ItemCol In colTermNames
                        mstrNames = Split(ItemCol, ";")
                        NumberKlemmnik = CInt(mstrNames(0))
                        SymNameKlemmnik = mstrNames(1)
                        Set colShp = CreateObject("Scripting.Dictionary")
                        'По уникальным типам заполняем коллецию для перенумерации
                        For Each vsoShape In colTerms
                            If vsoShape.Cells("Prop.NumberKlemmnik").Result(0) = NumberKlemmnik And vsoShape.Cells("Prop.SymName").ResultStr(0) = SymNameKlemmnik Then
                                colShp.Add vsoShape
                            End If
                        Next
                        ReNumber colShp
                    Next
                End If
            End If
            If cbElCx Or cbDatCx Then
                If colElementNames.Count > 0 Then
                    For Each ItemCol In colElementNames
                        mstrNames = Split(ItemCol, ";")
                        SAType = CInt(mstrNames(0))
                        SymName = mstrNames(1)
                        Set colShp = CreateObject("Scripting.Dictionary")
                        'По уникальным типам заполняем коллецию для перенумерации
                        For Each vsoShape In colElements
                            If vsoShape.Cells("User.SAType").Result(0) = SAType And vsoShape.Cells("Prop.SymName").ResultStr(0) = SymName Then
                                colShp.Add vsoShape
                            End If
                        Next
                        ReNumber colShp
                    Next
                End If
            End If
            If obVseCx Then
                NazvanieShemy = cmbxNazvanieShemy.List(i)
            Else
                Exit For
            End If
        Next
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

Private Sub obVseTipObCx_Change()
    If obVseTipObCx = True Then
        cbElCx.Value = True
        cbProvCx.Value = True
        cbKlemCx.Value = True
        cbKabCx.Value = True
        cbDatCx.Value = True
    End If
End Sub

Private Sub obVseTipObFSA_Change()
    If obVseTipObFSA = True Then
        cbDatFSA.Value = True
        cbPodFSA.Value = True
    End If
End Sub

Private Sub obVydNaListeCx_Change()
    cbElCx.Value = False
    cbProvCx.Value = False
    cbKlemCx.Value = False
    cbKabCx.Value = False
    cbDatCx.Value = False
End Sub

Private Sub obVydNaListeFSA_Change()
    cbDatFSA.Value = False
    cbPodFSA.Value = False
End Sub

Private Sub cbDatFSA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    obVybTipObFSA.Value = True
End Sub

Private Sub cbPodFSA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    obVybTipObFSA.Value = True
End Sub

Private Sub cbElCx_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    obVybTipObCx.Value = True
End Sub

Private Sub cbProvCx_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    obVybTipObCx.Value = True
End Sub

Private Sub cbKlemCx_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    obVybTipObCx.Value = True
End Sub

Private Sub cbKabCx_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    obVybTipObCx.Value = True
End Sub

Private Sub cbDatCx_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    obVybTipObCx.Value = True
End Sub

Private Sub btnCloseCx_Click()
    Unload Me
End Sub

Private Sub btnCloseFSA_Click()
    Unload Me
End Sub
