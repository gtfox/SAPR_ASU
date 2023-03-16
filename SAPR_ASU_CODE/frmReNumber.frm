Dim NazvanieFSA As String
Dim NazvanieShkafa As String

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

Private Sub UserForm_Initialize()
    
    Fill_cmbxNazvanieShkafa
    Fill_cmbxNazvanieFSA
    
    cmbxNazvanieShkafa.style = fmStyleDropDownList
    cmbxNazvanieFSA.style = fmStyleDropDownList
    
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieShkafa", 0) Then
        NazvanieShkafa = ActivePage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
        cmbxNazvanieShkafa.text = NazvanieShkafa
    End If
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
        NazvanieFSA = ActivePage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
        cmbxNazvanieFSA.text = NazvanieFSA
    End If

    With mpRazdel
        .Left = Me.Left
        .Top = Me.Top
        .Width = Me.Width
        .Height = Me.Height
        .Value = IIf(NazvanieFSA = "", 0, 1)
    End With

    If NazvanieShkafa <> "" Then
        obVybCx.Value = True
    End If
    If NazvanieFSA <> "" Then
        obVybFSA.Value = True
    End If
    
    If ActiveWindow.Selection.Count > 0 Then
        obVydNaListeCx.Value = True
        obVydNaListeFSA.Value = True
    Else
        obVseTipObCx.Value = True 'Все obVybTipObCx.Value = True 'Выбранные
        obVseTipObFSA.Value = True 'Все obVybTipObFSA.Value = True 'Выбранные
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
    Dim colItems As Collection
    Dim colTermSelectNames As Collection
    Dim colElementSelectNames As Collection
    Dim ItemCol As Variant
    Dim mstrNames() As String
    Dim NumberKlemmnik As Integer
    Dim SymNameKlemmnik As String
    Dim SAType As Integer
    Dim SymName As String       'Буквенная часть нумерации
    Dim NazvanieShkafa As String   'Нумерация элементов идет в пределах одного шкафа
    Dim NazvanieShkafaOld As String
    Dim NazvanieLista As String
    Dim NazvanieListaOLD As String
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String      'Имена листов где возможна нумерация
    Dim colCxem As Collection
    Dim colCxemNames As Collection
    Dim colCxemNamesSelection As Collection
    Dim Cxema As classCxema
    Dim List As classListCxemy
    Dim NextWire As Integer
    Dim NextCableSH As Integer
    Dim NextTerm As Integer
    Dim NextElement As Integer
    Dim bWireSelect As Boolean, bCableSHSelect As Boolean, bTermSelect As Boolean, bElementSelect As Boolean
    Dim nCountCxemNames As Double
    Dim nCountColListov As Double
    Dim CxemaSelection As classCxemaSelection
    Dim colCxemaSelection As Collection
    Dim i As Integer
    Dim j As Integer
    
    PageName = cListNameCxema  'Имена листов где возможна нумерация
    
    'Заполняем фильтры на основе выделенных шейпов
    If obVydNaListeCx Then 'Выделенные на листе
        Set ThePage = ActivePage.PageSheet
'        If ThePage.CellExists("Prop.SA_NazvanieShkafa", 0) Then NazvanieShkafa = ThePage.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
        Set colTermSelectNames = New Collection
        Set colElementSelectNames = New Collection
        Set colCxemNamesSelection = New Collection
        Set colCxemaSelection = New Collection
        If ActiveWindow.Selection.Count > 0 Then
            'Заполняем коллекцию уникальными типами элементов
            For Each vsoShape In ActiveWindow.Selection
                UserType = ShapeSAType(vsoShape)
                If UserType > 1 Then   'Берем только шейпы САПР АСУ
                    If vsoShape.CellExists("Prop.AutoNum", 0) Then
                        If vsoShape.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                            If vsoShape.CellExists("User.Shkaf", 0) Then
                                NazvanieShkafa = vsoShape.Cells("User.Shkaf").ResultStr(0)
                                nCountCxemNames = colCxemNamesSelection.Count
                                On Error Resume Next
                                colCxemNamesSelection.Add NazvanieShkafa, NazvanieShkafa
                                err.Clear
                                On Error GoTo 0
                                If colCxemNamesSelection.Count > nCountCxemNames Then
                                    Set CxemaSelection = New classCxemaSelection
                                    Set CxemaSelection.colTermSelectNames = New Collection
                                    Set CxemaSelection.colElementSelectNames = New Collection
                                    CxemaSelection.NameCxemaSelection = NazvanieShkafa
                                    colCxemaSelection.Add CxemaSelection, NazvanieShkafa
                                End If
                                
                                Select Case UserType
                                    Case typeWire 'Провода
                                        colCxemaSelection(NazvanieShkafa).bWireSelect = True
                                    Case typeCableSH 'Кабели на схеме электрической
                                        colCxemaSelection(NazvanieShkafa).bCableSHSelect = True
                                    Case typeTerm 'Клеммы
                                        colCxemaSelection(NazvanieShkafa).bTermSelect = True
                                        On Error Resume Next
                                        colCxemaSelection(NazvanieShkafa).colTermSelectNames.Add vsoShape.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0), vsoShape.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0)
                                        err.Clear
                                        On Error GoTo 0
                                    Case typeCoil, typeParent, typeElement, typePLCParent, typeSensor, typeActuator ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                                        colCxemaSelection(NazvanieShkafa).bElementSelect = True
                                        On Error Resume Next
                                        colCxemaSelection(NazvanieShkafa).colElementSelectNames.Add vsoShape.Cells("User.SAType").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0), vsoShape.Cells("User.SAType").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0)
                                        err.Clear
                                        On Error GoTo 0
                                End Select
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    'Заполнение коллекции схем со всеми листами шейпами и фильтрами
    Set colCxem = New Collection
    Set colCxemNames = New Collection

    NLista = 0
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            NazvanieLista = vsoPage.name
            For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе+
                UserType = ShapeSAType(vsoShapeOnPage)
                If UserType > 1 Then   'Берем только шейпы САПР АСУ
                    If vsoShapeOnPage.CellExists("Prop.AutoNum", 0) Then
                        If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                            If vsoShapeOnPage.CellExists("User.Shkaf", 0) Then
                                NazvanieShkafa = vsoShapeOnPage.Cells("User.Shkaf").ResultStr(0)
                                
                                'Провека на присутствие схемы в выделении
                                'Если есть, заполняем переменные из неё
                                If colCxemaSelection Is Nothing Then
                                    bWireSelect = False
                                    bCableSHSelect = False
                                    bTermSelect = False
                                    Set colTermSelectNames = Nothing
                                    bElementSelect = False
                                    Set colElementSelectNames = Nothing
                                Else
                                    nCountCxemNames = colCxemaSelection.Count
                                    On Error Resume Next
                                    colCxemaSelection.Add NazvanieShkafa, NazvanieShkafa
                                    err.Clear
                                    On Error GoTo 0
                                    If colCxemaSelection.Count > nCountCxemNames Then
                                        colCxemaSelection.Remove NazvanieShkafa
                                        bWireSelect = False
                                        bCableSHSelect = False
                                        bTermSelect = False
                                        Set colTermSelectNames = Nothing
                                        bElementSelect = False
                                        Set colElementSelectNames = Nothing
                                    Else
                                        bWireSelect = colCxemaSelection(NazvanieShkafa).bTermSelect
                                        bCableSHSelect = colCxemaSelection(NazvanieShkafa).bCableSHSelect
                                        bTermSelect = colCxemaSelection(NazvanieShkafa).bTermSelect
                                        Set colTermSelectNames = colCxemaSelection(NazvanieShkafa).colTermSelectNames
                                        bElementSelect = colCxemaSelection(NazvanieShkafa).bElementSelect
                                        Set colElementSelectNames = colCxemaSelection(NazvanieShkafa).colElementSelectNames
                                    End If
                                End If

                                nCountCxemNames = colCxemNames.Count
                                On Error Resume Next
                                colCxemNames.Add NazvanieShkafa, NazvanieShkafa
                                err.Clear
                                On Error GoTo 0
                                If colCxemNames.Count > nCountCxemNames Then 'Что-то всунулось => создаём новую схему
                                    Set Cxema = New classCxema
                                    Set Cxema.colListov = New Collection
                                    Cxema.NameCxema = NazvanieShkafa
                                    If cbKlemCx Or bTermSelect Then Set Cxema.colTermNames = New Collection
                                    If cbElCx Or cbDatCx Or bElementSelect Then Set Cxema.colElementNames = New Collection
                                    colCxem.Add Cxema, Cxema.NameCxema
                                End If
                                'Создаём лист на основе фильтов
                                Set List = New classListCxemy
                                nCountColListov = colCxem(NazvanieShkafa).colListov.Count
                                On Error Resume Next
                                colCxem(NazvanieShkafa).colListov.Add List, NazvanieLista
                                err.Clear
                                On Error GoTo 0
                                If colCxem(NazvanieShkafa).colListov.Count > nCountColListov Then 'Что-то всунулось => создаём новый лист
                                    With colCxem(NazvanieShkafa).colListov(NazvanieLista)
                                        .NameListCxema = NazvanieLista
                                        If cbProvCx Or bWireSelect Then Set .colWires = New Collection
                                        If cbKabCx Or bCableSHSelect Then Set .colCableSHs = New Collection
                                        If cbKlemCx Or bTermSelect Then Set .colTerms = New Collection
                                        If cbElCx Or cbDatCx Or bElementSelect Then Set .colElements = New Collection
                                    End With
                                End If
    
                                Select Case UserType
                                    Case typeWire 'Провода
                                        If cbProvCx Or (obVydNaListeCx And bWireSelect) Then
                                            colCxem(NazvanieShkafa).colListov(NazvanieLista).colWires.Add vsoShapeOnPage
                                        End If
                                    Case typeCableSH 'Кабели на схеме электрической
                                        If cbKabCx Or (obVydNaListeCx And bCableSHSelect) Then
                                            colCxem(NazvanieShkafa).colListov(NazvanieLista).colCableSHs.Add vsoShapeOnPage
                                        End If
                                    Case typeTerm 'Клеммы
                                        If cbKlemCx Or (obVydNaListeCx And bTermSelect) Then
                                            colCxem(NazvanieShkafa).colListov(NazvanieLista).colTerms.Add vsoShapeOnPage
                                            If Not (obVydNaListeCx And bTermSelect) Then
                                                On Error Resume Next
                                                colCxem(NazvanieShkafa).colTermNames.Add vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                                                err.Clear
                                                On Error GoTo 0
                                            End If
                                        End If
                                    Case typeCoil, typeParent, typeElement, typePLCParent, typeSensor, typeActuator ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                                        If cbElCx Or cbDatCx Or (obVydNaListeCx And bElementSelect) Then
                                            colCxem(NazvanieShkafa).colListov(NazvanieLista).colElements.Add vsoShapeOnPage
                                            If Not (obVydNaListeCx And bElementSelect) Then
                                                On Error Resume Next
                                                colCxem(NazvanieShkafa).colElementNames.Add vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                                                err.Clear
                                                On Error GoTo 0
                                            End If
                                        End If
                                End Select
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next

    'Для перенумерации на основе выделенных присваиваем коллекции фильтров выделенных
    If obVydNaListeCx Then
        For i = 1 To colCxemNamesSelection.Count
            NazvanieShkafa = colCxemNamesSelection.Item(i)
            If colCxemaSelection(NazvanieShkafa).bTermSelect Then
                Set colCxem(NazvanieShkafa).colTermNames = colCxemaSelection(NazvanieShkafa).colTermSelectNames
            End If
            If colCxemaSelection(NazvanieShkafa).bElementSelect Then
                Set colCxem(NazvanieShkafa).colElementNames = colCxemaSelection(NazvanieShkafa).colElementSelectNames
            End If
        Next
    End If

    'Перенумеровываем коллекции
    For i = 1 To colCxem.Count
        If obVseCx And Not obVydNaListeCx Then
            NazvanieShkafa = colCxem.Item(i).NameCxema
            GoSub RenWireKab
            GoSub RenTerm
            GoSub RenElement
        Else
            If obVydNaListeCx Then
                For j = 1 To colCxemNamesSelection.Count
                    NazvanieShkafa = colCxemNamesSelection.Item(j)
                    GoSub RenWireKab
                    GoSub RenTerm
                    GoSub RenElement
                Next
            Else
                NazvanieShkafa = cmbxNazvanieShkafa.text
                GoSub RenWireKab
                GoSub RenTerm
                GoSub RenElement
            End If
            Exit For
        End If
    Next

Exit Sub

RenWireKab:
    NextWire = 0
    NextCableSH = 0
    For Each List In colCxem(NazvanieShkafa).colListov
        If cbProvCx Or bWireSelect Then
            NextWire = ReNumber(List.colWires, NextWire)
        End If
        If cbKabCx Or bCableSHSelect Then
            NextCableSH = ReNumber(List.colCableSHs, NextCableSH)
        End If
    Next
Return

RenTerm:
    If cbKlemCx Or bTermSelect Then
        If Not colCxem(NazvanieShkafa).colTermNames Is Nothing Then
            If colCxem(NazvanieShkafa).colTermNames.Count > 0 Then
                For Each ItemCol In colCxem(NazvanieShkafa).colTermNames
                    mstrNames = Split(ItemCol, ";")
                    NumberKlemmnik = CInt(mstrNames(0))
                    SymNameKlemmnik = mstrNames(1)
                    NextTerm = 0
                    For Each List In colCxem(NazvanieShkafa).colListov
                        'По фильтрам заполняем коллецию для перенумерации
                        Set colItems = New Collection
                        For Each vsoShapeOnPage In List.colTerms
                            If vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) = NumberKlemmnik And vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) = SymNameKlemmnik Then
                                colItems.Add vsoShapeOnPage
                            End If
                        Next
                        NextTerm = ReNumber(colItems, NextTerm)
                    Next
                Next
            End If
        End If
    End If
Return

RenElement:
    If cbElCx Or cbDatCx Or bElementSelect Then
        If Not colCxem(NazvanieShkafa).colElementNames Is Nothing Then
            If colCxem(NazvanieShkafa).colElementNames.Count > 0 Then
                For Each ItemCol In colCxem(NazvanieShkafa).colElementNames
                    mstrNames = Split(ItemCol, ";")
                    SAType = CInt(mstrNames(0))
                    SymName = mstrNames(1)
                    NextElement = 0
                    For Each List In colCxem(NazvanieShkafa).colListov
                        'По фильтрам заполняем коллецию для перенумерации
                        Set colItems = New Collection
                        For Each vsoShapeOnPage In List.colElements
                            If vsoShapeOnPage.Cells("User.SAType").Result(0) = SAType And vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) = SymName Then
                                colItems.Add vsoShapeOnPage
                            End If
                        Next
                        NextElement = ReNumber(colItems, NextElement)
                    Next
                Next
            End If
        End If
    End If
Return

End Sub

Public Sub ReNumberFSA()
'------------------------------------------------------------------------------------------------------------
' Macros        : ReNumberFSA - Перенумерация элементов ФСА

                'Нумерация ведется с учетом имени контура
                'Перенумерация происходит слева направо, сверху вниз
                'независимо от порядка появления элементов на схеме
                'и независимо от их номеров до перенумерации.
                'Если в элементе Prop.AutoNum=0 то он не участвует в перенумерации
                'Перенумерация элементов идет в пределах одной ФСА или всех ФСА
                'Параметры перенумерации задаются в форме frmReNumber
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim ThePage As Visio.Shape
    Dim vsoShapeOnPage As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim colItems As Collection
    Dim colElementSelectNames As Collection
    Dim ItemCol As Variant
    Dim mstrNames() As String
    Dim SAType As Integer
    Dim NameKontur As String
    Dim SymName As String       'Буквенная часть нумерации
    Dim NazvanieFSA As String   'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String      'Имена листов где возможна нумерация
    Dim colFSA As Collection
    Dim FSA As classFSA
    Dim List As classListFSA
    Dim NazvanieFSAOld As String
    Dim NextPodval As Integer
    Dim NextElement As Integer
    Dim bPodvalSelect As Boolean, bElementSelect As Boolean
    Dim i As Integer
    
    PageName = cListNameFSA  'Имена листов где возможна нумерация
    
    'Заполняем фильтры на основе выделенных шейпов
    If obVydNaListeFSA Then 'Выделенные на листе
        Set ThePage = ActivePage.PageSheet
        If ThePage.CellExists("Prop.SA_NazvanieFSA", 0) Then NazvanieFSA = ThePage.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
        Set colElementSelectNames = New Collection
        If ActiveWindow.Selection.Count > 0 Then
            'Заполняем коллекцию уникальными типами элементов
            For Each vsoShape In ActiveWindow.Selection
                If ShapeSAType(vsoShape) > 1 Then   'Берем только шейпы САПР АСУ
                    UserType = ShapeSAType(vsoShape)
                    If vsoShape.CellExists("Prop.AutoNum", 0) Then
                        If vsoShape.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                            Select Case UserType
                                Case typeFSAPodval 'Подвал на ФСА
                                    bPodvalSelect = True
                                Case typeFSASensor 'Датчик на ФСА
                                    bElementSelect = True
                                    On Error Resume Next
                                    colElementSelectNames.Add vsoShape.Cells("User.SAType").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0) & ";" & vsoShape.Cells("Prop.NameKontur").ResultStr(0), vsoShape.Cells("User.SAType").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0) & ";" & vsoShape.Cells("Prop.NameKontur").ResultStr(0)
                            End Select
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    'Заполнение коллекции схем со всеми листами шейпами и фильтрами
    Set colFSA = New Collection
    Set List = New classListFSA
    NLista = 0
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            NazvanieFSA = vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
            If NazvanieFSA <> NazvanieFSAOld Then
                Set FSA = New classFSA
                Set FSA.colListov = New Collection
                FSA.NameFSA = NazvanieFSA
                NazvanieFSAOld = NazvanieFSA
                If cbDatFSA Then Set FSA.colElementNames = New Collection
            End If
            On Error Resume Next
            colFSA.Add FSA, FSA.NameFSA

            'Собираем шейпы и коллекции фильтов
            If cbPodFSA Or bPodvalSelect Then Set List.colPodvals = New Collection
            If cbDatFSA Or bElementSelect Then Set List.colElements = New Collection
            For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
                If ShapeSAType(vsoShapeOnPage) > 1 Then   'Берем только шейпы САПР АСУ
                    UserType = ShapeSAType(vsoShapeOnPage)
                    If vsoShapeOnPage.CellExists("Prop.AutoNum", 0) Then
                        If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                            Select Case UserType
                                Case typeFSAPodval 'Подвал на ФСА
                                    If cbPodFSA Or (obVydNaListeFSA And bPodvalSelect) Then
                                        List.colPodvals.Add vsoShapeOnPage
                                    End If
                                Case typeFSASensor 'Датчик на ФСА
                                    If cbDatFSA Or (obVydNaListeFSA And bElementSelect) Then
                                        List.colElements.Add vsoShapeOnPage
                                        If Not (obVydNaListeFSA And bElementSelect) Then
                                            On Error Resume Next
                                            colFSA(FSA.NameFSA).colElementNames.Add vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) & ";" & vsoShapeOnPage.Cells("Prop.NameKontur").ResultStr(0), vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) & ";" & vsoShapeOnPage.Cells("Prop.NameKontur").ResultStr(0)
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                End If
            Next
            
            'Для перенумерации на основе выделенных присваиваем коллекции фильтров выделенных
            If obVydNaListeFSA Then
                If bElementSelect Then
                    Set colFSA(FSA.NameFSA).colElementNames = colElementSelectNames
                End If
            End If

            colFSA(FSA.NameFSA).colListov.Add List, CStr(colFSA(FSA.NameFSA).colListov.Count + 1)
            Set List = New classListFSA
        End If
    Next

    'Перенумеровываем коллекции
    For i = 1 To colFSA.Count
        If obVseFSA And Not obVydNaListeFSA Then
            NazvanieFSA = cmbxNazvanieFSA.List(i - 1)
            GoSub RenPodval
            GoSub RenElement
        Else
            If obVydNaListeFSA Then
                NazvanieFSA = ThePage.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
            Else
                NazvanieFSA = cmbxNazvanieFSA.text
            End If
            GoSub RenPodval
            GoSub RenElement
            Exit For
        End If
    Next

Exit Sub

RenPodval:
    NextPodval = 0
    For Each List In colFSA(NazvanieFSA).colListov
        If cbPodFSA Or bPodvalSelect Then
            NextPodval = ReNumber(List.colPodvals, NextPodval)
        End If
    Next
Return

RenElement:
    If cbDatFSA Or bElementSelect Then
        If colFSA(NazvanieFSA).colElementNames.Count > 0 Then
            For Each ItemCol In colFSA(NazvanieFSA).colElementNames
                mstrNames = Split(ItemCol, ";")
                SAType = CInt(mstrNames(0))
                SymName = mstrNames(1)
                NameKontur = mstrNames(2)
                NextElement = 0
                For Each List In colFSA(NazvanieFSA).colListov
                    'По фильтрам заполняем коллецию для перенумерации
                    Set colItems = New Collection
                    For Each vsoShapeOnPage In List.colElements
                        If vsoShapeOnPage.Cells("User.SAType").Result(0) = SAType And vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) = SymName And vsoShapeOnPage.Cells("Prop.NameKontur").ResultStr(0) = NameKontur Then
                            colItems.Add vsoShapeOnPage
                        End If
                    Next
                    NextElement = ReNumber(colItems, NextElement)
                Next
            Next
        End If
    End If
Return

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

Sub Fill_cmbxNazvanieFSA()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameFSA
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieFSA.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieFSA.AddItem mstrPropPageSheet(i)
    Next
    cmbxNazvanieFSA.text = ""
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
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnCloseFSA_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub
