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
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String      'Имена листов где возможна нумерация
    Dim colCxem As Collection
    Dim Cxema As classCxema
    Dim List As classListCxemy
    Dim NazvanieShkafaOld As String
    Dim NextWire As Integer
    Dim NextCableSH As Integer
    Dim NextTerm As Integer
    Dim NextElement As Integer
    Dim bWireSelect As Boolean, bCableSHSelect As Boolean, bTermSelect As Boolean, bElementSelect As Boolean
    Dim i As Integer
    
    PageName = cListNameCxema  'Имена листов где возможна нумерация
    
    'Заполняем фильтры на основе выделенных шейпов
    If obVydNaListeCx Then 'Выделенные на листе
        Set ThePage = ActivePage.PageSheet
        If ThePage.CellExists("Prop.SA_NazvanieShkafa", 0) Then NazvanieShkafa = ThePage.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
        Set colTermSelectNames = New Collection
        Set colElementSelectNames = New Collection
        If ActiveWindow.Selection.Count > 0 Then
            'Заполняем коллекцию уникальными типами элементов
            For Each vsoShape In ActiveWindow.Selection
                If ShapeSAType(vsoShape) > 1 Then   'Берем только шейпы САПР АСУ
                    UserType = ShapeSAType(vsoShape)
                    If vsoShape.CellExists("Prop.AutoNum", 0) Then
                        If vsoShape.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                            Select Case UserType
                                Case typeWire 'Провода
                                    bWireSelect = True
                                Case typeCableSH 'Кабели на схеме электрической
                                    bCableSHSelect = True
                                Case typeTerm 'Клеммы
                                    bTermSelect = True
                                    On Error Resume Next
                                    colTermSelectNames.Add vsoShape.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0), vsoShape.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0)
                                Case typeCoil, typeParent, typeElement, typePLCParent, typeSensor, typeActuator ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                                    bElementSelect = True
                                    On Error Resume Next
                                    colElementSelectNames.Add vsoShape.Cells("User.SAType").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0), vsoShape.Cells("User.SAType").Result(0) & ";" & vsoShape.Cells("Prop.SymName").ResultStr(0)
                            End Select
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    'Заполнение коллекции схем со всеми листами шейпами и фильтрами
    Set colCxem = New Collection
    Set List = New classListCxemy
    NLista = 0
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            NazvanieShkafa = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
            If NazvanieShkafa <> NazvanieShkafaOld Then
                Set Cxema = New classCxema
                Set Cxema.colListov = New Collection
                Cxema.NameCxema = NazvanieShkafa
                NazvanieShkafaOld = NazvanieShkafa
                If cbKlemCx Then Set Cxema.colTermNames = New Collection
                If cbElCx Or cbDatCx Then Set Cxema.colElementNames = New Collection
            End If
            On Error Resume Next
            colCxem.Add Cxema, Cxema.NameCxema

            'Собираем шейпы и коллекции фильтов
            If cbProvCx Or bWireSelect Then Set List.colWires = New Collection
            If cbKabCx Or bCableSHSelect Then Set List.colCableSHs = New Collection
            If cbKlemCx Or bTermSelect Then Set List.colTerms = New Collection
            If cbElCx Or cbDatCx Or bElementSelect Then Set List.colElements = New Collection
            For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
                If ShapeSAType(vsoShapeOnPage) > 1 Then   'Берем только шейпы САПР АСУ
                    UserType = ShapeSAType(vsoShapeOnPage)
                    If vsoShapeOnPage.CellExists("Prop.AutoNum", 0) Then
                        If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                            Select Case UserType
                                Case typeWire 'Провода
                                    If cbProvCx Or (obVydNaListeCx And bWireSelect) Then
                                        List.colWires.Add vsoShapeOnPage
                                    End If
                                Case typeCableSH 'Кабели на схеме электрической
                                    If cbKabCx Or (obVydNaListeCx And bCableSHSelect) Then
                                        List.colCableSHs.Add vsoShapeOnPage
                                    End If
                                Case typeTerm 'Клеммы
                                    If cbKlemCx Or (obVydNaListeCx And bTermSelect) Then
                                        List.colTerms.Add vsoShapeOnPage
                                        If Not (obVydNaListeCx And bTermSelect) Then
                                            On Error Resume Next
                                            colCxem(Cxema.NameCxema).colTermNames.Add vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                                        End If
                                    End If
                                Case typeCoil, typeParent, typeElement, typePLCParent, typeSensor, typeActuator ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                                    If cbElCx Or cbDatCx Or (obVydNaListeCx And bElementSelect) Then
                                        List.colElements.Add vsoShapeOnPage
                                        If Not (obVydNaListeCx And bElementSelect) Then
                                            On Error Resume Next
                                            colCxem(Cxema.NameCxema).colElementNames.Add vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                End If
            Next
            
            'Для перенумерации на основе выделенных присваиваем коллекции фильтров выделенных
            If obVydNaListeCx Then
                If bTermSelect Then
                    Set colCxem(Cxema.NameCxema).colTermNames = colTermSelectNames
                End If
                If bElementSelect Then
                    Set colCxem(Cxema.NameCxema).colElementNames = colElementSelectNames
                End If
            End If

            colCxem(Cxema.NameCxema).colListov.Add List, CStr(colCxem(Cxema.NameCxema).colListov.Count + 1)
            Set List = New classListCxemy
        End If
    Next

    'Перенумеровываем коллекции
    For i = 1 To colCxem.Count
        If obVseCx And Not obVydNaListeCx Then
            NazvanieShkafa = cmbxNazvanieShkafa.List(i - 1)
            GoSub RenWireKab
            GoSub RenTerm
            GoSub RenElement
        Else
            If obVydNaListeCx Then
                NazvanieShkafa = ThePage.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
            Else
                NazvanieShkafa = cmbxNazvanieShkafa.text
            End If
            GoSub RenWireKab
            GoSub RenTerm
            GoSub RenElement
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
Return

RenElement:
    If cbElCx Or cbDatCx Or bElementSelect Then
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
