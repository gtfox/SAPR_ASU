
'Function FillColWires(shpSensorIO As Visio.Shape) As Collection
''------------------------------------------------------------------------------------------------------------
'' Function        : FillColWires - Находим подключенные провода и суем их в коллекцию
''------------------------------------------------------------------------------------------------------------
'    Dim colWires As Collection
'    Dim shpSensorTerm As Visio.Shape
'
'    Set colWires = New Collection
'    For Each shpSensorTerm In shpSensorIO.Shapes
'        If ShapeSATypeIs(shpSensorTerm, typeSensorTerm) Then
'            If shpSensorTerm.FromConnects.Count = 1 Then
'                If ShapeSATypeIs(shpSensorTerm.FromConnects.FromSheet, typeWire) Then
'                    colWires.Add shpSensorTerm.FromConnects.FromSheet, shpSensorTerm.FromConnects.FromSheet.name
'                End If
'            End If
'        End If
'    Next
'    Set FillColWires = colWires
'End Function
'
'Function FillColWiresOnPage(shpSensorIO As Visio.Shape) As Collection
''------------------------------------------------------------------------------------------------------------
'' Function        : FillColWiresOnPage - Находим подключенные провода находящиеся на листе (не в группе кабеля) и суем их в коллекцию
''------------------------------------------------------------------------------------------------------------
'    Dim colWires As Collection
'    Dim shpSensorTerm As Visio.Shape
'
'    Set colWires = New Collection
'    For Each shpSensorTerm In shpSensorIO.Shapes
'        If ShapeSATypeIs(shpSensorTerm, typeSensorTerm) Then
'            If shpSensorTerm.FromConnects.Count = 1 Then
'                If ShapeSATypeIs(shpSensorTerm.FromConnects.FromSheet, typeWire) Then
'                    If Not shpSensorTerm.FromConnects.FromSheet.Parent.Type = visTypeGroup Then
'                        colWires.Add shpSensorTerm.FromConnects.FromSheet, shpSensorTerm.FromConnects.FromSheet.name
'                    End If
'                End If
'            End If
'        End If
'    Next
'    Set FillColWiresOnPage = colWires
'End Function
'
'Function FillColTerms(colWires As Collection) As Collection
''------------------------------------------------------------------------------------------------------------
'' Function        : FillColTerms - Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
''------------------------------------------------------------------------------------------------------------
'    Dim colTerms As Collection
'    Dim shpWire As Visio.Shape
'
'    Set colTerms = New Collection
'
'    For Each shpWire In colWires
'        If ShapeSATypeIs(shpWire, typeWire) Then
'            If shpWire.Connects.Count = 2 Then
'                For i = 1 To shpWire.Connects.Count
'                    If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeTerm) Then
'                        colTerms.Add shpWire.Connects(i).ToSheet
'                    End If
'                Next
'            End If
'        End If
'    Next
'    Set FillColTerms = colTerms
'End Function
'
'Function FindSensorFromKabel(shpKabel As Visio.Shape) As Visio.Shape
''------------------------------------------------------------------------------------------------------------
'' Function        : FindSensorFromKabel - Находим датчик/привод подключенный кабелем
''------------------------------------------------------------------------------------------------------------
'    Dim shpWire As Visio.Shape
'
'    For Each shpWire In shpKabel.Shapes
'        If ShapeSATypeIs(shpWire, typeWire) Then
'            If shpWire.Connects.Count = 2 Then
'                For i = 1 To shpWire.Connects.Count
'                    If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeSensorTerm) Then
'                        Set FindSensorFromKabel = shpWire.Connects(i).ToSheet.Parent.Parent
'                        Exit Function
'                    End If
'                Next
'            End If
'        End If
'    Next
'End Function
'
'Public Sub PageSVPAddKabeliFrm()
'    Load frmPageSVPAddKabeli
'    frmPageSVPAddKabeli.Show
'End Sub
'
'Public Sub AddPagesSVP(NazvanieShkafa As String)
''------------------------------------------------------------------------------------------------------------
'' Macros        : AddPagesSVP - Создает листы СВП
'                'Заполняет листы СВП датчиками, отсортированными по возрастанию их координаты Х на эл. схеме
''------------------------------------------------------------------------------------------------------------
''    Dim NazvanieShkafa As String
'    Dim ThePage As Visio.Shape
'    Dim vsoShapeOnPage As Visio.Shape
'    Dim vsoPage As Visio.Page
'    Dim PageName As String
'    Dim shpElement As Shape
'    Dim Prev As Shape
'    Dim colShpPage As Collection
'    Dim colShpDoc As Collection
'    Dim shpMas() As Shape
'    Dim shpTemp As Shape
'    Dim Index As Integer
'    Dim ShinaNumber As Boolean 'Нумерация проводов кабеля по типу ШИНЫ(Номер=Клемме), или Номер провода кабеля = Порядковому номеру жилы в кабеле
'    Dim ss As String
'    Dim i As Integer, ii As Integer, j As Integer, n As Integer
'
'    ShinaNumber = 1
'
'    PastePoint = NachaloVstavki
'
'    Set ThePage = ActivePage.PageSheet
'
'    Set colShpDoc = New Collection
'
'    PageName = cListNameCxema
'    'If ThePage.CellExists("Prop.SA_NazvanieShkafa", 0) Then NazvanieShkafa = ThePage.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер.
''    NazvanieShkafa = 4
'
'    'Цикл поиска датчиков и приводов
'    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
'        If vsoPage.name Like PageName & "*" Then    'Берем те, что содержат "Схема" в имени
'            Set colShpPage = New Collection
'            For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
'                If vsoShapeOnPage.CellExists("User.Shkaf", 0) Then
'                    If vsoShapeOnPage.Cells("User.Shkaf").ResultStr(0) = NazvanieShkafa Then 'Берем все шкафы с именем того, на который вставляем элемент
'                        Select Case ShapeSAType(vsoShapeOnPage) 'Если в шейпе есть тип, то -
'                            Case typeSensor, typeActuator
'                                'Собираем в коллекцию нужные для сортировки шейпы
'                                colShpPage.Add vsoShapeOnPage
'                            Case Else
'                        End Select
'                    End If
'                End If
'            Next
'
'            'Сортируем то что нашли на листе
'
'            'из коллекции передаем в массив для сортировки
'            If colShpPage.Count > 0 Then
'                ReDim shpMas(colShpPage.Count - 1)
'                i = 0
'                For Each shpElement In colShpPage
'                    Set shpMas(i) = shpElement
'                    i = i + 1
'                Next
'
'                ' "Сортировка вставками" массива шейпов по возрастанию коордонаты Х
'                '--V--Сортируем по возрастанию коордонаты Х
'                UbMas = UBound(shpMas)
'                For j = 1 To UbMas
'                    Set shpTemp = shpMas(j)
'                    i = j
'                    'If shpMas(i) Is Nothing Then Exit Sub
'                    While shpMas(i - 1).Cells("PinX").Result("mm") > shpTemp.Cells("PinX").Result("mm") '>:возрастание, <:убывание
'                        Set shpMas(i) = shpMas(i - 1)
'                        i = i - 1
'                        If i <= 0 Then GoTo ExitWhileX
'                    Wend
'ExitWhileX:                  Set shpMas(i) = shpTemp
'                Next
'                '--Х--Сортировка по возрастанию коордонаты Х
'
'                'Собираем отсортированные листы в коллекцию документа
'                For i = 0 To UbMas
'                    colShpDoc.Add shpMas(i)
'                Next
'                Set colShpPage = Nothing
'            End If
'        End If
'    Next
'
'    If colShpDoc.Count > 0 Then
'        'Берем первую страницу СВП
'        Set vsoPage = ActivePage 'GetSAPageExist(cListNameSVP) 'ActiveDocument.Pages(cListNameSVP)
''        If vsoPage Is Nothing Then Set vsoPage = AddSAPage(cListNameSVP)
'        SetPageSVP vsoPage
'        'Вставляем на лист СВП найденные и отсортированные датчики/приводы
'        For i = 1 To colShpDoc.Count
'            AddSensorOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
'            'Если лист кончился
'            If PastePoint > vsoPage.PageSheet.Cells("PageWidth").Result(0) - KonecLista Then
'                'Положение текущей страницы
'                Index = vsoPage.Index
'                'Создаем новую страницу СВП
'                Set vsoPage = AddSAPage(cListNameSVP)
'                SetPageSVP vsoPage
'                'Положение новой страницы сразу за текущей
'                vsoPage.Index = Index + 1
'                PastePoint = NachaloVstavki
'                'Вставляем этот же датчик только на следующем листе
'                AddSensorOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
'            End If
'        Next
'    End If
'
'    ActiveWindow.DeselectAll
'End Sub
'
'Sub SetPageSVP(vsoPage As Visio.Page)
'    Dim shpShkaf As Visio.Shape
'    'Подвал
'    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("PodvalCxemy"), 0, 0
'    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).FormulaU = "(25 mm-TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale"
'    'Шкаф
'    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("ShkafMesto"), 0, 0
'    Set shpShkaf = Application.ActiveWindow.Selection(1)
'    With shpShkaf
'        .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Formula = NachaloVstavki + Interval
'        .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Klemma - 5 / 25.4
'        .CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "37.5 mm"
'        .CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = "382.5 mm"
'        .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0"
'        .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).FormulaU = "Height * 0"
'        .CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.5 mm"
'        .CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "1"
'        .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "24 pt"
'        .CellsSRC(visSectionObject, visRowTextXForm, visXFormLocPinX).FormulaU = "TxtWidth * 0.5"
'        .Cells("Controls.TextPos").FormulaU = "Width * 0.5"
'        .AddSection visSectionConnectionPts
'        .AddRow visSectionConnectionPts, visRowLast, visTagDefault
'        .CellsSRC(visSectionConnectionPts, 0, visCnnctX).FormulaForceU = "Width*1"
'        .CellsSRC(visSectionConnectionPts, 0, visCnnctY).FormulaForceU = "Height*0"
'        .CellsSRC(visSectionConnectionPts, 0, visCnnctDirX).FormulaForceU = "0 mm"
'        .CellsSRC(visSectionConnectionPts, 0, visCnnctDirY).FormulaForceU = "0 mm"
'        .CellsSRC(visSectionConnectionPts, 0, visCnnctType).FormulaForceU = "0 mm"
'        .CellsSRC(visSectionConnectionPts, 0, visCnnctAutoGen).FormulaForceU = "0 mm"
'        .CellsSRC(visSectionConnectionPts, 0, 6).FormulaForceU = ""
'    End With
'    'Клеммник
'    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("klemmnik"), 0, 0
'    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Formula = NachaloVstavki + Interval
'    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Klemma - 5 / 25.4
'    Application.ActiveWindow.Selection(1).Cells("Controls.Line").GlueTo shpShkaf.Cells("Connections.X1")
'End Sub
'
'Public Sub svpDEL()
''------------------------------------------------------------------------------------------------------------
'' Macros        : svpDEL - Удаляет листы схемы внешних проводок
''------------------------------------------------------------------------------------------------------------
'    If MsgBox("Удалить листы схемы внешних проводок?", vbQuestion + vbOKCancel, "САПР-АСУ: Удалить листы СВП") = vbOK Then
'        del_pages cListNameSVP
'        'MsgBox "Старая версия спецификации удалена", vbInformation
'    End If
'End Sub
'
'
'