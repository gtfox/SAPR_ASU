
'------------------------------------------------------------------------------------------------------------
' Module        : KabeliPLAN - Кабели на планах
' Author        : gtfox
' Date          : 2020.10.09/2022.05.03(Дейкстра)
' Description   : Автопрокладка кабелей по лоткам, подсчет длины, выноски на плане
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Dim selLines As Visio.Selection
Dim vsoShape As Visio.Shape
Dim colShapePoints As Collection
Dim PointNumber As Integer
Dim AntiScale As Double
Const kRound = 4 'Округление коорданат (число цифр после запятой)

Sub AddRouteCablesOnPlan()
'------------------------------------------------------------------------------------------------------------
' Macros        : AddRouteCablesOnPlan - Прокладывает кабели для всех датчиков по ближайшим лоткам
'------------------------------------------------------------------------------------------------------------
    Dim vsoShape As Visio.Shape
    Dim shpSensor As Visio.Shape
    Dim shpKabel As Visio.Shape
    Dim shpSensorFSA As Visio.Shape
    Dim colCables As Collection
    Dim colCablesTemp As Collection
    Dim vsoCollection As Collection
    Dim colShapeFSA As Collection
    Dim nCount As Double

    Set colShapeFSA = New Collection
    'Находим датчики
    For Each shpSensorFSA In ActivePage.Shapes
        If ShapeSATypeIs(shpSensorFSA, typePlanSensor) Then
            colShapeFSA.Add shpSensorFSA
        End If
    Next
    
    'Находим датчики и прокладываем кабели
    For Each shpSensorFSA In colShapeFSA
        If ShapeSATypeIs(shpSensorFSA, typePlanSensor) Then
            Set colCables = New Collection
            Set colCablesTemp = New Collection

            'Находим датчик на схеме
            Set shpSensor = ShapeByHyperLink(shpSensorFSA.Cells("Hyperlink.Shema.SubAddress").ResultStr(0))
            If Not shpSensor Is Nothing Then

                'Находим кабели на плане (чтобы не проложить повторно)
                For Each shpKabel In shpSensorFSA.ContainingPage.Shapes 'Перебираем все кабели
                    If ShapeSATypeIs(shpKabel, typeCablePL) Then
                        colCablesTemp.Add shpKabel, IIf(shpKabel.Cells("Prop.SymName").ResultStr(0) = "", CStr(shpKabel.Cells("Prop.Number").Result(0)), shpKabel.Cells("Prop.SymName").ResultStr(0) & shpKabel.Cells("Prop.Number").Result(0))
                    End If
                Next

                'Находим кабель/кабели подключенные к датчику исключая существующие(уже проложенные)
                For Each vsoShape In shpSensor.Shapes 'Перебираем все входы датчика
                    If ShapeSATypeIs(vsoShape, typeSensorIO) Then
                        'Находим подключенные провода
                        Set vsoCollection = FillColWires(vsoShape)
                        nCount = colCablesTemp.Count
                        On Error Resume Next
                        colCablesTemp.Add vsoCollection.Item(1).Parent, IIf(vsoCollection.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), vsoCollection.Item(1).Parent.Cells("Prop.SymName").ResultStr(0) & vsoCollection.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(vsoCollection.Item(1).Parent.Cells("Prop.Number").Result(0)))
                        If colCablesTemp.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                            colCables.Add vsoCollection.Item(1).Parent
                            nCount = colCablesTemp.Count
                        End If
                    End If
                Next
                'Отключаем On Error Resume Next
                err.Clear
                On Error GoTo 0

                'Прокладываем кабель/кабели для датчика
                If colCables.Count > 0 Then
                    RouteCable shpSensorFSA
                    DoEvents
                End If

            End If
        End If
    Next
End Sub

Sub RouteCable(shpSensorFSA As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : RouteCable - Прокладывает кабель по ближайшему лотку
                'Определяет ближайший лоток и прокладывает кабель до шкафа
'------------------------------------------------------------------------------------------------------------
    
    Dim shpKabel As Visio.Shape
    Dim shpKabelPL As Visio.Shape
    Dim shpKabelPLPattern As Visio.Shape

    Dim colCables As Collection
    Dim colCablesTemp As Collection
    
    Dim shpLotok As Visio.Shape
    Dim shpSensor As Visio.Shape
    Dim shpSensorFSATemp As Visio.Shape
    Dim shpShortLine As Visio.Shape

    Dim vsoShapeTemp As Visio.Shape
    Dim vsoCollection As Collection
    
    Dim shpLine As Visio.Shape
    Dim shpLineUp As Visio.Shape
    Dim shpLineDown As Visio.Shape
    Dim shpLineLeft As Visio.Shape
    Dim shpLineRight As Visio.Shape
    Dim selLineUp As Visio.Selection
    Dim selLineDown As Visio.Selection
    Dim selLineLeft As Visio.Selection
    Dim selLineRight As Visio.Selection
    Dim selSelection As Visio.Selection
    Dim selSelectionTemp As Visio.Selection
    
    Dim colLine As Collection
    Dim colLotok As Collection
    Dim colLineShort As Collection
    
    Dim StartRoute As Long
    Dim EndRoute As Long
    Dim BeginX As Double
    Dim BeginY As Double
    Dim EndX As Double
    Dim EndY As Double
    Dim PinX As Double
    Dim PinY As Double

    Dim vsoLayer1 As Visio.Layer
    Dim vsoLayer2 As Visio.Layer
    Dim vsoLayer3 As Visio.Layer
    Dim vsoLayer4 As Visio.Layer
    
    Dim SensorFSAPinX As Double
    Dim SensorFSAPinY As Double
    Dim dXSensorFSAPinX As Double
    Dim dYSensorFSAPinY As Double
    Dim PageWidth As Double
    Dim PageHeight As Double

    Dim DlinaKabelya As Double
    Dim nCount As Double
    Dim Key As String
    
    Dim BoxNumber As Integer 'Номер шкафа к которому подключен кабель/датчик
    Dim NazvanieShemy As String 'Название схемы шкафа к которому подключен кабель/датчик
    Dim i As Integer
    Dim n As Integer
    Dim MultiCable As Boolean

    Dim colShkafov As Collection
    Dim LastPointNumber As Integer
    Dim clsShapePoint As classShapePoint
    Dim clsShpPnt As classShapePoint

    Dim graph() As Vertex
    Dim masRoute()
    Dim clsLotokFSA As classLotokFSA

    AntiScale = ActivePage.PageSheet.Cells("DrawingScale").Result(0) / ActivePage.PageSheet.Cells("PageScale").Result(0)
    
    Set vsoCollection = New Collection
    Set colShkafov = New Collection
    Set colShapePoints = New Collection
    
    Set vsoLayer3 = Application.ActiveWindow.Page.Layers.Add("temp3") 'слой для коротких линий(для удаления)
    Set vsoLayer4 = Application.ActiveWindow.Page.Layers.Add("SA_Kabeli") 'слой для кабелей
    
    'Находим шкафы и точки их подключения
    For Each shpLotok In ActivePage.Shapes
        If ShapeSATypeIs(shpLotok, typeDuctPlan) Then
            For i = 1 To shpLotok.Connects.Count 'Перебираем подключенные концы лотка
                If ShapeSATypeIs(shpLotok.Connects(i).ToSheet, typeBox) Then 'Выбираем только шкафы
                    Set clsShapePoint = New classShapePoint
                    clsShapePoint.PointNumber = colShapePoints.Count + 1
                    Select Case shpLotok.Connects(i).FromPart
                        Case visBegin
                            clsShapePoint.x = Round(shpLotok.Cells("BeginX").Result(0), kRound)
                            clsShapePoint.y = Round(shpLotok.Cells("BeginY").Result(0), kRound)
                        Case visEnd
                            clsShapePoint.x = Round(shpLotok.Cells("EndX").Result(0), kRound)
                            clsShapePoint.y = Round(shpLotok.Cells("EndY").Result(0), kRound)
                    End Select
                    Set clsShapePoint.ShapeOnFSA = shpLotok.Connects(i).ToSheet
                    colShapePoints.Add clsShapePoint, CStr(clsShapePoint.PointNumber)
                End If
            Next
        End If
    Next

    'Находим точку подключения датчика
    Set clsShapePoint = New classShapePoint
    clsShapePoint.PointNumber = colShapePoints.Count + 1
    clsShapePoint.x = Round(shpSensorFSA.Cells("PinX").Result(0), kRound)
    clsShapePoint.y = Round(shpSensorFSA.Cells("PinY").Result(0), kRound)
    Set clsShapePoint.ShapeOnFSA = shpSensorFSA
    colShapePoints.Add clsShapePoint, CStr(clsShapePoint.PointNumber)

    
    'Для датчика находим кратчайшую линию
    Set colLine = New Collection
    Set colLotok = New Collection
    Set colLineShort = New Collection
    ActiveWindow.DeselectAll
    Set selSelection = ActiveWindow.Selection

    SensorFSAPinX = shpSensorFSA.Cells("PinX").Result(0)
    SensorFSAPinY = shpSensorFSA.Cells("PinY").Result(0)
    PageWidth = shpSensorFSA.ContainingPage.PageSheet.Cells("PageWidth").Result(0)
    PageHeight = shpSensorFSA.ContainingPage.PageSheet.Cells("PageHeight").Result(0)
    
    'Рисуем линии во все стороны
    Set shpLineUp = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, SensorFSAPinX, PageHeight)
    Set shpLineDown = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, SensorFSAPinX, 0)
    Set shpLineLeft = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, 0, SensorFSAPinY)
    Set shpLineRight = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, PageWidth, SensorFSAPinY)
    
    'Находим все пересечения
    Set selLineUp = shpLineUp.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
    Set selLineDown = shpLineDown.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
    Set selLineLeft = shpLineLeft.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
    Set selLineRight = shpLineRight.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
    
    'Выбираем лотки и линии
    AddLotokToCol shpLineUp, selLineUp, colLine, colLotok ', NazvanieShemy 'BoxNumber
    AddLotokToCol shpLineDown, selLineDown, colLine, colLotok ', NazvanieShemy 'BoxNumber
    AddLotokToCol shpLineLeft, selLineLeft, colLine, colLotok ', NazvanieShemy 'BoxNumber
    AddLotokToCol shpLineRight, selLineRight, colLine, colLotok ', NazvanieShemy 'BoxNumber
    If colLotok.Count = 0 Then 'нет лотков - выходим
        'Чистим вспомогательную графику
        shpLineUp.Delete
        shpLineDown.Delete
        shpLineLeft.Delete
        shpLineRight.Delete
        MsgBox "Нет лотков поблизости или лоток не приклеен к ящику", vbCritical + vbOKOnly, "САПР-АСУ: Ошибка"
        Exit Sub
    End If

    'Выделяем их
    selSelection.Select shpSensorFSA, visSelect
    For Each vsoShape In colLine
        selSelection.Select vsoShape, visSelect
    Next
    For Each vsoShape In colLotok
        selSelection.Select vsoShape, visSelect
    Next

    'Копируем и вставляем на временном слое
    selSelection.Copy
    Set vsoLayer1 = Application.ActiveWindow.Page.Layers.Add("temp") 'новый слой
    vsoLayer1.CellsC(visLayerActive).FormulaU = "1" 'активируем
    Application.ActiveWindow.Page.Paste

    'Находим смещение вставленного, относительно копированного
    For Each vsoShape In ActiveWindow.Selection
        If ShapeSATypeIs(vsoShape, typePlanSensor) Then
            dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
            dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
            Set shpSensorFSATemp = vsoShape
        End If
        vsoShape.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
        vsoLayer1.Add vsoShape, 0 'Добавляем все на временный слой
    Next
    'и сдвигаем на место
    ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY
    
    'Отключаем события автоматизации
    Application.EventsEnabled = 0
    shpSensorFSATemp.Delete 'убираем лишнее перед trim
    'Включаем события автоматизации
    Application.EventsEnabled = -1
    
    ActiveWindow.Selection.Trim 'разбиваем

    'находим ближайшие линии которые касаются лотка
    Set selLines = ActivePage.SpatialSearch(SensorFSAPinX, SensorFSAPinY, visSpatialTouching, 0.02 * AntiScale, 0)
    For Each vsoShape In selLines
        If vsoShape.LayerCount > 0 Then
            If vsoShape.Layer(1).name = vsoLayer1.name Then
                'Проверяем что линия касается лотка
                Set selSelectionTemp = vsoShape.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
                For Each vsoShapeTemp In selSelectionTemp
                    If ShapeSATypeIs(vsoShapeTemp, typeDuctPlan) Then
                        colLineShort.Add vsoShape
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    'находим самую короткую
    If colLineShort.Count = 0 Then
        MsgBox "Нет лотков поблизости от " & shpSensorFSA.Cells("User.NameParent").ResultStr(0) & " (" & shpSensorFSA.NameID & ")", vbExclamation, "САПР-АСУ: Ошибка"
    ElseIf colLineShort.Count = 1 Then
        Set shpShortLine = colLineShort.Item(1)
    ElseIf colLineShort.Count > 1 Then
        Set shpShortLine = colLineShort.Item(1)
        For i = 2 To colLineShort.Count
            If colLineShort.Item(i).Cells("Width").Result(0) < shpShortLine.Cells("Width").Result(0) Then
                Set shpShortLine = colLineShort.Item(i)
            End If
        Next
    End If
    
    'Убираем ее с временного слоя
     vsoLayer1.Remove shpShortLine, 0
     
    'Чистим вспомогательную графику
    shpLineUp.Delete
    shpLineDown.Delete
    shpLineLeft.Delete
    shpLineRight.Delete
    vsoLayer1.Delete True
    
    'Создаем свойства для линии (тип как у лотка typeDuctPlan = 170)
    With shpShortLine
        .AddSection visSectionUser
        .AddRow visSectionUser, visRowLast, visTagDefault
        .CellsSRC(visSectionUser, visRowLast, visUserValue).RowNameU = "SAType"
        .CellsSRC(visSectionUser, visRowLast, visUserValue).FormulaForceU = "170"
        .AddSection visSectionProp
        .AddRow visSectionProp, visRowLast, visTagDefault
        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "SymName"
        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Букв. обозначение"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Букв. обозначение"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "0"
        .AddRow visSectionProp, visRowLast, visTagDefault
        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "Ac3"
        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Номер кабеля"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Номер кабеля"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "2"
        .AddRow visSectionProp, visRowLast, visTagDefault
        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "Dlina"
        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Длина кабеля, м."""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Длина кабеля, м."""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "2"
    End With
    
    shpShortLine.Cells("Prop.SymName").Formula = """G"""
    shpShortLine.Cells("Prop.Ac3").Formula = """1"""

    vsoLayer3.Add shpShortLine, 0

    'Берем все лотки и кратчайшие линии
    ActiveWindow.DeselectAll
    Set selSelection = ActiveWindow.Selection
    
    For Each shpLotok In ActivePage.Shapes
        If ShapeSATypeIs(shpLotok, typeDuctPlan) Then
            selSelection.Select shpLotok, visSelect
        End If
    Next
    
    'Добавляем туда датчик для поиска смещения
    SensorFSAPinX = shpSensorFSA.Cells("PinX").Result(0)
    SensorFSAPinY = shpSensorFSA.Cells("PinY").Result(0)
    selSelection.Select shpSensorFSA, visSelect

    'Копируем и вставляем на временном слое
    selSelection.Copy
    Set vsoLayer1 = Application.ActiveWindow.Page.Layers.Add("temp") 'новый слой
    vsoLayer1.CellsC(visLayerActive).FormulaU = "1" 'активируем
    Application.ActiveWindow.Page.Paste
     
    'Находим смещение вставленного, относительно копированного
    For Each vsoShape In ActiveWindow.Selection
        If ShapeSATypeIs(vsoShape, typePlanSensor) Then
            dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
            dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
            'Отключаем события автоматизации
            Application.EventsEnabled = 0
            vsoShape.Delete 'убираем временный датчик
            'Включаем события автоматизации
            Application.EventsEnabled = -1
            Exit For
        End If
    Next
    'и сдвигаем на место
    ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY
    'разбиваем
    ActiveWindow.Selection.Trim 'разбиваем
    
    'Создаем из линий маршруты
    Set selSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer1.name)
    For Each shpLine In selSelection
        SetRoute shpLine
    Next
    LastPointNumber = colShapePoints.Count
    PointNumber = 1

    'Сканируем маршруты в этой точке (максимум 8(4 стороны света +45 градусов))находим ближайшие линии + Заполняем пути именами точек
    Set selLines = ActivePage.SpatialSearch(colShapePoints(1).x, colShapePoints(1).y, visSpatialTouching, 0.02 * AntiScale, 0)
    For Each vsoShape In selLines
        If vsoShape.LayerCount > 0 Then
            If vsoShape.Layer(1).name = vsoLayer1.name Then
                'Находим точки начала и конца
                If vsoShape.OneD Then '1-D фигура
                    'Находим точки начала и конца линии в 1D фигуре
                    BeginX = Round(vsoShape.Cells("BeginX").Result(0), kRound)
                    BeginY = Round(vsoShape.Cells("BeginY").Result(0), kRound)
                    EndX = Round(vsoShape.Cells("EndX").Result(0), kRound)
                    EndY = Round(vsoShape.Cells("EndY").Result(0), kRound)
                Else '2-D фигура
                    'Находим точки начала и конца линии в 2D фигуре
                    BeginX = Round(vsoShape.Cells("PinX").Result(0) - vsoShape.Cells("Width").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 0).Result(0), kRound)
                    BeginY = Round(vsoShape.Cells("PinY").Result(0) - vsoShape.Cells("Height").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 1).Result(0), kRound)
                    EndX = Round(vsoShape.Cells("PinX").Result(0) - vsoShape.Cells("Width").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowLast, 0).Result(0), kRound)
                    EndY = Round(vsoShape.Cells("PinY").Result(0) - vsoShape.Cells("Height").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowLast, 1).Result(0), kRound)
                End If
                
                vsoShape.Cells("User.BeginX").Formula = BeginX
                vsoShape.Cells("User.BeginY").Formula = BeginY
                vsoShape.Cells("User.EndX").Formula = EndX
                vsoShape.Cells("User.EndY").Formula = EndY
                
                'Именуем начало или ...
                If BeginX = colShapePoints(1).x And BeginY = colShapePoints(1).y Then
                    vsoShape.Cells("Prop.Begin").Formula = PointNumber
                
                '... именуем конец в этой точке
                ElseIf EndX = colShapePoints(1).x And EndY = colShapePoints(1).y Then
                    vsoShape.Cells("Prop.End").Formula = PointNumber
                End If
                
                'Заполняем свойство длина
                vsoShape.Cells("Prop.Dlina").Formula = CableLength(vsoShape)
                
                'Заполняем пути именами точек
                FillRoute vsoShape, vsoLayer1
            End If
        End If
    Next

    'Датчику и шкафам присваиваем номера точек маршрутов
    For Each clsShapePoint In colShapePoints
        Set selLines = ActivePage.SpatialSearch(clsShapePoint.x, clsShapePoint.y, visSpatialTouching, 0.02 * AntiScale, 0)
        
        'Перебираем маршруты
        For Each shpRoute In selLines
            If shpRoute.LayerCount > 0 Then
                If shpRoute.Layer(1).name = vsoLayer1.name Then
                    If shpRoute.Cells("User.BeginX").Result(0) = clsShapePoint.x And shpRoute.Cells("User.BeginY").Result(0) = clsShapePoint.y Then
                        clsShapePoint.PointNumber = shpRoute.Cells("Prop.Begin").Result(0)
                    ElseIf shpRoute.Cells("User.EndX").Result(0) = clsShapePoint.x And shpRoute.Cells("User.EndY").Result(0) = clsShapePoint.y Then
                        clsShapePoint.PointNumber = shpRoute.Cells("Prop.End").Result(0)
                    End If
                End If
            End If
        Next
    Next
    
    'Создаем граф маршрутов
    MakeGraph graph, vsoLayer1

    'Находим датчик ФСА в коллекции
    For Each clsShapePoint In colShapePoints
        If ShapeSATypeIs(clsShapePoint.ShapeOnFSA, typePlanSensor) Then
            Set colCables = New Collection
            Set colCablesTemp = New Collection
            Set shpSensorFSATemp = clsShapePoint.ShapeOnFSA
            
            'Находим датчик на схеме
            Set shpSensor = ShapeByHyperLink(shpSensorFSATemp.Cells("Hyperlink.Shema.SubAddress").ResultStr(0))
            If Not shpSensor Is Nothing Then
                MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)
            Else
                vsoLayer1.Delete True
                vsoLayer3.Delete True
                MsgBox "Датчик " & shpSensor.Cells("User.NameParent").ResultStr(0) & " (" & shpSensor.NameID & ") не связан", vbExclamation, "САПР-АСУ: Ошибка"
                Exit Sub
            End If
            
            'Находим кабели на плане (чтобы не проложить повторно)
            For Each shpKabel In shpSensorFSATemp.ContainingPage.Shapes 'Перебираем все кабели
                If ShapeSATypeIs(shpKabel, typeCablePL) Then
                    colCablesTemp.Add shpKabel, IIf(shpKabel.Cells("Prop.SymName").ResultStr(0) = "", CStr(shpKabel.Cells("Prop.Number").Result(0)), shpKabel.Cells("Prop.SymName").ResultStr(0) & shpKabel.Cells("Prop.Number").Result(0))
                End If
            Next
            
            'Находим кабель/кабели подключенные к датчику исключая существующие(уже проложенные)
            For Each vsoShape In shpSensor.Shapes 'Перебираем все входы датчика
                If ShapeSATypeIs(vsoShape, typeSensorIO) Then
                    'Находим подключенные провода
                    Set vsoCollection = FillColWires(vsoShape)
                    nCount = colCablesTemp.Count
                    On Error Resume Next
                    colCablesTemp.Add vsoCollection.Item(1).Parent, IIf(vsoCollection.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), vsoCollection.Item(1).Parent.Cells("Prop.SymName").ResultStr(0) & vsoCollection.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(vsoCollection.Item(1).Parent.Cells("Prop.Number").Result(0)))
                    If colCablesTemp.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                        colCables.Add vsoCollection.Item(1).Parent
                        nCount = colCablesTemp.Count
                    End If
                End If
            Next
            
            'Отключаем On Error Resume Next
            err.Clear
            On Error GoTo 0
            
            If colCables.Count = 0 Then Exit Sub 'MsgBox "Не найдены кабели", vbExclamation + vbOKOnly, "САПР-АСУ: Info": Exit Sub
            'Шкаф к которому подключен кабель (Предполагается что 1 датчик подключен к 1 шкафу (даже многокабельный)
        '    BoxNumber = colCables.Item(1).Cells("User.LinkToBox").Result(0)
        '    NazvanieShemy = colCables.Item(1).ContainingPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0)
            NazvanieShemy = colCables.Item(1).Cells("User.LinkToBox").ResultStr(0)
            
            'Номер точки начала машрута
            StartRoute = clsShapePoint.PointNumber
            
            'Находим шкаф по названию схемы
            For Each clsShpPnt In colShapePoints
                If ShapeSATypeIs(clsShpPnt.ShapeOnFSA, typeBox) Then
                    If clsShpPnt.ShapeOnFSA.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then
                        EndRoute = clsShpPnt.PointNumber 'Номер точки конца машрута
                        Exit For
                    End If
                End If
            Next
            If EndRoute = 0 Then
                vsoLayer1.Delete True
                vsoLayer3.Delete True
                MsgBox "Нет шкафа " & NazvanieShemy & " для датчика " & clsShapePoint.ShapeOnFSA.Cells("User.NameParent").ResultStr(0) & " (" & clsShapePoint.ShapeOnFSA.NameID & ")", vbCritical + vbOKOnly, "САПР-АСУ: Ошибка"
                Exit Sub
            End If
            
            'Очищаем предыдущий маршрут
            For i = 1 To UBound(graph, 1)
                graph(i).d = INF
                graph(i).p = 0
                graph(i).id = 0
                graph(i).u = False
            Next
            
            'Находим кратчайший маршрут по алгоритму Дейкстры
            masRoute = MyDijkstra(graph, StartRoute, EndRoute)
            
            'Перебираем куски маршрута
            ActiveWindow.DeselectAll
            Set selLines = ActiveWindow.Selection
            Set colLotok = New Collection
            For i = 1 To UBound(masRoute, 1)
                If IsEmpty(masRoute(i, 1)) Then Exit For
                On Error GoTo er1 'Маршрут 3-5 или 5-3
                Set shpRoute = ActivePage.Shapes(masRoute(i, 1) & "-" & masRoute(i, 2))

                'Находим точку на куске маршрута
                If shpRoute.OneD Then '1-D фигура
                    'Точка по середине линии
                    PinX = Round(shpRoute.Cells("PinX").Result(0), kRound)
                    PinY = Round(shpRoute.Cells("PinY").Result(0), kRound)
                Else '2-D фигура
                    'Точка на первом изгибе 2D фигуры
                    PinX = Round(shpRoute.Cells("PinX").Result(0) - shpRoute.Cells("Width").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowFirst + 2, 0).Result(0), kRound)
                    PinY = Round(shpRoute.Cells("PinY").Result(0) - shpRoute.Cells("Height").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowFirst + 2, 1).Result(0), kRound)
                End If
                
                'Находим лоток под куском маршрута
                Set selSelection = ActivePage.SpatialSearch(PinX, PinY, visSpatialTouching, 0.02 * AntiScale, 0)
                For Each vsoShape In selSelection
                    If ShapeSATypeIs(vsoShape, typeDuctPlan) Then
                        Set shpLotok = vsoShape
                    End If
                Next
    
                'Собираем куски лотков для КЖ и их длины
                Set clsLotokFSA = New classLotokFSA
                clsLotokFSA.NameLotok = shpLotok.Cells("Prop.SymName").ResultStr(0) & " " & shpLotok.Cells("Prop.Ac3").ResultStr(0)
                clsLotokFSA.DlinaLotok = shpRoute.Cells("Prop.Dlina").Result(0)
                Key = shpLotok.Cells("Prop.SymName").ResultStr(0) & shpLotok.Cells("Prop.Ac3").ResultStr(0)
                nCount = colLotok.Count
                On Error Resume Next
                colLotok.Add clsLotokFSA, Key
                If colLotok.Count = nCount Then 'Если кол-во не увеличелось, значит оно уже есть - складываем длину
                    colLotok(Key).DlinaLotok = colLotok(Key).DlinaLotok + shpRoute.Cells("Prop.Dlina").Result(0)
                End If
                
                'Отключаем On Error Resume Next
                err.Clear
                On Error GoTo 0
                
                'Выделяем куски маршрута
                selLines.Select shpRoute, visSelect
            Next

            'Выделяем датчик для поиска смещения
            selLines.Select shpSensorFSATemp, visSelect
            
            'Сохраняем координаты датчика
            SensorFSAPinX = shpSensorFSATemp.Cells("PinX").Result(0)
            SensorFSAPinY = shpSensorFSATemp.Cells("PinY").Result(0)
                        
            'Копируем и вставляем на временном слое
            selLines.Copy
            Set vsoLayer2 = Application.ActiveWindow.Page.Layers.Add("temp2") 'Временный слой
            vsoLayer2.CellsC(visLayerActive).FormulaU = "1" 'активируем
            Application.ActiveWindow.Page.Paste
            
            'Находим смещение вставленного, относительно копированного
            For Each vsoShape In ActiveWindow.Selection
                If ShapeSATypeIs(vsoShape, typePlanSensor) Then
                    dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
                    dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
                    'Отключаем события автоматизации
                    Application.EventsEnabled = 0
                    vsoShape.Delete 'убираем временный датчик
                    'Включаем события автоматизации
                    Application.EventsEnabled = -1
                    Exit For
                End If
            Next
            ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY 'и сдвигаем на место

            'соединяем
            Application.ActiveWindow.Selection.Join
            Set selSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer2) 'выделяем все в слое
            Set shpKabelPLPattern = selSelection.PrimaryItem 'Таки профит! Гребаный кабель случился!
            'Убираем с временных слоёв
            shpKabelPLPattern.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
            'Чистим вспомогательную графику
            vsoLayer2.Delete True
            
            'Считаем длину кабеля
            DlinaKabelya = CableLength(shpKabelPLPattern)
            
            'Создаем свойтва шаблона кабеля на плане
            SetGofra shpKabelPLPattern
            'Второй раз, чтобы записались перекрестные формулы в разделах User. и Prop.
            SetGofra shpKabelPLPattern
            
            With shpKabelPLPattern
                .CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = 1 'Обычная линия
                .CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.2 mm"
            End With
            
            'Перебираем все кабели в датчике
            For Each shpKabel In colCables
                Set shpKabelPL = shpKabelPLPattern.Duplicate
                'Сдвигаем на место
                shpKabelPL.Cells("PinX").Formula = shpKabelPLPattern.Cells("PinX").Result(0)
                shpKabelPL.Cells("PinY").Formula = shpKabelPLPattern.Cells("PinY").Result(0)
                'На задний план
                Application.ActiveWindow.Selection.SendToBack
                'Переименовываем кабель на плане и заполняем свойства
                With shpKabelPL
                    .name = "KabelPL." & .id
                    .Cells("Prop.SymName").Formula = IIf(shpKabel.Cells("Prop.BukvOboz").Result(0), """" & shpKabel.Cells("Prop.SymName").ResultStr(0) & """", """""")
                    .Cells("Prop.Number").Formula = shpKabel.Cells("Prop.Number").Result(0)
                    .Cells("Prop.Dlina").Formula = DlinaKabelya
                End With
                
                'Заполняем куски маршрута для КЖ на плане
                shpKabelPL.AddSection visSectionScratch
                For i = 1 To colLotok.Count
                    shpKabelPL.AddRow visSectionScratch, visRowLast, visTagDefault
                    If colLotok.Item(i).NameLotok = "G 1" Then
                        shpKabelPL.CellsSRC(visSectionScratch, visRowLast, visScratchA).FormulaU = "User.FullName"
                    Else
                        shpKabelPL.CellsSRC(visSectionScratch, visRowLast, visScratchA).FormulaU = """" & colLotok.Item(i).NameLotok & """"
                    End If
                    shpKabelPL.CellsSRC(visSectionScratch, visRowLast, visScratchB).FormulaU = """" & colLotok.Item(i).DlinaLotok & """"
                Next
                
                'Переносим на слой кабелей
                vsoLayer4.Add shpKabelPL, 0
                
                'Заполняем длину кабеля на эл.схеме (длина кабеля СВП ссылается формулой на эл.сх.)
                shpKabel.Cells("Prop.Dlina").FormulaU = "Pages[" + shpKabelPL.ContainingPage.NameU + "]!" + shpKabelPL.NameID + "!Prop.Dlina"
                shpKabel.Cells("Hyperlink.Kabel.SubAddress").FormulaU = """" + shpKabelPL.ContainingPage.NameU + "/" + shpKabelPL.NameID + """"
                shpKabel.Cells("Hyperlink.Kabel.ExtraInfo").FormulaU = """" + shpKabelPL.Cells("Prop.SymName").ResultStr(0) + CStr(shpKabelPL.Cells("Prop.Number").Result(0)) + """"
                
                vsoLayer1.Remove shpKabelPL, 0
            Next
            
            'Удаляем шаблон кабеля
            shpKabelPLPattern.Delete
            
            Application.ActiveWindow.DeselectAll
            
        End If
    Next
    
    vsoLayer1.Delete True
    vsoLayer3.Delete True
    
    'Включаем события автоматизации
    Application.EventsEnabled = -1
    
    Exit Sub
    
er1: 'Маршрут 3-5 или 5-3
    Set shpRoute = ActivePage.Shapes(masRoute(i, 2) & "-" & masRoute(i, 1))
Resume Next

End Sub

Sub FillRoute(shpRouteToPoint As Visio.Shape, vsoLayer As Visio.Layer)
'------------------------------------------------------------------------------------------------------------
' Macros        : FillRoute - Заполняет пути именами точек и длиной (рекурсивная)
'------------------------------------------------------------------------------------------------------------
    Dim shpRoute As Visio.Shape
    Dim colRoute As Collection
    Dim BeginX As Double
    Dim BeginY As Double
    Dim EndX As Double
    Dim EndY As Double
    Dim clsPoint As classPoint
    
    Set clsPoint = New classPoint

    If shpRouteToPoint.OneD Then '1-D фигура
        'Находим точки начала и конца линии в 1D фигуре
        BeginX = Round(shpRouteToPoint.Cells("BeginX").Result(0), kRound)
        BeginY = Round(shpRouteToPoint.Cells("BeginY").Result(0), kRound)
        EndX = Round(shpRouteToPoint.Cells("EndX").Result(0), kRound)
        EndY = Round(shpRouteToPoint.Cells("EndY").Result(0), kRound)
    Else '2-D фигура
        'Находим точки начала и конца линии в 2D фигуре
        BeginX = Round(shpRouteToPoint.Cells("PinX").Result(0) - shpRouteToPoint.Cells("Width").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 0).Result(0), kRound)
        BeginY = Round(shpRouteToPoint.Cells("PinY").Result(0) - shpRouteToPoint.Cells("Height").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 1).Result(0), kRound)
        EndX = Round(shpRouteToPoint.Cells("PinX").Result(0) - shpRouteToPoint.Cells("Width").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowLast, 0).Result(0), kRound)
        EndY = Round(shpRouteToPoint.Cells("PinY").Result(0) - shpRouteToPoint.Cells("Height").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowLast, 1).Result(0), kRound)
    End If
    
    shpRouteToPoint.Cells("User.BeginX").Formula = BeginX
    shpRouteToPoint.Cells("User.BeginY").Formula = BeginY
    shpRouteToPoint.Cells("User.EndX").Formula = EndX
    shpRouteToPoint.Cells("User.EndY").Formula = EndY
    
    If shpRouteToPoint.Cells("Prop.Begin").Result(0) = 0 Or shpRouteToPoint.Cells("Prop.End").Result(0) = 0 Then
        'Находим точку на другом конце
        PointNumber = PointNumber + 1
        clsPoint.PointNumber = PointNumber
        If shpRouteToPoint.Cells("Prop.Begin").Result(0) = 0 Then
            clsPoint.x = BeginX
            clsPoint.y = BeginY
            shpRouteToPoint.Cells("Prop.Begin").Formula = clsPoint.PointNumber
        ElseIf shpRouteToPoint.Cells("Prop.End").Result(0) = 0 Then
            clsPoint.x = EndX
            clsPoint.y = EndY
            shpRouteToPoint.Cells("Prop.End").Formula = clsPoint.PointNumber
        End If
        shpRouteToPoint.name = shpRouteToPoint.Cells("Prop.Begin").Result(0) & "-" & shpRouteToPoint.Cells("Prop.End").Result(0)
    Else
        Exit Sub
    End If
    
    'Сканируем маршруты в этой точке (максимум 8(4 стороны света +45 градусов))находим ближайшие линии
    Set selLines = ActivePage.SpatialSearch(clsPoint.x, clsPoint.y, visSpatialTouching, 0.02 * AntiScale, 0)

    Set colRoute = New Collection
    For Each vsoShape In selLines
        If vsoShape.LayerCount > 0 Then
            If vsoShape.Layer(1).name = vsoLayer.name Then
                colRoute.Add vsoShape
            End If
        End If
    Next
    
    'Перебираем маршруты
    For Each shpRoute In colRoute

        If shpRoute.OneD Then '1-D фигура
        
            'Находим точки начала и конца линии в 1D фигуре
            BeginX = Round(shpRoute.Cells("BeginX").Result(0), kRound)
            BeginY = Round(shpRoute.Cells("BeginY").Result(0), kRound)
            EndX = Round(shpRoute.Cells("EndX").Result(0), kRound)
            EndY = Round(shpRoute.Cells("EndY").Result(0), kRound)
            
            shpRoute.Cells("User.BeginX").Formula = BeginX
            shpRoute.Cells("User.BeginY").Formula = BeginY
            shpRoute.Cells("User.EndX").Formula = EndX
            shpRoute.Cells("User.EndY").Formula = EndY

            'Нет именованных концов
            If shpRoute.Cells("Prop.Begin").Result(0) = 0 And shpRoute.Cells("Prop.End").Result(0) = 0 Then
    
                'Именуем конец в этой точке
                If BeginX = clsPoint.x And BeginY = clsPoint.y Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                ElseIf EndX = clsPoint.x And EndY = clsPoint.y Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
                End If
                
                'Заполняем свойство длина
                shpRoute.Cells("Prop.Dlina").Formula = CableLength(shpRoute)
                
            'Именован один конец
            ElseIf shpRoute.Cells("Prop.Begin").Result(0) = 0 Or shpRoute.Cells("Prop.End").Result(0) = 0 Then
                'Именован другой конец - именуем наш, другой не трогаем
                'Именуем конец в этой точке
                If BeginX = clsPoint.x And BeginY = clsPoint.y And shpRoute.Cells("Prop.Begin").Result(0) = 0 Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                    shpRoute.name = shpRoute.Cells("Prop.Begin").Result(0) & "-" & shpRoute.Cells("Prop.End").Result(0)
                ElseIf EndX = clsPoint.x And EndY = clsPoint.y And shpRoute.Cells("Prop.End").Result(0) = 0 Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
                    shpRoute.name = shpRoute.Cells("Prop.Begin").Result(0) & "-" & shpRoute.Cells("Prop.End").Result(0)
    
                'Именован наш конец - исключение (мы не должны попасть в точку в которой есть именованные концы)
                ElseIf BeginX = clsPoint.x And BeginY = clsPoint.y And shpRoute.Cells("Prop.Begin").Result(0) <> 0 Then
                    MsgBox "Именованый конец в точке: " & clsPoint.PointNumber & ". Конец: " & shpRoute.Cells("Prop.Begin").Result(0) & ". Маршрут: " & shpRoute.Cells("Prop.Begin").Result(0) & " - " & shpRoute.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
                    Exit Sub
                ElseIf shpRoute.Cells("EndX").Result(0) = clsPoint.x And shpRoute.Cells("EndY").Result(0) = clsPoint.y And shpRoute.Cells("Prop.End").Result(0) <> 0 Then
                    MsgBox "Именованый конец в точке: " & clsPoint.PointNumber & ". Конец: " & shpRoute.Cells("Prop.End").Result(0) & ". Маршрут: " & shpRoute.Cells("Prop.Begin").Result(0) & " - " & shpRoute.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
                    Exit Sub
                End If
                
            'Именованы оба конца
            ElseIf shpRoute.Cells("Prop.Begin").Result(0) <> 0 And shpRoute.Cells("Prop.End").Result(0) <> 0 Then
                'Маршрут уже обработан полностью (с двух концов)
            End If
  
        Else '2-D фигура
        
            'Находим точки начала и конца линии в 2D фигуре
            BeginX = Round(shpRoute.Cells("PinX").Result(0) - shpRoute.Cells("Width").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 0).Result(0), kRound)
            BeginY = Round(shpRoute.Cells("PinY").Result(0) - shpRoute.Cells("Height").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 1).Result(0), kRound)
            EndX = Round(shpRoute.Cells("PinX").Result(0) - shpRoute.Cells("Width").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowLast, 0).Result(0), kRound)
            EndY = Round(shpRoute.Cells("PinY").Result(0) - shpRoute.Cells("Height").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowLast, 1).Result(0), kRound)

            shpRoute.Cells("User.BeginX").Formula = BeginX
            shpRoute.Cells("User.BeginY").Formula = BeginY
            shpRoute.Cells("User.EndX").Formula = EndX
            shpRoute.Cells("User.EndY").Formula = EndY

            'Нет именованных концов
            If shpRoute.Cells("Prop.Begin").Result(0) = 0 And shpRoute.Cells("Prop.End").Result(0) = 0 Then
    
                'Именуем конец в этой точке
                If BeginX = clsPoint.x And BeginY = clsPoint.y Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                ElseIf EndX = clsPoint.x And EndY = clsPoint.y Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
                End If
                
                'Заполняем свойство длина
                shpRoute.Cells("Prop.Dlina").Formula = CableLength(shpRoute)
                
            'Именован один конец
            ElseIf shpRoute.Cells("Prop.Begin").Result(0) = 0 Or shpRoute.Cells("Prop.End").Result(0) = 0 Then
                'Именован другой конец - именуем наш, другой не трогаем
                'Именуем конец в этой точке
                If BeginX = clsPoint.x And BeginY = clsPoint.y And shpRoute.Cells("Prop.Begin").Result(0) = 0 Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                    shpRoute.name = shpRoute.Cells("Prop.Begin").Result(0) & "-" & shpRoute.Cells("Prop.End").Result(0)
                ElseIf EndX = clsPoint.x And EndY = clsPoint.y And shpRoute.Cells("Prop.End").Result(0) = 0 Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
                    shpRoute.name = shpRoute.Cells("Prop.Begin").Result(0) & "-" & shpRoute.Cells("Prop.End").Result(0)
    
                'Именован наш конец - исключение (мы не должны попасть в точку в которой есть именованные концы)
                ElseIf BeginX = clsPoint.x And BeginY = clsPoint.y And shpRoute.Cells("Prop.Begin").Result(0) <> 0 Then
                    MsgBox "Именованый конец в точке: " & clsPoint.PointNumber & ". Конец: " & shpRoute.Cells("Prop.Begin").Result(0) & ". Маршрут: " & shpRoute.Cells("Prop.Begin").Result(0) & " - " & shpRoute.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
                    Exit Sub
                ElseIf EndX = clsPoint.x And EndY = clsPoint.y And shpRoute.Cells("Prop.End").Result(0) <> 0 Then
                    MsgBox "Именованый конец в точке: " & clsPoint.PointNumber & ". Конец: " & shpRoute.Cells("Prop.End").Result(0) & ". Маршрут: " & shpRoute.Cells("Prop.Begin").Result(0) & " - " & shpRoute.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
                    Exit Sub
                End If
                
            'Именованы оба конца
            ElseIf shpRoute.Cells("Prop.Begin").Result(0) <> 0 And shpRoute.Cells("Prop.End").Result(0) <> 0 Then
                'Маршрут уже обработан полностью (с двух концов)
            End If
        End If
    Next
    'Перебираем маршруты
    For Each shpRoute In colRoute
        'Заполняет пути именами точек (рекурсия)
        FillRoute shpRoute, vsoLayer
    Next
End Sub


Sub AddLotokToCol(shpLine As Visio.Shape, selLine As Visio.Selection, ByRef colLine As Collection, ByRef colLotok As Collection) ', NazvanieShemy As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddLotokToCol - Заполняет коллекции лотков и линий
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim vsoShape As Visio.Shape
    Dim shpLotok As Visio.Shape
    
    For Each vsoShape In selLine 'Шейпы в выделении
        If ShapeSATypeIs(vsoShape, typeDuctPlan) Then 'Нашли лоток
            colLotok.Add vsoShape
        End If
    Next
    If colLotok.Count > 0 Then 'Линия пересекла лоток
        colLine.Add shpLine
    End If
End Sub

Public Sub PagePLANAddElementsFrm()
    Load frmPagePLANAddElements
    frmPagePLANAddElements.Show
End Sub

Public Sub AddSensorsFSAOnPlan(NazvanieFSA As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddSensorsFSAOnPlan - Копирует все датчики из ФСА на ПЛАН
                'Копирует все датчики из ФСА на ПЛАН, если датчик уже есть, то не копирут его.
'------------------------------------------------------------------------------------------------------------
    Dim PageParent As String, NameIdParent As String, AdrParent As String
    Dim PageChild  As String, NameIdChild As String, AdrChild As String
    Dim vsoPagePlan As Visio.Page
    Dim vsoPageFSA As Visio.Page
    Dim colPagesFSA As Collection
    Dim shpSensorOnFSA As Visio.Shape
    Dim shpSensorOnPLAN As Visio.Shape
    Dim colSensorOnPLAN As Collection
    Dim colSensorToPLAN As Collection
    Dim vsoSelection As Visio.Selection
    Dim vsoGroup As Visio.Shape
    Dim nCount As Double
    Dim DropX As Double
    
    If NazvanieFSA = "" Then
        MsgBox "Нет ФСА для вставки. Название ФСА пустое", vbExclamation, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set colSensorOnPLAN = New Collection
    Set colSensorToPLAN = New Collection
    Set colPagesFSA = New Collection
    Set vsoSelection = ActiveWindow.Selection
    Set vsoPagePlan = Application.ActivePage  '.Pages("План")
    Set vsoPageFSA = ActiveDocument.Pages(cListNameFSA)

    'Берем все листы одной ФСА
    For Each vsoPageFSA In ActiveDocument.Pages
        If vsoPageFSA.name Like cListNameFSA & "*" Then
            If vsoPageFSA.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
                If vsoPageFSA.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0) = NazvanieFSA Then
                    colPagesFSA.Add vsoPageFSA
                End If
            End If
        End If
    Next

    'Находим что уже есть на плане
    For Each shpSensorOnPLAN In vsoPagePlan.Shapes
        If ShapeSATypeIs(shpSensorOnPLAN, typePlanSensor) Then
            colSensorOnPLAN.Add shpSensorOnPLAN, shpSensorOnPLAN.Cells("User.Name").ResultStr(0) '& ";" & shpSensorOnPLAN.Cells("User.NameParent").ResultStr(0)
        End If
    Next
    
    'Суем туда же все из ФСА. Одинаковое не влезает => ошибка. Что не влезло: нам оно то и нужно
    For Each vsoPageFSA In colPagesFSA
        For Each shpSensorOnFSA In vsoPageFSA.Shapes
            If ShapeSATypeIs(shpSensorOnFSA, typeFSASensor) Then
                nCount = colSensorOnPLAN.Count
                On Error Resume Next
                colSensorOnPLAN.Add shpSensorOnFSA, shpSensorOnFSA.Cells("User.Name").ResultStr(0) '& ";" & shpSensorOnFSA.Cells("User.NameParent").ResultStr(0)
                If colSensorOnPLAN.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                    colSensorToPLAN.Add shpSensorOnFSA
                    nCount = colSensorOnPLAN.Count
                End If
            End If
        Next
    Next
    
    'Очищаем коллекцию для вставляемых датчиков
    Set colSensorOnPLAN = New Collection
    
    'Копируем недостающие датчики на план
    For Each shpSensorOnFSA In colSensorToPLAN
    
        PageParent = shpSensorOnFSA.ContainingPage.NameU
        NameIdParent = shpSensorOnFSA.NameID
        AdrParent = "Pages[" + PageParent + "]!" + NameIdParent

        shpSensorOnFSA.CellsU("EventDrop").FormulaU = """"""
        shpSensorOnFSA.CellsU("EventMultiDrop").FormulaU = """"""
        
        vsoSelection.Select shpSensorOnFSA, visSelect
        vsoSelection.Copy
        
        shpSensorOnFSA.CellsU("EventDrop").FormulaU = "CALLTHIS(""ThisDocument.EventDropAutoNum"")"
        shpSensorOnFSA.CellsU("EventMultiDrop").FormulaU = "CALLTHIS(""AutoNumber.AutoNumFSA"")"

        'Активируем лист план
        ActiveWindow.Page = ActiveDocument.Pages(vsoPagePlan.name)
        'Отключаем события автоматизации (чтобы не перенумеровалось все)
        Application.EventsEnabled = 0
        'Вставляем на листе план
        ActivePage.Paste
        'Включаем события автоматизации
        Application.EventsEnabled = -1
        'Заполняем данные
        With ActiveWindow.Selection(1)
        
            PageChild = .ContainingPage.NameU
            NameIdChild = .NameID
            AdrChild = "Pages[" + PageChild + "]!" + NameIdChild
            
            shpSensorOnFSA.CellsU("Hyperlink.Plan.SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
            shpSensorOnFSA.CellsU("Hyperlink.Plan.ExtraInfo").FormulaU = AdrChild + "!User.Location"   'Pages[Схема.3]!Sheet.4!User.Location
            
            ActiveWindow.Selection.Move DropX + .Cells("Width").Result(0) * 2, 0
            DropX = DropX + .Cells("Width").Result(0) * 2
            
            .Cells("Actions.Kabel.Invisible").Formula = 0
            .Cells("Actions.AddReference.Invisible").Formula = 1
            .Cells("Prop.KanalNumber").Formula = 0
            .Cells("Prop.Kanal").Formula = 0
            .Cells("EventDblClick").Formula = ""
            .Cells("User.NameParent").Formula = AdrParent + "!User.NameParent"
            .Cells("Hyperlink.Shema.SubAddress").Formula = AdrParent + "!Hyperlink.Shema.SubAddress"
            .Cells("Hyperlink.Shema.ExtraInfo").Formula = AdrParent + "!Hyperlink.Shema.ExtraInfo"
            .Cells("Hyperlink.FSA.SubAddress").Formula = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"
            .Cells("Hyperlink.FSA.ExtraInfo").Formula = AdrParent + "!User.Location"
            .Cells("Hyperlink.Plan.SubAddress").Formula = """"""
            .Cells("Hyperlink.Plan.ExtraInfo").Formula = """"""
            .Cells("Prop.Place").Formula = AdrParent + "!Prop.Place"
            .Cells("Prop.AutoNum").Formula = 0
            .Cells("Prop.SymName").Formula = AdrParent + "!Prop.SymName"
            .Cells("Prop.Number").Formula = AdrParent + "!Prop.Number"
            .Cells("Prop.NameKontur").Formula = AdrParent + "!Prop.NameKontur"
            .Cells("Prop.Forma").Formula = AdrParent + "!Prop.Forma"
            .Cells("Prop.ResizeWithText").Formula = AdrParent + "!Prop.ResizeWithText"
            .Cells("Prop.NameParent").Formula = AdrParent + "!Prop.NameParent"
            .Cells("Controls.Impuls").FormulaU = "BOUND(Width*0.5,0,FALSE,Width*-0.6667,Width*-0.6667,FALSE,Width*0.5,Width*0.5,FALSE,Width*1.6667,Width*1.6667)"
            .Cells("Controls.Impuls.Y").FormulaU = "BOUND(Height*0.5,0,FALSE,Height*-0.6667,Height*-0.6667,FALSE,Height*0.5,Height*0.5,FALSE,Height*1.6667,Height*1.6667)"
            .Cells("User.SAType").FormulaU = typePlanSensor
            
        End With
        'Собираем в коллецию вставленные датчики
        colSensorOnPLAN.Add ActiveWindow.Selection(1)
        vsoSelection.DeselectAll
    Next
    'Выделяем вставленные датчики
    ActiveWindow.DeselectAll
    For Each shpSensorOnPLAN In colSensorOnPLAN
        ActiveWindow.Select shpSensorOnPLAN, visSelect
    Next
    With ActiveWindow.Selection
        'Выравниваем по горизонтали
        .Align visHorzAlignNone, visVertAlignMiddle, False
        'Распределяем по горизонтали
        .Distribute visDistHorzSpace, False
        DoEvents
        'Поднимаем вверх
        .Move 0, ActivePage.PageSheet.Cells("PageHeight").Result(0) - .PrimaryItem.Cells("PinY").Result(0)
    End With

End Sub




Public Sub VynoskaPlan(Connects As IVConnects)
'------------------------------------------------------------------------------------------------------------
' Macros        : VynoskaPlan - Заполняет выноску на плане (лоток и кабели)
                'Клеим выноску на лоток, по которому проложены кабели
                'Выноска заполняется названием лотка и номерами кабелей
'------------------------------------------------------------------------------------------------------------
    Dim shpVynoska As Visio.Shape
    Dim shpTouchingShapes As Visio.Shape
    Dim vsoSelection As Visio.Selection
    Dim strProvoda As String
    Dim strLotok As String
    Dim strCablePL As String
    Dim colNum As Collection
    Dim masShape() As Visio.Shape
    Dim CabTemp As Visio.Shape
    Dim i As Integer
    Dim j As Integer
    Dim UbMas As Long
    Dim AntiScale As Double
    
    AntiScale = ActivePage.PageSheet.Cells("DrawingScale").Result(0) / ActivePage.PageSheet.Cells("PageScale").Result(0)
    
    Set colNum = New Collection
    Set shpVynoska = Connects.FromSheet
    
    
    Select Case shpVynoska.Connects.Count 'кол-во соединенных концов у выноски
        Case 0 'Нет соединенных концов - отцепили
                shpVynoska.Cells("Prop.Lotok").FormulaU = """"""
                shpVynoska.Cells("Prop.Provoda").FormulaU = """"""
        Case 1, 2 'С одной стороны
            Set vsoSelection = shpVynoska.ContainingPage.SpatialSearch(shpVynoska.Cells("EndX").Result(0), shpVynoska.Cells("EndY").Result(0), visSpatialTouching, 0.02 * AntiScale, 0)
            For Each shpTouchingShapes In vsoSelection
                If ShapeSATypeIs(shpTouchingShapes, typeCablePL) Then
                    strCablePL = shpTouchingShapes.Cells("User.FullName").ResultStr(0)
                    colNum.Add shpTouchingShapes
                ElseIf ShapeSATypeIs(shpTouchingShapes, typeDuctPlan) Then
                    strLotok = shpTouchingShapes.Cells("User.FullName").ResultStr(0)
                End If
            Next
        'Case 2 'С двух сторон - не обрабатываем 2-ю сторону
        
    End Select
    
    'Провода
    If colNum.Count > 0 Then
       'из коллекции передаем номера проводов в массив для сортировки
       ReDim masShape(colNum.Count - 1)
       i = 0
       For Each CabTemp In colNum
           Set masShape(i) = CabTemp
           i = i + 1
       Next
       
       ' "Сортировка вставками" номеров проводов
       '--V--Сортируем по возрастанию номеров проводов
       UbMas = UBound(masShape)
       For j = 1 To UbMas
           Set CabTemp = masShape(j)
           i = j
           While masShape(i - 1).Cells("Prop.Number").Result(0) > CabTemp.Cells("Prop.Number").Result(0) '>:возрастание, <:убывание
               Set masShape(i) = masShape(i - 1)
               i = i - 1
               If i <= 0 Then GoTo ExitWhile
           Wend
ExitWhile:    Set masShape(i) = CabTemp
       Next
       '--Х--Сортировка по возрастанию номеров проводов
       
       strProvoda = "("
        
       For i = 0 To UbMas
           strProvoda = strProvoda & masShape(i).Cells("Prop.SymName").ResultStr(0) & masShape(i).Cells("Prop.Number").Result(0) & ";"
       Next
                       
       strProvoda = Left(strProvoda, Len(strProvoda) - 1)
       If Len(strProvoda) > 1 Then
           strProvoda = strProvoda & ")"
       End If
       
    Else
        strProvoda = ""
    End If
    
    If colNum.Count > 0 And strLotok = "" Then
        shpVynoska.Cells("Prop.Lotok").FormulaU = """" & strCablePL & """"
    ElseIf strLotok <> "" Then
        shpVynoska.Cells("Prop.Lotok").FormulaU = """" & strLotok & """"
    Else
        shpVynoska.Cells("Prop.Lotok").FormulaU = """"""
    End If
    
    shpVynoska.Cells("Prop.Provoda").FormulaU = """" & strProvoda & """"
    
End Sub
 
Function CableLength(shpCable As Shape) As Double
'------------------------------------------------------------------------------------------------------------
' Function      : CableLength - Вычисление длины ломанной линии в метрах
' Author        : Surrogate
' Date          : 2012.10.15
' Description   : Вычисление длины ломанной линии в метрах
' Link          : https://visio.getbb.ru/viewtopic.php?f=15&t=209&st=0&sk=t&sd=a
'------------------------------------------------------------------------------------------------------------
    Dim Summa As Double ' сумма длин
    Dim dX As Double ' определяем разности координат между концами отрезка
    Dim dY As Double ' определяем разности координат между концами отрезка
    Dim nRows As Integer  ' счетчик количества изломов линии
    Dim i As Integer
    
    nRows = shpCable.RowCount(visSectionFirstComponent) - 1
    Summa = 0
    For i = 1 To nRows - 1  ' пошагово перебираются узлы линии и вычисляются расстояния между узлами:
        dX = (shpCable.CellsSRC(visSectionFirstComponent, i, 0) - shpCable.CellsSRC(visSectionFirstComponent, i + 1, 0)) * 25.4 ' по оси X
        dY = (shpCable.CellsSRC(visSectionFirstComponent, i, 1) - shpCable.CellsSRC(visSectionFirstComponent, i + 1, 1)) * 25.4 ' по оси Y
        Summa = Summa + Sqr(dX ^ 2 + dY ^ 2) ' Вычисляем длину текущего отрезка и прибавляем к сумме длин предыдущих отрезков
    Next
    CableLength = Round(Summa * 0.001, 3)
    
End Function

Sub FillNazvanieShemyInBox(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : FillNazvanieShemyInBox - Заполняет Prop.SA_NazvanieShemy в шейпе шкафа/коробки на плане
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim PageName As String
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            vsoShape.Cells("Prop.SA_NazvanieShemy.Format").Formula = """" & vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").ResultStr(0) & """"
            Exit Sub
        End If
    Next
End Sub

Sub SetGofra(vsoObject As Object)
'Делает из линии гофру
    Dim mastshp As Visio.Shape
    Dim arrRowValue()
    Dim arrRowName()
    Dim arrMast()
    Dim SectionNumber As Long
    Dim RowNumber As Long

SectionNumber = visSectionUser 'User 242
sSectionName = "User."
            arrRowName = Array("Dropped", "SAType", "Name", "AdrSource", "FullName", "KodProizvoditelyaDB", "KodPoziciiDB")
            arrRowValue = Array("0|""""", _
                            "90|", _
                            "IF(Prop.HideNumber,"""",Prop.Number)&IF(Prop.HideName,"""","": ""&Prop.SymName)|", _
                            "0|""""", _
                            "Prop.FullName&"" ""&Prop.Ac3|""""", _
                            "0|""""", _
                            "0|""Код позиции/Код производителя/Код единицы""")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

SectionNumber = visSectionProp 'Prop 243
            arrRowName = Array("SymName", "Number", "AutoNum", "HideName", "HideNumber", "FullName", "Ac3", "Dlina", "NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
            arrRowValue = Array("""Название""|""Название""|0|""""|""|""10""|FALSE|FALSE|1049|0", _
                            """Номер кабеля""|""Номер кабеля""|2|""""||""20""|FALSE|FALSE|1049|0", _
                            """Автонумерация""|""Автонумерация""|3|""""|FALSE|""50""|TRUE|FALSE|1049|0", _
                            """Скрыть название""|""Скрыть название провода""|3|""""|TRUE|""30""|TRUE|FALSE|1049|0", _
                            """Скрыть номер""|""Скрыть номер провода""|3|""""|TRUE|""40""|TRUE|FALSE|1049|0", _
                            """Оболочка""|""Оболочка""|1|""Гофра;Металлорукав""|INDEX(0,Prop.FullName.Format)|""30""|FALSE|FALSE|1049|0", _
                            """Тип""|""Тип""|1|IF(LOOKUP(Prop.FullName,Prop.FullName.Format),""Р3-ЦХ-15;Р3-ЦХ-18;Р3-ЦХ-22;Р3-ЦХ-32"",""d16;d20;d25;d32"")|INDEX(0,Prop.Ac3.Format)|""40""|||1049|", _
                            """Длина кабеля, м.""|""Длина кабеля, м.""|2|""""||""50""|TRUE|FALSE|1049|0", _
                            """Название из БД""|""Название из БД""|0|""""|""""|""60""|FALSE|FALSE|1049|0", _
                            """Артикул из БД""|""Код заказа из БД""|0|""""|""""|""61""|FALSE|FALSE|1049|0", _
                            """Производитель из БД""|""Производитель из БД""|0|""""|""""|""62""|FALSE|FALSE|1049|0", _
                            """Цена из БД""|""Цена из БД""|0|""""|""""|""63""|FALSE|FALSE|1049|0", _
                            """Единица из БД""|""Единица измерения из БД""|0|""""|""""|""64""|FALSE|FALSE|1049|0")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

SectionNumber = visSectionObject
RowNumber = visRowLine 'Line Format
            arrRowName = Array("")
                    arrRowValue = Array("1|0.2 mm|0|0|0|0|0%|1|1|0 mm")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

End Sub

Sub SetRoute(vsoObject As Object)
'Делает из линии маршрут
    Dim mastshp As Visio.Shape
    Dim arrRowValue()
    Dim arrRowName()
    Dim arrMast()
    Dim SectionNumber As Long
    Dim RowNumber As Long
    
SectionNumber = visSectionUser 'User 242
sSectionName = "User."
            arrRowName = Array("BeginX", "BeginY", "EndX", "EndY")
            arrRowValue = Array("0|""""", _
                            "0|""""", _
                            "0|""""", _
                            "0|""""")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

SectionNumber = visSectionProp 'Prop 243
            arrRowName = Array("Begin", "End", "Dlina")
            arrRowValue = Array("""""|""""|2|""""|0|""""|FALSE|FALSE|1049|0", _
                            """""|""""|2|""""|0|""""|FALSE|FALSE|1049|0", _
                            """""|""""|2|""""|0|""""|FALSE|FALSE|1049|0")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

vsoObject.Characters.AddCustomFieldU "Prop.Begin&""-""&Prop.End", visFmtNumGenNoUnits


End Sub