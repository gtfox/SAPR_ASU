
'------------------------------------------------------------------------------------------------------------
' Module        : KabeliPLAN - Кабели на планах
' Author        : gtfox
' Date          : 2020.10.09/2022.02.21
' Description   : Автопрокладка кабелей по лоткам, подсчет длины, выноски на плане
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Dim selLines As Visio.Selection
Dim vsoShape As Visio.Shape
Dim vsoLayer As Visio.Layer
Dim colShapePoints As Collection
Dim PointNumber As Integer


Public Sub RouteCable() '(shpSensorFSA As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : RouteCable - Прокладывает кабель по ближайшему лотку
                'Определяет ближайший лоток и прокладывает кабель до шкафа
'------------------------------------------------------------------------------------------------------------
    
    Dim shpKabel As Visio.Shape
    Dim shpKabelPL As Visio.Shape
    Dim shpKabelPLPattern As Visio.Shape
    
    Dim colWires As Collection
    Dim colWiresIO As Collection
    Dim colCables As Collection
    Dim colCablesTemp As Collection
    Dim vsoMaster As Visio.Master
    
    Dim shpLotok As Visio.Shape
    Dim shpLotokTemp As Visio.Shape
    Dim shpSensor As Visio.Shape
    'Dim shpSensorFSA As Visio.Shape
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
    'Dim selLines As Visio.Selection
    
    Dim colLine As Collection
    Dim colLotok As Collection
    Dim colLineShort As Collection
    
    
    Dim vsoShapeLayer As Visio.Layer
    
    Dim SensorFSAPinX As Double
    Dim SensorFSAPinY As Double
    Dim dXSensorFSAPinX As Double
    Dim dYSensorFSAPinY As Double
'    Dim BoxX As Double
'    Dim BoxY As Double
'    Dim LineX As Double
'    Dim LineY As Double
    Dim PageWidth As Double
    Dim PageHeight As Double
    Dim AntiScale As Double
    
    Dim DlinaKabelya As Double
    Dim nCount As Double
    
    Dim BoxNumber As Integer 'Номер шкафа к которому подключен кабель/датчик
    Dim NazvanieShemy As String 'Название схемы шкафа к которому подключен кабель/датчик
    Dim i As Integer
    Dim n As Integer
    Dim MultiCable As Boolean
    
    Dim shpSensorFSA As Visio.Shape
    Dim colSensorFSA As Collection

    Dim LastPointNumber As Integer
    Dim clsShpPnt As classShapePoint

    Dim colShkafov As Collection
    
    AntiScale = ActivePage.PageSheet.Cells("DrawingScale").Result(0) / ActivePage.PageSheet.Cells("PageScale").Result(0)
    
    Set colLine = New Collection
    Set colLotok = New Collection
    Set colLineShort = New Collection
    Set vsoCollection = New Collection
    Set colCables = New Collection
    Set colCablesTemp = New Collection
    

    
    
    Set selSelection = ActiveWindow.Selection
    
    Set colShkafov = New Collection
    Set colSensorFSA = New Collection
    Set colShapePoints = New Collection

    'Находим шкафы и точки их подключения
    For Each shpLotok In ActivePage.Shapes
        If ShapeSATypeIs(shpLotok, typeDuctPlan) Then
            For i = 1 To shpLotok.Connects.Count 'Перебираем подключенные концы лотка
                If ShapeSATypeIs(shpLotok.Connects(i).ToSheet, typeBox) Then 'Выбираем только шкафы
                    Set clsShpPnt = New classShapePoint
                    clsShpPnt.PointNumber = colShapePoints.Count + 1
                    Select Case shpLotok.Connects(i).FromPart
                        Case visBegin
                            clsShpPnt.X = Round(shpLotok.Cells("BeginX").Result(0), 14)
                            clsShpPnt.Y = Round(shpLotok.Cells("BeginY").Result(0), 14)
                        Case visEnd
                            clsShpPnt.X = Round(shpLotok.Cells("EndX").Result(0), 14)
                            clsShpPnt.Y = Round(shpLotok.Cells("EndY").Result(0), 14)
                    End Select
                    Set clsShpPnt.ShapeOnFSA = shpLotok.Connects(i).ToSheet
                    colShapePoints.Add clsShpPnt, CStr(clsShpPnt.PointNumber)
                End If
            Next
        End If
    Next

    'Находим датчики и точки их подключения
    For Each shpSensorFSA In ActivePage.Shapes
        If ShapeSATypeIs(shpSensorFSA, typeFSASensor) Then
            Set clsShpPnt = New classShapePoint
            clsShpPnt.PointNumber = colShapePoints.Count + 1
            clsShpPnt.X = Round(shpSensorFSA.Cells("PinX").Result(0), 14)
            clsShpPnt.Y = Round(shpSensorFSA.Cells("PinY").Result(0), 14)
            Set clsShpPnt.ShapeOnFSA = shpSensorFSA
            colShapePoints.Add clsShpPnt, CStr(clsShpPnt.PointNumber)
        End If
    Next
    
    'Для каждого датчика находим кратчайшую линию
    For Each shpSensorFSA In ActivePage.Shapes
        If ShapeSATypeIs(shpSensorFSA, typeFSASensor) Then
        
            SensorFSAPinX = shpSensorFSA.Cells("PinX").Result(0)
            SensorFSAPinY = shpSensorFSA.Cells("PinY").Result(0)
            PageWidth = shpSensorFSA.ContainingPage.PageSheet.Cells("PageWidth").Result(0)
            PageHeight = shpSensorFSA.ContainingPage.PageSheet.Cells("PageHeight").Result(0)
            
            'UndoScopeID1 = Application.BeginUndoScope("Вспомогательные построения")
            
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
                MsgBox "Нет лотков поблизости или не приклеен к ящику"
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
            Set vsoLayer = Application.ActiveWindow.Page.Layers.Add("temp") 'новый слой
            vsoLayer.CellsC(visLayerActive).FormulaU = "1" 'активируем
            Application.ActiveWindow.Page.Paste
        
             
            'Находим смещение вставленного, относительно копированного
            For Each vsoShape In ActiveWindow.Selection
                If ShapeSATypeIs(vsoShape, typeFSASensor) Then
                    dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
                    dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
                    Set shpSensorFSATemp = vsoShape
                End If
                vsoShape.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
                vsoLayer.Add vsoShape, 0 'Добавляем все на временный слой
            Next
            'и сдвигаем на место
            ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY
            
            'разбиваем
            shpSensorFSATemp.Delete 'убираем лишнее перед trim
            ActiveWindow.Selection.Trim 'разбиваем
        
            'находим ближайшие линии
            Set selLines = ActivePage.SpatialSearch(SensorFSAPinX, SensorFSAPinY, visSpatialTouching, 0.02 * AntiScale, 0)
            For Each vsoShape In selLines
                If vsoShape.LayerCount > 0 Then
                    If vsoShape.Layer(1).name = vsoLayer.name Then
                        colLineShort.Add vsoShape
                    End If
                End If
            Next
            
            'находим самую короткую
            Set shpShortLine = colLineShort.Item(1)
            For i = 2 To colLineShort.Count
                If colLineShort.Item(i).Cells("Width").Result(0) < shpShortLine.Cells("Width").Result(0) Then
                    Set shpShortLine = colLineShort.Item(i)
                End If
            Next
            
            'Убираем ее с временного слоя
             vsoLayer.Remove shpShortLine, 0
             
            'Чистим вспомогательную графику
            shpLineUp.Delete
            shpLineDown.Delete
            shpLineLeft.Delete
            shpLineRight.Delete
            vsoLayer.Delete True
            
            'Создаем свойства для гофры в линии (тип как у лотка typeDuctPlan = 170)
            SetGofra shpShortLine
        End If
    Next
    
    'Берем все лотки и кратчайшие линии
    ActiveWindow.DeselectAll
    Set selSelection = ActiveWindow.Selection
    
    For Each shpLotok In ActivePage.Shapes
        If ShapeSATypeIs(shpLotok, typeDuctPlan) Then
            selSelection.Select shpLotok, visSelect
        End If
    Next
    
    'Добавляем туда 1 датчик для поиска смещения
    For Each vsoShape In ActivePage.Shapes
        If ShapeSATypeIs(vsoShape, typeFSASensor) Then
            SensorFSAPinX = vsoShape.Cells("PinX").Result(0)
            SensorFSAPinY = vsoShape.Cells("PinY").Result(0)
            Set shpSensorFSA = vsoShape
            Exit For
        End If
    Next
    selSelection.Select shpSensorFSA, visSelect

    'Копируем и вставляем на временном слое
    selSelection.Copy
    Set vsoLayer = Application.ActiveWindow.Page.Layers.Add("temp") 'новый слой
    vsoLayer.CellsC(visLayerActive).FormulaU = "1" 'активируем
    Application.ActiveWindow.Page.Paste
     
    'Находим смещение вставленного, относительно копированного
    For Each vsoShape In ActiveWindow.Selection
        If ShapeSATypeIs(vsoShape, typeFSASensor) Then
            dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
            dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
            Set shpSensorFSATemp = vsoShape
        End If
        vsoShape.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
        vsoLayer.Add vsoShape, 0 'Добавляем все на временный слой
    Next
    'и сдвигаем на место
    ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY
    
    'разбиваем
    shpSensorFSATemp.Delete 'убираем лишнее перед trim
    ActiveWindow.Selection.Trim 'разбиваем
    
    'Создаем из линий маршруты
    Set selSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer.name)
    For Each shpLine In selSelection
        SetRoute shpLine
    Next

    LastPointNumber = colShapePoints.Count
    
    PointNumber = 1

    'Сканируем маршруты в этой точке (максимум 4(4 стороны света))находим ближайшие линии
    Set selLines = ActivePage.SpatialSearch(colShapePoints(1).X, colShapePoints(1).Y, visSpatialTouching, 0.02 * AntiScale, 0)
    For Each vsoShape In selLines
        If vsoShape.LayerCount > 0 Then
            If vsoShape.Layer(1).name = vsoLayer.name Then
                'Именуем начало
                If vsoShape.OneD Then '1-D фигура
                    'Находим точки начала и конца линии в 1D фигуре
                    BeginX = Round(vsoShape.Cells("BeginX").Result(0), 14)
                    BeginY = Round(vsoShape.Cells("BeginY").Result(0), 14)
                    EndX = Round(vsoShape.Cells("EndX").Result(0), 14)
                    EndY = Round(vsoShape.Cells("EndY").Result(0), 14)
                Else '2-D фигура
                    'Находим точки начала и конца линии в 2D фигуре
                    BeginX = Round(vsoShape.Cells("PinX").Result(0) - vsoShape.Cells("Width").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 0).Result(0), 14)
                    BeginY = Round(vsoShape.Cells("PinY").Result(0) - vsoShape.Cells("Height").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 1).Result(0), 14)
                    EndX = Round(vsoShape.Cells("PinX").Result(0) - vsoShape.Cells("Width").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowLast, 0).Result(0), 14)
                    EndY = Round(vsoShape.Cells("PinY").Result(0) - vsoShape.Cells("Height").Result(0) * 0.5 + vsoShape.CellsSRC(visSectionFirstComponent, visRowLast, 1).Result(0), 14)
                End If
                
                'Именуем конец в этой точке
                If BeginX = colShapePoints(1).X And BeginY = colShapePoints(1).Y Then
                    vsoShape.Cells("Prop.Begin").Formula = PointNumber
                ElseIf EndX = colShapePoints(1).X And EndY = colShapePoints(1).Y Then
                    vsoShape.Cells("Prop.End").Formula = PointNumber
                End If
                
                'Заполняем свойство длина
                vsoShape.Cells("Prop.Dlina").Formula = CableLength(vsoShape)
                
                'Заполняет пути именами точек
                FillRoute vsoShape
            End If
        End If
    Next


'----------------------------------------------------------------------------
Exit Sub
















    'Dim UndoScopeID1 As Long


    Set shpSensor = ShapeByHyperLink(shpSensorFSA.Cells("Hyperlink.Shema.SubAddress").ResultStr(0))
    If Not shpSensor Is Nothing Then
        MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)
    Else
        MsgBox "Датчик не связан"
        Exit Sub
    End If
    
    'Находим кабели на плане (чтобы не проложить повторно)
    For Each shpKabel In shpSensorFSA.ContainingPage.Shapes 'Перебираем все кабели
        If ShapeSATypeIs(shpKabel, typeCablePL) Then
            colCablesTemp.Add shpKabel, CStr(shpKabel.Cells("Prop.Number").ResultStr(0))
        End If
    Next
    
    'Находим кабель/кабели подключенные к датчику исключая существующие(уже проложенные)
    For Each vsoShape In shpSensor.Shapes 'Перебираем все входы датчика
        If ShapeSATypeIs(vsoShape, typeSensorIO) Then
            'Находим подключенные провода
            Set vsoCollection = FillColWires(vsoShape)
            nCount = colCablesTemp.Count
            On Error Resume Next
            colCablesTemp.Add vsoCollection.Item(1).Parent, IIf(vsoCollection.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), vsoCollection.Item(1).Parent.Cells("Prop.SymName").ResultStr(0) & vsoCollection.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(vsoCollection.Item(1).Parent.Cells("Prop.Number").ResultStr(0)))
            If colCablesTemp.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                colCables.Add vsoCollection.Item(1).Parent
                nCount = colCablesTemp.Count
            End If
        End If
    Next
    If colCables.Count = 0 Then Exit Sub 'MsgBox "Не найдены кабели", vbExclamation + vbOKOnly, "Info": Exit Sub
    'Шкаф к которому подключен кабель (Предполагается что 1 датчик подключен к 1 шкафу (даже многокабельный)
'    BoxNumber = colCables.Item(1).Cells("User.LinkToBox").Result(0)
'    NazvanieShemy = colCables.Item(1).ContainingPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0)
    NazvanieShemy = colCables.Item(1).Cells("User.LinkToBox").ResultStr(0)

    
    
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
     
    'Находим лоток идущий в наш шкаф и которого касается кратчайшая линия
    Set selSelection = shpShortLine.SpatialNeighbors(visSpatialTouching, 0.02, 0)
    For Each vsoShape In selSelection 'Шейпы в выделении
        If ShapeSATypeIs(vsoShape, typeDuctPlan) Then  'Нашли лоток
'            'Находим координаты точки в которой лоток подключен к шкафу
'            For i = 1 To vsoShape.Connects.Count 'Перебираем подключенные концы лотка
'                If vsoShape.Connects(i).ToSheet.Name Like "Box*" Then 'Выбираем только шкафы
'                    If vsoShape.Connects(i).FromCell.Name Like "Begin*" Then
'                        BoxX = vsoShape.Connects(i).FromSheet.Cells("BeginX").Result(0)
'                        BoxY = vsoShape.Connects(i).FromSheet.Cells("BeginY").Result(0)
'                    ElseIf vsoShape.Connects(i).FromCell.Name Like "End*" Then
'                        BoxX = vsoShape.Connects(i).FromSheet.Cells("EndX").Result(0)
'                        BoxY = vsoShape.Connects(i).FromSheet.Cells("EndY").Result(0)
'                    End If
'                End If
'            Next
            Set shpLotok = vsoShape
            Exit For
        End If
    Next
    
'    'Находим координаты точки в которой лоток пересекается с кратчайшей линией
'    If shpShortLine.Cells("BeginX").Result(0) = SensorFSAPinX Then
'        LineX = shpShortLine.Cells("EndX").Result(0)
'        LineY = shpShortLine.Cells("EndY").Result(0)
'    ElseIf shpShortLine.Cells("EndX").Result(0) = SensorFSAPinX Then
'        LineX = shpShortLine.Cells("BeginX").Result(0)
'        LineY = shpShortLine.Cells("BeginY").Result(0)
'    End If
    
    'Выделяем лоток и кратчайшую линию до лотка
    ActiveWindow.DeselectAll
    Set selSelection = ActiveWindow.Selection
    selSelection.Select shpSensorFSA, visSelect
    selSelection.Select shpLotok, visSelect
    selSelection.Select shpShortLine, visSelect
    
    'Копируем и вставляем на временном слое
    selSelection.Copy
    Set vsoLayer = Application.ActiveWindow.Page.Layers.Add("temp") 'новый слой
    vsoLayer.CellsC(visLayerActive).FormulaU = "1" 'активируем
    Application.ActiveWindow.Page.Paste
     
    'Находим смещение вставленного, относительно копированного
    For Each vsoShape In ActiveWindow.Selection
        If ShapeSATypeIs(vsoShape, typeFSASensor) Then
            dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
            dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
            Set shpSensorFSATemp = vsoShape
        End If
        vsoShape.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
        vsoLayer.Add vsoShape, 0 'Добавляем все на временный слой
    Next
    'и сдвигаем на место
    ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY
    
    'разбиваем
    shpSensorFSATemp.Delete 'убираем лишнее перед trim
    ActiveWindow.Selection.Trim 'разбиваем
    
    'Находим ту часть разбитого лотка которая идет от кратчайшей линии до шкафа
    Set selSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer) 'выделяем все в слое
    'Для каждого делаем спатиал и ищем шкаф
    For Each vsoShape In selSelection
        Set selSelectionTemp = vsoShape.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0.02 * AntiScale, 0)
        For Each vsoShapeTemp In selSelectionTemp
            If ShapeSATypeIs(vsoShapeTemp, typeBox) Then
                If vsoShapeTemp.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then
                    Set shpLotokTemp = vsoShape
                End If
            End If
        Next
    Next
    
    'Убираем его с временного слоя
    vsoLayer.Remove shpLotokTemp, 0
    'Чистим вспомогательную графику
    vsoLayer.Delete True
     
    'Соединяем найденный кусок с кратчайшей линией
    Set vsoLayer = Application.ActiveWindow.Page.Layers.Add("temp") 'Временный слой
    vsoLayer.CellsC(visLayerActive).FormulaU = "1" 'активируем
    shpLotokTemp.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
    shpShortLine.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
    vsoLayer.Add shpLotokTemp, 0 'Добавляем на временный слой
    vsoLayer.Add shpShortLine, 0 'Добавляем на временный слой
    ActiveWindow.DeselectAll
    ActiveWindow.Select shpLotokTemp, visSelect
    ActiveWindow.Select shpShortLine, visSelect
    Application.ActiveWindow.Selection.Join 'соединяем
    Set selSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer) 'выделяем все в слое
    Set shpKabelPLPattern = selSelection.PrimaryItem 'Таки профит! Гребаный кабель случился!
    'Убираем с временного слоя
    vsoLayer.Remove shpKabelPLPattern, 0
    'Чистим вспомогательную графику
    vsoLayer.Delete True
    
    'Считаем длину кабеля
    DlinaKabelya = CableLength(shpKabelPLPattern)
    
    'Создаем свойтва шаблона кабеля на плане
    With shpKabelPLPattern
        .CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = 1 'Обычная линия
        .CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.2 mm"
        .AddSection visSectionUser
        .AddRow visSectionUser, visRowLast, visTagDefault
        .CellsSRC(visSectionUser, visRowLast, visUserValue).RowNameU = "SAType"
        .CellsSRC(visSectionUser, visRowLast, visUserValue).FormulaForceU = "90"
        .AddSection visSectionProp
        .AddRow visSectionProp, visRowLast, visTagDefault
        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "SymName"
        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Букв. обозначение"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Букв. обозначение"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "0"
        .AddRow visSectionProp, visRowLast, visTagDefault
        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "Number"
        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Номер кабеля"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Номер кабеля"""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "2"
        .AddRow visSectionProp, visRowLast, visTagDefault
        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "Dlina"
        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Длина кабеля, м."""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Длина кабеля, м."""
        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "2"
        
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
        
        'Заполняем длину кабеля на эл.схеме (длина кабеля СВП ссылается формулой на эл.сх.)
        shpKabel.Cells("Prop.Dlina").FormulaU = "Pages[" + shpKabelPL.ContainingPage.NameU + "]!" + shpKabelPL.NameID + "!Prop.Dlina"
    Next
    
    'Удаляем шаблон кабеля
    shpKabelPLPattern.Delete
    
    Application.ActiveWindow.DeselectAll

'    'Application.EndUndoScope UndoScopeID1, True

End Sub

Sub FillRoute(shpRouteToPoint As Visio.Shape)
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
        BeginX = Round(shpRouteToPoint.Cells("BeginX").Result(0), 14)
        BeginY = Round(shpRouteToPoint.Cells("BeginY").Result(0), 14)
        EndX = Round(shpRouteToPoint.Cells("EndX").Result(0), 14)
        EndY = Round(shpRouteToPoint.Cells("EndY").Result(0), 14)
    Else '2-D фигура
        'Находим точки начала и конца линии в 2D фигуре
        BeginX = Round(shpRouteToPoint.Cells("PinX").Result(0) - shpRouteToPoint.Cells("Width").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 0).Result(0), 14)
        BeginY = Round(shpRouteToPoint.Cells("PinY").Result(0) - shpRouteToPoint.Cells("Height").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 1).Result(0), 14)
        EndX = Round(shpRouteToPoint.Cells("PinX").Result(0) - shpRouteToPoint.Cells("Width").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowLast, 0).Result(0), 14)
        EndY = Round(shpRouteToPoint.Cells("PinY").Result(0) - shpRouteToPoint.Cells("Height").Result(0) * 0.5 + shpRouteToPoint.CellsSRC(visSectionFirstComponent, visRowLast, 1).Result(0), 14)
    End If
    
    If shpRouteToPoint.Cells("Prop.Begin").Result(0) = 0 Or shpRouteToPoint.Cells("Prop.End").Result(0) = 0 Then
        'Находим точку на другом конце
        PointNumber = PointNumber + 1
        clsPoint.PointNumber = PointNumber
        If shpRouteToPoint.Cells("Prop.Begin").Result(0) = 0 Then
            clsPoint.X = BeginX
            clsPoint.Y = BeginY
            shpRouteToPoint.Cells("Prop.Begin").Formula = clsPoint.PointNumber
        ElseIf shpRouteToPoint.Cells("Prop.End").Result(0) = 0 Then
            clsPoint.X = EndX
            clsPoint.Y = EndY
            shpRouteToPoint.Cells("Prop.End").Formula = clsPoint.PointNumber
        End If
    Else
        Exit Sub
    End If
    
    'Сканируем маршруты в этой точке (максимум 4(4 стороны света))находим ближайшие линии
    Set selLines = ActivePage.SpatialSearch(clsPoint.X, clsPoint.Y, visSpatialTouching, 0.02 * AntiScale, 0)

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
            BeginX = Round(shpRoute.Cells("BeginX").Result(0), 14)
            BeginY = Round(shpRoute.Cells("BeginY").Result(0), 14)
            EndX = Round(shpRoute.Cells("EndX").Result(0), 14)
            EndY = Round(shpRoute.Cells("EndY").Result(0), 14)

            'Нет именованных концов
            If shpRoute.Cells("Prop.Begin").Result(0) = 0 And shpRoute.Cells("Prop.End").Result(0) = 0 Then
    
                'Именуем конец в этой точке
                If BeginX = clsPoint.X And BeginY = clsPoint.Y Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                ElseIf EndX = clsPoint.X And EndY = clsPoint.Y Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
                End If
                
                'Заполняем свойство длина
                shpRoute.Cells("Prop.Dlina").Formula = CableLength(shpRoute)
                
            'Именован один конец
            ElseIf shpRoute.Cells("Prop.Begin").Result(0) = 0 Or shpRoute.Cells("Prop.End").Result(0) = 0 Then
                'Именован другой конец - именуем наш, другой не трогаем
                'Именуем конец в этой точке
                If BeginX = clsPoint.X And BeginY = clsPoint.Y And shpRoute.Cells("Prop.Begin").Result(0) = 0 Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                ElseIf EndX = clsPoint.X And EndY = clsPoint.Y And shpRoute.Cells("Prop.End").Result(0) = 0 Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
    
                'Именован наш конец - исключение (мы не должны попасть в точку в которой есть именованные концы)
                ElseIf BeginX = clsPoint.X And BeginY = clsPoint.Y And shpRoute.Cells("Prop.Begin").Result(0) <> 0 Then
                    MsgBox "Именованый конец в точке: " & clsPoint.PointNumber & ". Конец: " & shpRoute.Cells("Prop.Begin").Result(0) & ". Маршрут: " & shpRoute.Cells("Prop.Begin").Result(0) & " - " & shpRoute.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
                    Exit Sub
                ElseIf shpRoute.Cells("EndX").Result(0) = clsPoint.X And shpRoute.Cells("EndY").Result(0) = clsPoint.Y And shpRoute.Cells("Prop.End").Result(0) <> 0 Then
                    MsgBox "Именованый конец в точке: " & clsPoint.PointNumber & ". Конец: " & shpRoute.Cells("Prop.End").Result(0) & ". Маршрут: " & shpRoute.Cells("Prop.Begin").Result(0) & " - " & shpRoute.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
                    Exit Sub
                End If
                
            'Именованы оба конца
            ElseIf shpRoute.Cells("Prop.Begin").Result(0) <> 0 And shpRoute.Cells("Prop.End").Result(0) <> 0 Then
                'Маршрут уже обработан полностью (с двух концов)
            End If
        
        
                
        Else '2-D фигура
        
            'Находим точки начала и конца линии в 2D фигуре
            BeginX = Round(shpRoute.Cells("PinX").Result(0) - shpRoute.Cells("Width").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 0).Result(0), 14)
            BeginY = Round(shpRoute.Cells("PinY").Result(0) - shpRoute.Cells("Height").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 1).Result(0), 14)
            EndX = Round(shpRoute.Cells("PinX").Result(0) - shpRoute.Cells("Width").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowLast, 0).Result(0), 14)
            EndY = Round(shpRoute.Cells("PinY").Result(0) - shpRoute.Cells("Height").Result(0) * 0.5 + shpRoute.CellsSRC(visSectionFirstComponent, visRowLast, 1).Result(0), 14)

            'Нет именованных концов
            If shpRoute.Cells("Prop.Begin").Result(0) = 0 And shpRoute.Cells("Prop.End").Result(0) = 0 Then
    
                'Именуем конец в этой точке
                If BeginX = clsPoint.X And BeginY = clsPoint.Y Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                ElseIf EndX = clsPoint.X And EndY = clsPoint.Y Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
                End If
                
                'Заполняем свойство длина
                shpRoute.Cells("Prop.Dlina").Formula = CableLength(shpRoute)
                
            'Именован один конец
            ElseIf shpRoute.Cells("Prop.Begin").Result(0) = 0 Or shpRoute.Cells("Prop.End").Result(0) = 0 Then
                'Именован другой конец - именуем наш, другой не трогаем
                'Именуем конец в этой точке
                If BeginX = clsPoint.X And BeginY = clsPoint.Y And shpRoute.Cells("Prop.Begin").Result(0) = 0 Then
                    shpRoute.Cells("Prop.Begin").Formula = clsPoint.PointNumber
                ElseIf EndX = clsPoint.X And EndY = clsPoint.Y And shpRoute.Cells("Prop.End").Result(0) = 0 Then
                    shpRoute.Cells("Prop.End").Formula = clsPoint.PointNumber
    
                'Именован наш конец - исключение (мы не должны попасть в точку в которой есть именованные концы)
                ElseIf BeginX = clsPoint.X And BeginY = clsPoint.Y And shpRoute.Cells("Prop.Begin").Result(0) <> 0 Then
                    MsgBox "Именованый конец в точке: " & clsPoint.PointNumber & ". Конец: " & shpRoute.Cells("Prop.Begin").Result(0) & ". Маршрут: " & shpRoute.Cells("Prop.Begin").Result(0) & " - " & shpRoute.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
                    Exit Sub
                ElseIf EndX = clsPoint.X And EndY = clsPoint.Y And shpRoute.Cells("Prop.End").Result(0) <> 0 Then
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
        FillRoute shpRoute
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
        MsgBox "Нет ФСА для вставки", vbExclamation, "Название ФСА пустое"
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
        If ShapeSATypeIs(shpSensorOnPLAN, typeFSASensor) Then
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
            .Cells("Hyperlink.Shema.ExtraInfo").Formula = AdrParent + "!Hyperlink.Shema.ExtraInfo"
            .Cells("Hyperlink.FSA.ExtraInfo").Formula = AdrParent + "!Hyperlink.FSA.ExtraInfo"
            .Cells("Hyperlink.Shema.SubAddress").Formula = AdrParent + "!Hyperlink.Shema.SubAddress"
            .Cells("Hyperlink.FSA.SubAddress").Formula = AdrParent + "!Hyperlink.FSA.SubAddress"
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

'    If vsoSelection.Count = 0 Then
'        Exit Sub
'    End If
'    Set vsoGroup = vsoSelection.Group
'    vsoSelection.DeselectAll
'    vsoSelection.Select vsoGroup, visSelect
'    For Each shpSensorOnPLAN In vsoSelection
'       shpSensorOnPLAN.CellsU("EventDrop").FormulaU = """"""
'       shpSensorOnPLAN.CellsU("EventMultiDrop").FormulaU = """"""
'    Next
'    'Копируем на план что насобирали
'    vsoSelection.Copy
'    For Each shpSensorOnPLAN In vsoSelection
'       shpSensorOnPLAN.CellsU("EventDrop").FormulaU = "CALLTHIS(""ThisDocument.EventDropAutoNum"")"
'       shpSensorOnPLAN.CellsU("EventMultiDrop").FormulaU = "CALLTHIS(""AutoNumber.AutoNumFSA"")"
'    Next
'    vsoSelection.Ungroup
'    'Отключаем события автоматизации (чтобы не перенумеровалось все)
'    DoEvents
'    Application.EventsEnabled = 0
'    'Вставляем на листе план
'    ActiveWindow.Page = ActiveDocument.Pages(vsoPagePlan.Name)
'    ActivePage.Paste
'    ActiveWindow.Selection.Ungroup
'    'Включаем пункт меню "Проложить кабель"
'    For Each shpSensorOnPLAN In ActiveWindow.Selection
'       shpSensorOnPLAN.Cells("Actions.Kabel.Invisible").Formula = 0
'       shpSensorOnPLAN.Cells("Actions.AddReference.Invisible").Formula = 1
'       shpSensorOnPLAN.Cells("Prop.KanalNumber").Formula = 0
''       shpSensorOnPLAN.Cells("User.Dropped").Formula = 0
'    Next
'    With ActiveWindow.Selection
'        'Выравниваем по горизонтали
'        .Align visHorzAlignNone, visVertAlignMiddle, False
'        'Распределяем по горизонтали
'        .Distribute visDistHorzSpace, False
'        DoEvents
'        'Поднимаем вверх
'        .Move 0, ActivePage.PageSheet.Cells("PageHeight").Result(0) - .PrimaryItem.Cells("PinY").Result(0)
'
'    End With
'    'Включаем события автоматизации
'    Application.EventsEnabled = -1
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
    
    If colNum.Count > 0 And strLotok = "" Then strLotok = "Гофра d16"
    
    shpVynoska.Cells("Prop.Lotok").FormulaU = """" & strLotok & """"
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
                            "170|", _
                            "IF(Prop.HideNumber,"""",Prop.Number)&IF(Prop.HideName,"""","": ""&Prop.SymName)|", _
                            "0|""""", _
                            "Prop.SymName&"" ""&Prop.Ac3|""""", _
                            "0|""""", _
                            "0|""Код позиции/Код производителя/Код единицы""")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

SectionNumber = visSectionProp 'Prop 243
            arrRowName = Array("SymName", "Number", "AutoNum", "HideName", "HideNumber", "FullName", "Ac3", "Dlina", "NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
            arrRowValue = Array("""Название""|""Название""|0||""Гофра""|""10""|FALSE|FALSE|1049|0", _
                            """Номер провода""|""Номер провода""|2|""""||""20""|TRUE|FALSE|1033|0", _
                            """Автонумерация""|""Автонумерация""|3|""""|FALSE|""50""|TRUE|FALSE|1033|0", _
                            """Скрыть название""|""Скрыть название провода""|3|""""|TRUE|""30""|TRUE|FALSE|1033|0", _
                            """Скрыть номер""|""Скрыть номер провода""|3|""""|TRUE|""40""|TRUE|FALSE|1033|0", _
                            """Имя лотка""|""Имя лотка""|0|""""|User.FullName|""91""|TRUE|FALSE|1033|0", _
                            """Сечение""|""Сечение""|1|""d16;d20;d25;d32""|INDEX(0,Prop.Ac3.Format)|""40""|||1033|", _
                            """Длина кабеля, м.""|""Длина кабеля, м.""|2|""""||""50""|TRUE|FALSE|1033|0", _
                            """Название из БД""|""Название из БД""|0|""""|""""|""60""|FALSE|FALSE|1033|0", _
                            """Артикул из БД""|""Код заказа из БД""|0|""""|""""|""61""|FALSE|FALSE|1033|0", _
                            """Производитель из БД""|""Производитель из БД""|0|""""|""""|""62""|FALSE|FALSE|1033|0", _
                            """Цена из БД""|""Цена из БД""|0|""""|""""|""63""|FALSE|FALSE|1033|0", _
                            """Единица из БД""|""Единица измерения из БД""|0|""""|""""|""64""|FALSE|FALSE|1033|0")
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

SectionNumber = visSectionProp 'Prop 243
            arrRowName = Array("Begin", "End", "Dlina")
            arrRowValue = Array("""""|""""|2|""""|0|""""|FALSE|FALSE|1049|0", _
                            """""|""""|2|""""|0|""""|FALSE|FALSE|1049|0", _
                            """""|""""|2|""""|0|""""|FALSE|FALSE|1049|0")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

vsoObject.Characters.AddCustomFieldU "Prop.Begin&""-""&Prop.End", visFmtNumGenNoUnits


End Sub


'Public Sub RouteCable_old(shpSensorFSA As Visio.Shape)
''------------------------------------------------------------------------------------------------------------
'' Macros        : RouteCable - Прокладывает кабель по ближайшему лотку
'                'Определяет ближайший лоток и прокладывает кабель до шкафа
'                '
'                'Устарел, т.к. используется чисто графический способ, который сильно ограничен
''------------------------------------------------------------------------------------------------------------
'
'    Dim shpKabel As Visio.Shape
'    Dim shpKabelPL As Visio.Shape
'    Dim shpKabelPLPattern As Visio.Shape
'
'    Dim colWires As Collection
'    Dim colWiresIO As Collection
'    Dim colCables As Collection
'    Dim colCablesTemp As Collection
'    Dim vsoMaster As Visio.Master
'
'    Dim shpLotok As Visio.Shape
'    Dim shpLotokTemp As Visio.Shape
'    Dim shpSensor As Visio.Shape
'    'Dim shpSensorFSA As Visio.Shape
'    Dim shpSensorFSATemp As Visio.Shape
'    Dim shpShortLine As Visio.Shape
'    Dim vsoShape As Visio.Shape
'    Dim vsoShapeTemp As Visio.Shape
'    Dim vsoCollection As Collection
'
'    Dim shpLineUp As Visio.Shape
'    Dim shpLineDown As Visio.Shape
'    Dim shpLineLeft As Visio.Shape
'    Dim shpLineRight As Visio.Shape
'    Dim selLineUp As Visio.Selection
'    Dim selLineDown As Visio.Selection
'    Dim selLineLeft As Visio.Selection
'    Dim selLineRight As Visio.Selection
'    Dim selSelection As Visio.Selection
'    Dim selSelectionTemp As Visio.Selection
'    Dim selLines As Visio.Selection
'
'    Dim colLine As Collection
'    Dim colLotok As Collection
'    Dim colLineShort As Collection
'
'    Dim vsoLayer As Visio.Layer
'    Dim vsoShapeLayer As Visio.Layer
'
'    Dim SensorFSAPinX As Double
'    Dim SensorFSAPinY As Double
'    Dim dXSensorFSAPinX As Double
'    Dim dYSensorFSAPinY As Double
''    Dim BoxX As Double
''    Dim BoxY As Double
''    Dim LineX As Double
''    Dim LineY As Double
'    Dim PageWidth As Double
'    Dim PageHeight As Double
'    Dim AntiScale As Double
'
'    Dim DlinaKabelya As Double
'    Dim nCount As Double
'
'    Dim BoxNumber As Integer 'Номер шкафа к которому подключен кабель/датчик
'    Dim NazvanieShemy As String 'Название схемы шкафа к которому подключен кабель/датчик
'    Dim i As Integer
'    Dim n As Integer
'    Dim MultiCable As Boolean
'
'    'Dim UndoScopeID1 As Long
'
'    AntiScale = ActivePage.PageSheet.Cells("DrawingScale").Result(0) / ActivePage.PageSheet.Cells("PageScale").Result(0)
'
'    Set colLine = New Collection
'    Set colLotok = New Collection
'    Set colLineShort = New Collection
'    Set vsoCollection = New Collection
'    Set colCables = New Collection
'    Set colCablesTemp = New Collection
'
'    Set selSelection = ActiveWindow.Selection
'    Set shpSensor = ShapeByHyperLink(shpSensorFSA.Cells("Hyperlink.Shema.SubAddress").ResultStr(0))
'    If Not shpSensor Is Nothing Then
'        MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)
'    Else
'        MsgBox "Датчик не связан"
'        Exit Sub
'    End If
'
'    'Находим кабели на плане (чтобы не проложить повторно)
'    For Each shpKabel In shpSensorFSA.ContainingPage.Shapes 'Перебираем все кабели
'        If ShapeSATypeIs(shpKabel, typeCablePL) Then
'            colCablesTemp.Add shpKabel, CStr(shpKabel.Cells("Prop.Number").ResultStr(0))
'        End If
'    Next
'
'    'Находим кабель/кабели подключенные к датчику исключая существующие(уже проложенные)
'    For Each vsoShape In shpSensor.Shapes 'Перебираем все входы датчика
'        If ShapeSATypeIs(vsoShape, typeSensorIO) Then
'            'Находим подключенные провода
'            Set vsoCollection = FillColWires(vsoShape)
'            nCount = colCablesTemp.Count
'            On Error Resume Next
'            colCablesTemp.Add vsoCollection.Item(1).Parent, IIf(vsoCollection.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), vsoCollection.Item(1).Parent.Cells("Prop.SymName").ResultStr(0) & vsoCollection.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(vsoCollection.Item(1).Parent.Cells("Prop.Number").ResultStr(0)))
'            If colCablesTemp.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
'                colCables.Add vsoCollection.Item(1).Parent
'                nCount = colCablesTemp.Count
'            End If
'        End If
'    Next
'    If colCables.Count = 0 Then Exit Sub 'MsgBox "Не найдены кабели", vbExclamation + vbOKOnly, "Info": Exit Sub
'    'Шкаф к которому подключен кабель (Предполагается что 1 датчик подключен к 1 шкафу (даже многокабельный)
''    BoxNumber = colCables.Item(1).Cells("User.LinkToBox").Result(0)
''    NazvanieShemy = colCables.Item(1).ContainingPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0)
'    NazvanieShemy = colCables.Item(1).Cells("User.LinkToBox").ResultStr(0)
'
'    SensorFSAPinX = shpSensorFSA.Cells("PinX").Result(0)
'    SensorFSAPinY = shpSensorFSA.Cells("PinY").Result(0)
'    PageWidth = shpSensorFSA.ContainingPage.PageSheet.Cells("PageWidth").Result(0)
'    PageHeight = shpSensorFSA.ContainingPage.PageSheet.Cells("PageHeight").Result(0)
'
'    'UndoScopeID1 = Application.BeginUndoScope("Вспомогательные построения")
'
'    'Рисуем линии во все стороны
'    Set shpLineUp = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, SensorFSAPinX, PageHeight)
'    Set shpLineDown = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, SensorFSAPinX, 0)
'    Set shpLineLeft = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, 0, SensorFSAPinY)
'    Set shpLineRight = ActivePage.DrawLine(SensorFSAPinX, SensorFSAPinY, PageWidth, SensorFSAPinY)
'
'    'Находим все пересечения
'    Set selLineUp = shpLineUp.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
'    Set selLineDown = shpLineDown.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
'    Set selLineLeft = shpLineLeft.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
'    Set selLineRight = shpLineRight.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
'
'    'Выбираем лотки и линии
'    AddLotokToCol shpLineUp, selLineUp, colLine, colLotok, NazvanieShemy 'BoxNumber
'    AddLotokToCol shpLineDown, selLineDown, colLine, colLotok, NazvanieShemy 'BoxNumber
'    AddLotokToCol shpLineLeft, selLineLeft, colLine, colLotok, NazvanieShemy 'BoxNumber
'    AddLotokToCol shpLineRight, selLineRight, colLine, colLotok, NazvanieShemy 'BoxNumber
'    If colLotok.Count = 0 Then 'нет лотков - выходим
'        'Чистим вспомогательную графику
'        shpLineUp.Delete
'        shpLineDown.Delete
'        shpLineLeft.Delete
'        shpLineRight.Delete
'        MsgBox "Нет лотков поблизости или не приклеен к ящику"
'        Exit Sub
'    End If
'
'    'Выделяем их
'    selSelection.Select shpSensorFSA, visSelect
'    For Each vsoShape In colLine
'        selSelection.Select vsoShape, visSelect
'    Next
'    For Each vsoShape In colLotok
'        selSelection.Select vsoShape, visSelect
'    Next
'
'    'Копируем и вставляем на временном слое
'    selSelection.Copy
'    Set vsoLayer = Application.ActiveWindow.Page.Layers.Add("temp") 'новый слой
'    vsoLayer.CellsC(visLayerActive).FormulaU = "1" 'активируем
'    Application.ActiveWindow.Page.Paste
'
'
'    'Находим смещение вставленного, относительно копированного
'    For Each vsoShape In ActiveWindow.Selection
'        If ShapeSATypeIs(vsoShape, typeFSASensor) Then
'            dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
'            dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
'            Set shpSensorFSATemp = vsoShape
'        End If
'        vsoShape.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
'        vsoLayer.Add vsoShape, 0 'Добавляем все на временный слой
'    Next
'    'и сдвигаем на место
'    ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY
'
'    'разбиваем
'    shpSensorFSATemp.Delete 'убираем лишнее перед trim
'    ActiveWindow.Selection.Trim 'разбиваем
'
'    'находим ближайшие линии
'    Set selLines = ActivePage.SpatialSearch(SensorFSAPinX, SensorFSAPinY, visSpatialTouching, 0.02 * AntiScale, 0)
'    For Each vsoShape In selLines
'        If vsoShape.LayerCount > 0 Then
'            If vsoShape.Layer(1).name = vsoLayer.name Then
'                colLineShort.Add vsoShape
'            End If
'        End If
'    Next
'
'    'находим самую короткую
'    Set shpShortLine = colLineShort.Item(1)
'    For i = 2 To colLineShort.Count
'        If colLineShort.Item(i).Cells("Width").Result(0) < shpShortLine.Cells("Width").Result(0) Then
'            Set shpShortLine = colLineShort.Item(i)
'        End If
'    Next
'
'    'Убираем ее с временного слоя
'     vsoLayer.Remove shpShortLine, 0
'
'    'Чистим вспомогательную графику
'    shpLineUp.Delete
'    shpLineDown.Delete
'    shpLineLeft.Delete
'    shpLineRight.Delete
'    vsoLayer.Delete True
'
'    'Находим лоток идущий в наш шкаф и которого касается кратчайшая линия
'    Set selSelection = shpShortLine.SpatialNeighbors(visSpatialTouching, 0.02, 0)
'    For Each vsoShape In selSelection 'Шейпы в выделении
'        If ShapeSATypeIs(vsoShape, typeDuctPlan) And (PointLotokToBox(vsoShape, NazvanieShemy)) Then 'Нашли лоток
''            'Находим координаты точки в которой лоток подключен к шкафу
''            For i = 1 To vsoShape.Connects.Count 'Перебираем подключенные концы лотка
''                If vsoShape.Connects(i).ToSheet.Name Like "Box*" Then 'Выбираем только шкафы
''                    If vsoShape.Connects(i).FromCell.Name Like "Begin*" Then
''                        BoxX = vsoShape.Connects(i).FromSheet.Cells("BeginX").Result(0)
''                        BoxY = vsoShape.Connects(i).FromSheet.Cells("BeginY").Result(0)
''                    ElseIf vsoShape.Connects(i).FromCell.Name Like "End*" Then
''                        BoxX = vsoShape.Connects(i).FromSheet.Cells("EndX").Result(0)
''                        BoxY = vsoShape.Connects(i).FromSheet.Cells("EndY").Result(0)
''                    End If
''                End If
''            Next
'            Set shpLotok = vsoShape
'            Exit For
'        End If
'    Next
'
''    'Находим координаты точки в которой лоток пересекается с кратчайшей линией
''    If shpShortLine.Cells("BeginX").Result(0) = SensorFSAPinX Then
''        LineX = shpShortLine.Cells("EndX").Result(0)
''        LineY = shpShortLine.Cells("EndY").Result(0)
''    ElseIf shpShortLine.Cells("EndX").Result(0) = SensorFSAPinX Then
''        LineX = shpShortLine.Cells("BeginX").Result(0)
''        LineY = shpShortLine.Cells("BeginY").Result(0)
''    End If
'
'    'Выделяем лоток и кратчайшую линию до лотка
'    ActiveWindow.DeselectAll
'    Set selSelection = ActiveWindow.Selection
'    selSelection.Select shpSensorFSA, visSelect
'    selSelection.Select shpLotok, visSelect
'    selSelection.Select shpShortLine, visSelect
'
'    'Копируем и вставляем на временном слое
'    selSelection.Copy
'    Set vsoLayer = Application.ActiveWindow.Page.Layers.Add("temp") 'новый слой
'    vsoLayer.CellsC(visLayerActive).FormulaU = "1" 'активируем
'    Application.ActiveWindow.Page.Paste
'
'    'Находим смещение вставленного, относительно копированного
'    For Each vsoShape In ActiveWindow.Selection
'        If ShapeSATypeIs(vsoShape, typeFSASensor) Then
'            dXSensorFSAPinX = SensorFSAPinX - vsoShape.Cells("PinX").Result(0)
'            dYSensorFSAPinY = SensorFSAPinY - vsoShape.Cells("PinY").Result(0)
'            Set shpSensorFSATemp = vsoShape
'        End If
'        vsoShape.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
'        vsoLayer.Add vsoShape, 0 'Добавляем все на временный слой
'    Next
'    'и сдвигаем на место
'    ActiveWindow.Selection.Move dXSensorFSAPinX, dYSensorFSAPinY
'
'    'разбиваем
'    shpSensorFSATemp.Delete 'убираем лишнее перед trim
'    ActiveWindow.Selection.Trim 'разбиваем
'
'    'Находим ту часть разбитого лотка которая идет от кратчайшей линии до шкафа
'    Set selSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer) 'выделяем все в слое
'    'Для каждого делаем спатиал и ищем шкаф
'    For Each vsoShape In selSelection
'        Set selSelectionTemp = vsoShape.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0.02 * AntiScale, 0)
'        For Each vsoShapeTemp In selSelectionTemp
'            If ShapeSATypeIs(vsoShapeTemp, typeBox) Then
'                If vsoShapeTemp.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then
'                    Set shpLotokTemp = vsoShape
'                End If
'            End If
'        Next
'    Next
'
'    'Убираем его с временного слоя
'    vsoLayer.Remove shpLotokTemp, 0
'    'Чистим вспомогательную графику
'    vsoLayer.Delete True
'
'    'Соединяем найденный кусок с кратчайшей линией
'    Set vsoLayer = Application.ActiveWindow.Page.Layers.Add("temp") 'Временный слой
'    vsoLayer.CellsC(visLayerActive).FormulaU = "1" 'активируем
'    shpLotokTemp.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
'    shpShortLine.Cells("LayerMember").FormulaU = "" 'Чистим старые слои
'    vsoLayer.Add shpLotokTemp, 0 'Добавляем на временный слой
'    vsoLayer.Add shpShortLine, 0 'Добавляем на временный слой
'    ActiveWindow.DeselectAll
'    ActiveWindow.Select shpLotokTemp, visSelect
'    ActiveWindow.Select shpShortLine, visSelect
'    Application.ActiveWindow.Selection.Join 'соединяем
'    Set selSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer) 'выделяем все в слое
'    Set shpKabelPLPattern = selSelection.PrimaryItem 'Таки профит! Гребаный кабель случился!
'    'Убираем с временного слоя
'    vsoLayer.Remove shpKabelPLPattern, 0
'    'Чистим вспомогательную графику
'    vsoLayer.Delete True
'
'    'Считаем длину кабеля
'    DlinaKabelya = CableLength(shpKabelPLPattern)
'
'    'Создаем свойтва шаблона кабеля на плане
'    With shpKabelPLPattern
'        .CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = 1 'Обычная линия
'        .CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.2 mm"
'        .AddSection visSectionUser
'        .AddRow visSectionUser, visRowLast, visTagDefault
'        .CellsSRC(visSectionUser, visRowLast, visUserValue).RowNameU = "SAType"
'        .CellsSRC(visSectionUser, visRowLast, visUserValue).FormulaForceU = "90"
'        .AddSection visSectionProp
'        .AddRow visSectionProp, visRowLast, visTagDefault
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "SymName"
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Букв. обозначение"""
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Букв. обозначение"""
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "0"
'        .AddRow visSectionProp, visRowLast, visTagDefault
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "Number"
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Номер кабеля"""
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Номер кабеля"""
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "2"
'        .AddRow visSectionProp, visRowLast, visTagDefault
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "Dlina"
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Длина кабеля, м."""
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Длина кабеля, м."""
'        .CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "2"
'
'    End With
'
'    'Перебираем все кабели в датчике
'    For Each shpKabel In colCables
'        Set shpKabelPL = shpKabelPLPattern.Duplicate
'        'Сдвигаем на место
'        shpKabelPL.Cells("PinX").Formula = shpKabelPLPattern.Cells("PinX").Result(0)
'        shpKabelPL.Cells("PinY").Formula = shpKabelPLPattern.Cells("PinY").Result(0)
'        'На задний план
'        Application.ActiveWindow.Selection.SendToBack
'        'Переименовываем кабель на плане и заполняем свойства
'        With shpKabelPL
'            .name = "KabelPL." & .id
'            .Cells("Prop.SymName").Formula = IIf(shpKabel.Cells("Prop.BukvOboz").Result(0), """" & shpKabel.Cells("Prop.SymName").ResultStr(0) & """", """""")
'            .Cells("Prop.Number").Formula = shpKabel.Cells("Prop.Number").Result(0)
'            .Cells("Prop.Dlina").Formula = DlinaKabelya
'        End With
'
'        'Заполняем длину кабеля на эл.схеме (длина кабеля СВП ссылается формулой на эл.сх.)
'        shpKabel.Cells("Prop.Dlina").FormulaU = "Pages[" + shpKabelPL.ContainingPage.NameU + "]!" + shpKabelPL.NameID + "!Prop.Dlina"
'    Next
'
'    'Удаляем шаблон кабеля
'    shpKabelPLPattern.Delete
'
'    Application.ActiveWindow.DeselectAll
'
''    'Application.EndUndoScope UndoScopeID1, True
'
'End Sub
'
'Sub FillWay_OLD(ByVal clsShapePoint As classShapePoint)
''------------------------------------------------------------------------------------------------------------
'' Macros        : FillWay - Заполняет пути именами точек и длиной (рекурсивная)
''------------------------------------------------------------------------------------------------------------
'    Dim shpWay As Visio.Shape
'    Dim colWay As Collection
'    Dim BeginX As Double
'    Dim BeginY As Double
'    Dim EndX As Double
'    Dim EndY As Double
'
'
'    'Если это новая точка (не именованая) - именуем
'    If clsShapePoint.PointNumber = 0 Then clsShapePoint.PointNumber = colShapePoints.Count + 1
'    'Сканируем маршруты в этой точке (максимум 4(4 стороны света))находим ближайшие линии
'    Set selLines = ActivePage.SpatialSearch(clsShapePoint.X, clsShapePoint.Y, visSpatialTouching, 0.02 * AntiScale, 0)
'    Set colWay = New Collection
'    For Each vsoShape In selLines
'        If vsoShape.LayerCount > 0 Then
'            If vsoShape.Layer(1).name = vsoLayer.name Then
'                colWay.Add vsoShape
'            End If
'        End If
'    Next
'
'    'Перебираем маршруты
'    For Each shpWay In colWay
'
'
'        If shpWay.OneD Then '1-D фигура
'
'            'Находим точки начала и конца линии в 2D фигуре
'            BeginX = Round(shpWay.Cells("BeginX").Result(0), 14)
'            BeginY = Round(shpWay.Cells("BeginY").Result(0), 14)
'            EndX = Round(shpWay.Cells("EndX").Result(0), 14)
'            EndY = Round(shpWay.Cells("EndY").Result(0), 14)
'
'            'Нет именованных концов
'            If shpWay.Cells("Prop.Begin").Result(0) = 0 And shpWay.Cells("Prop.End").Result(0) = 0 Then
'
'                'Именуем конец в этой точке
'                If BeginX = clsShapePoint.X And BeginY = clsShapePoint.Y Then
'                    shpWay.Cells("Prop.Begin").Formula = clsShapePoint.PointNumber
'                ElseIf EndX = clsShapePoint.X And EndY = clsShapePoint.Y Then
'                    shpWay.Cells("Prop.End").Formula = clsShapePoint.PointNumber
'                End If
'
'                'Заполняем свойство длина
'                shpWay.Cells("Prop.Dlina").Formula = CableLength(shpWay)
'
'                'Берем другой конец и создаем новую точку в этих координатах
'                If shpWay.Cells("Prop.Begin").Result(0) = 0 Then
'                    Set clsShapePoint = New classShapePoint
'                    clsShapePoint.PointNumber = colShapePoints.Count + 1
'                    clsShapePoint.X = BeginX
'                    clsShapePoint.Y = BeginY
'                    colShapePoints.Add clsShapePoint, CStr(clsShapePoint.PointNumber)
'                ElseIf shpWay.Cells("Prop.End").Result(0) = 0 Then
'                    Set clsShapePoint = New classShapePoint
'                    clsShapePoint.PointNumber = colShapePoints.Count + 1
'                    clsShapePoint.X = EndX
'                    clsShapePoint.Y = EndY
'                    colShapePoints.Add clsShapePoint, CStr(clsShapePoint.PointNumber)
'                End If
'
'                'Рекурсия
'                FillWay colShapePoints(clsShapePoint.PointNumber)
'
'            'Именован один конец
'            ElseIf shpWay.Cells("Prop.Begin").Result(0) = 0 Or shpWay.Cells("Prop.End").Result(0) = 0 Then
'                'Именован другой конец - именуем наш, другой не трогаем
'                'Именуем конец в этой точке
'                If BeginX = clsShapePoint.X And BeginY = clsShapePoint.Y And shpWay.Cells("Prop.Begin").Result(0) = 0 Then
'                    shpWay.Cells("Prop.Begin").Formula = clsShapePoint.PointNumber
'                ElseIf EndX = clsShapePoint.X And EndY = clsShapePoint.Y And shpWay.Cells("Prop.End").Result(0) = 0 Then
'                    shpWay.Cells("Prop.End").Formula = clsShapePoint.PointNumber
'
'                'Именован наш конец - исключение (мы не должны попасть в точку в которой есть именованные концы)
'                ElseIf BeginX = clsShapePoint.X And BeginY = clsShapePoint.Y And shpWay.Cells("Prop.Begin").Result(0) <> 0 Then
'                    MsgBox "Именованый конец в точке: " & clsShapePoint.PointNumber & ". Конец: " & shpWay.Cells("Prop.Begin").Result(0) & ". Маршрут: " & shpWay.Cells("Prop.Begin").Result(0) & " - " & shpWay.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
'                    Exit Sub
'                ElseIf shpWay.Cells("EndX").Result(0) = clsShapePoint.X And shpWay.Cells("EndY").Result(0) = clsShapePoint.Y And shpWay.Cells("Prop.End").Result(0) <> 0 Then
'                    MsgBox "Именованый конец в точке: " & clsShapePoint.PointNumber & ". Конец: " & shpWay.Cells("Prop.End").Result(0) & ". Маршрут: " & shpWay.Cells("Prop.Begin").Result(0) & " - " & shpWay.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
'                    Exit Sub
'                End If
'
'            'Именованы оба конца
'            ElseIf shpWay.Cells("Prop.Begin").Result(0) <> 0 And shpWay.Cells("Prop.End").Result(0) <> 0 Then
'                'Маршрут уже обработан полностью (с двух концов)
'            End If
'
'
'
'        Else '2-D фигура
'
'            'Находим точки начала и конца линии в 2D фигуре
'            BeginX = Round(shpWay.Cells("PinX").Result(0) - shpWay.Cells("Width").Result(0) * 0.5 + shpWay.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 0).Result(0), 14)
'            BeginY = Round(shpWay.Cells("PinY").Result(0) - shpWay.Cells("Height").Result(0) * 0.5 + shpWay.CellsSRC(visSectionFirstComponent, visRowFirst + 1, 1).Result(0), 14)
'            EndX = Round(shpWay.Cells("PinX").Result(0) - shpWay.Cells("Width").Result(0) * 0.5 + shpWay.CellsSRC(visSectionFirstComponent, visRowLast, 0).Result(0), 14)
'            EndY = Round(shpWay.Cells("PinY").Result(0) - shpWay.Cells("Height").Result(0) * 0.5 + shpWay.CellsSRC(visSectionFirstComponent, visRowLast, 1).Result(0), 14)
'
'            'Нет именованных концов
'            If shpWay.Cells("Prop.Begin").Result(0) = 0 And shpWay.Cells("Prop.End").Result(0) = 0 Then
'
'                'Именуем конец в этой точке
'                If BeginX = clsShapePoint.X And BeginY = clsShapePoint.Y Then
'                    shpWay.Cells("Prop.Begin").Formula = clsShapePoint.PointNumber
'                ElseIf EndX = clsShapePoint.X And EndY = clsShapePoint.Y Then
'                    shpWay.Cells("Prop.End").Formula = clsShapePoint.PointNumber
'                End If
'
'                'Заполняем свойство длина
'                shpWay.Cells("Prop.Dlina").Formula = CableLength(shpWay)
'
'                'Берем другой конец и создаем новую точку в этих координатах
'                If shpWay.Cells("Prop.Begin").Result(0) = 0 Then
'                    Set clsShapePoint = New classShapePoint
'                    clsShapePoint.PointNumber = colShapePoints.Count + 1
'                    clsShapePoint.X = BeginX
'                    clsShapePoint.Y = BeginY
'                    colShapePoints.Add clsShapePoint, CStr(clsShapePoint.PointNumber)
'                ElseIf shpWay.Cells("Prop.End").Result(0) = 0 Then
'                    Set clsShapePoint = New classShapePoint
'                    clsShapePoint.PointNumber = colShapePoints.Count + 1
'                    clsShapePoint.X = EndX
'                    clsShapePoint.Y = EndY
'                    colShapePoints.Add clsShapePoint, CStr(clsShapePoint.PointNumber)
'                End If
'
'                'Рекурсия
'                FillWay colShapePoints(clsShapePoint.PointNumber)
'
'            'Именован один конец
'            ElseIf shpWay.Cells("Prop.Begin").Result(0) = 0 Or shpWay.Cells("Prop.End").Result(0) = 0 Then
'                'Именован другой конец - именуем наш, другой не трогаем
'                'Именуем конец в этой точке
'                If BeginX = clsShapePoint.X And BeginY = clsShapePoint.Y And shpWay.Cells("Prop.Begin").Result(0) = 0 Then
'                    shpWay.Cells("Prop.Begin").Formula = clsShapePoint.PointNumber
'                ElseIf EndX = clsShapePoint.X And EndY = clsShapePoint.Y And shpWay.Cells("Prop.End").Result(0) = 0 Then
'                    shpWay.Cells("Prop.End").Formula = clsShapePoint.PointNumber
'
'                'Именован наш конец - исключение (мы не должны попасть в точку в которой есть именованные концы)
'                ElseIf BeginX = clsShapePoint.X And BeginY = clsShapePoint.Y And shpWay.Cells("Prop.Begin").Result(0) <> 0 Then
'                    MsgBox "Именованый конец в точке: " & clsShapePoint.PointNumber & ". Конец: " & shpWay.Cells("Prop.Begin").Result(0) & ". Маршрут: " & shpWay.Cells("Prop.Begin").Result(0) & " - " & shpWay.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
'                    Exit Sub
'                ElseIf EndX = clsShapePoint.X And EndY = clsShapePoint.Y And shpWay.Cells("Prop.End").Result(0) <> 0 Then
'                    MsgBox "Именованый конец в точке: " & clsShapePoint.PointNumber & ". Конец: " & shpWay.Cells("Prop.End").Result(0) & ". Маршрут: " & shpWay.Cells("Prop.Begin").Result(0) & " - " & shpWay.Cells("Prop.End").Result(0), vbCritical, "Ошибка"
'                    Exit Sub
'                End If
'
'            'Именованы оба конца
'            ElseIf shpWay.Cells("Prop.Begin").Result(0) <> 0 And shpWay.Cells("Prop.End").Result(0) <> 0 Then
'                'Маршрут уже обработан полностью (с двух концов)
'            End If
'        End If
'    Next
'End Sub
'