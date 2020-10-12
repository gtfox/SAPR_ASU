
'------------------------------------------------------------------------------------------------------------
' Module        : KabeliPLAN - Кабели на планах
' Author        : gtfox
' Date          : 2020.10.09
' Description   : Автопрокладка кабелей по лоткам, подсчет длины, выноски кабелей
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------



Public Sub RouteCable() '(shpSensorFSA As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : RouteCable - Прокладывает кабель по ближайшему лотку
                'Определяет ближайший лоток и прокладывает кабель до шкафа
'------------------------------------------------------------------------------------------------------------
    
    Dim shpKabel As Visio.Shape
    
    Dim colWires As Collection
    Dim colWiresIO As Collection
    Dim vsoMaster As Visio.Master
    
    Dim shpLotok As Visio.Shape
    Dim shpLotokTemp As Visio.Shape
    Dim shpSensor As Visio.Shape
    Dim shpSensorFSA As Visio.Shape
    Dim shpSensorFSATemp As Visio.Shape
    Dim shpShortLine As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim vsoShapeTemp As Visio.Shape
    Dim vsoCollection As Collection
    
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
    Dim selLines As Visio.Selection
    
    Dim colLine As Collection
    Dim colLotok As Collection
    Dim colLineShort As Collection
    
    Dim vsoLayer As Visio.Layer
    Dim vsoShapeLayer As Visio.Layer
    
    Dim SensorFSAPinX As Double
    Dim SensorFSAPinY As Double
    Dim dXSensorFSAPinX As Double
    Dim dYSensorFSAPinY As Double
    Dim BoxX As Double
    Dim BoxY As Double
    Dim LineX As Double
    Dim LineY As Double
    Dim PageWidth As Double
    Dim PageHeight As Double
    
    Dim BoxNumber As Integer 'Номер шкафа к которому подключен кабель/датчик
    Dim i As Integer
    Dim n As Integer
    
    Dim UndoScopeID1 As Long


    Set shpSensorFSA = ActiveWindow.Selection.PrimaryItem
    
    Set colLine = New Collection
    Set colLotok = New Collection
    Set colLineShort = New Collection
    Set vsoCollection = New Collection
    
    Set selSelection = ActiveWindow.Selection
    Set shpSensor = ShapeByHyperLink(shpSensorFSA.Cells("Hyperlink.Shema.SubAddress").ResultStr(0))
    'Находим кабель подключенный к датчику
    For Each vsoShape In shpSensor.Shapes 'Перебираем все входы датчика
        If vsoShape.Name Like "SensorIO*" Then
            'Находим подключенные провода
            Set vsoCollection = FillColWires(vsoShape)
            'По проводу находим кабель
            Set shpKabel = vsoCollection.Item(1).Parent 'Кабель (Предпологается что 1 датчик подключен к 1 шкафу (даже многокабельный))
            Exit For
        End If
    Next
    'Шкаф к которому подключен кабель
    BoxNumber = shpKabel.Cells("User.LinkToBox").Result(0)
    
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
    AddLotokToCol shpLineUp, selLineUp, colLine, colLotok, BoxNumber
    AddLotokToCol shpLineDown, selLineDown, colLine, colLotok, BoxNumber
    AddLotokToCol shpLineLeft, selLineLeft, colLine, colLotok, BoxNumber
    AddLotokToCol shpLineRight, selLineRight, colLine, colLotok, BoxNumber
    If colLotok.Count = 0 Then 'нет лотков - выходим
        'Чистим вспомогательную графику
        shpLineUp.Delete
'        shpLineDown.Delete
        shpLineLeft.Delete
        shpLineRight.Delete
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
        If vsoShape.Name Like "SensorFSA*" Then
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
    Set selLines = ActivePage.SpatialSearch(SensorFSAPinX, SensorFSAPinY, visSpatialTouching, 0.02, 0)
    For Each vsoShape In selLines
        If vsoShape.LayerCount > 0 Then
            If vsoShape.Layer(1).Name = vsoLayer.Name Then
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
     
    'Находим лоток идущий в наш шкаф и которого касается кратчайшая линия
    Set selSelection = shpShortLine.SpatialNeighbors(visSpatialTouching, 0, 0)
    For Each vsoShape In selSelection 'Шейпы в выделении
        If (vsoShape.Name Like "Lotok*") And (LotokToBox(vsoShape, BoxNumber)) Then 'Нашли лоток
            'Находим координаты точки в которой лоток подключен к шкафу
            For i = 1 To vsoShape.Connects.Count 'Перебираем подключенные концы лотка
                If vsoShape.Connects(i).ToSheet.Name Like "Box*" Then 'Выбираем только шкафы
                    If vsoShape.Connects(i).FromCell.Name Like "Begin*" Then
                        BoxX = vsoShape.Connects(i).FromSheet.Cells("BeginX").Result(0)
                        BoxY = vsoShape.Connects(i).FromSheet.Cells("BeginY").Result(0)
                    ElseIf vsoShape.Connects(i).FromCell.Name Like "End*" Then
                        BoxX = vsoShape.Connects(i).FromSheet.Cells("EndX").Result(0)
                        BoxY = vsoShape.Connects(i).FromSheet.Cells("EndY").Result(0)
                    End If
                End If
            Next
            Set shpLotok = vsoShape
            Exit For
        End If
    Next
    
    'Находим координаты точки в которой лоток пересекается с кратчайшей линией
    If shpShortLine.Cells("BeginX").Result(0) = SensorFSAPinX Then
        LineX = shpShortLine.Cells("EndX").Result(0)
        LineY = shpShortLine.Cells("EndY").Result(0)
    ElseIf shpShortLine.Cells("EndX").Result(0) = SensorFSAPinX Then
        LineX = shpShortLine.Cells("BeginX").Result(0)
        LineY = shpShortLine.Cells("BeginY").Result(0)
    End If
    
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
        If vsoShape.Name Like "SensorFSA*" Then
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
        Set selSelectionTemp = vsoShape.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
        For Each vsoShapeTemp In selSelectionTemp
            If vsoShapeTemp.Name Like "Box*" Then
                If vsoShapeTemp.Cells("Prop.BoxNumber").Result(0) = BoxNumber Then
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
    selSelection.SendToBack 'на задний план
    Set shpKabel = selSelection.PrimaryItem 'Таки профит! Гребаный кабель случился!
    'Убираем с временного слоя
    vsoLayer.Remove shpKabel, 0
    'Чистим вспомогательную графику
    vsoLayer.Delete True
    
    
    'Переименовываем и заполняем свойтва кабеля
    shpKabel.Name = "CablePL." & shpKabel.ID
    
    'Считаем длину кабеля
    
    
    
    
    
    


'    'Application.EndUndoScope UndoScopeID1, True
'    For Each vsoShape In ActivePage.Shapes
'        n = vsoShape.LayerCount
'        If n > 0 Then
'            For i = 1 To n
'                Set vsoShapeLayer = vsoShape.Layer(i)
'                If vsoShapeLayer.Name = vsoLayer.Name Then
'
'                End If
'            Next
'        End If
'    Next

'    Application.ActiveWindow.Selection.SendToBack
' Application.ActiveWindow.Shape.CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = 2

End Sub

Sub AddLotokToCol(shpLine As Visio.Shape, selLine As Visio.Selection, ByRef colLine As Collection, ByRef colLotok As Collection, BoxNumber As Integer)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddLotokToCol - Заполняет коллекции лотков и линий
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim vsoShape As Visio.Shape
    Dim shpLotok As Visio.Shape
    
    For Each vsoShape In selLine 'Шейпы в выделении
        If vsoShape.Name Like "Lotok*" Then 'Нашли лоток
            If (colLotok.Count = 0) And (LotokToBox(vsoShape, BoxNumber)) Then 'Первый в коллекции
                colLotok.Add vsoShape
                i = i + 1
            Else
                For Each shpLotok In colLotok 'Лотки в коллекции
                    If vsoShape.Name Like shpLotok.Name Then 'Лоток уже есть в коллекции
                        i = i + 1
                        Exit For
                    Else
                        If LotokToBox(vsoShape, BoxNumber) Then
                            colLotok.Add vsoShape
                            i = i + 1
                        End If
                    End If
                Next
            End If
        End If
    Next
    If i > 0 Then 'Линия пересекла лоток
        colLine.Add shpLine
    End If
End Sub

Function LotokToBox(shpLotok As Visio.Shape, BoxNumber As Integer) As Boolean
'------------------------------------------------------------------------------------------------------------
' Function        : LotokToBox - Проверяет что лоток приклеен к нужному шкафу на плане
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
            For i = 1 To shpLotok.Connects.Count 'Перебираем подключенные концы лотка
                If shpLotok.Connects(i).ToSheet.Name Like "Box*" Then 'Выбираем только шкафы
                    If shpLotok.Connects(i).ToSheet.Cells("Prop.BoxNumber").Result(0) = BoxNumber Then 'Сравниваем номер шкафа
                        LotokToBox = True
                        Exit Function
                    End If
                End If
            Next
            LotokToBox = False
End Function


Public Sub CableInfoPlan(Connects As IVConnects)
'------------------------------------------------------------------------------------------------------------
' Macros        : CableInfoPlan - Заполняет выноску на плане (лоток и кабели)
                'Клеим выноску на лоток, по которому проложены кабели
                'Выноска заполняется названием лотка и номерами кабелей
'------------------------------------------------------------------------------------------------------------
    Dim shpVynoska As Visio.Shape
    Dim shpTouchingShapes As Visio.Shape
    Dim vsoSelection As Visio.Selection
    Dim strProvoda As String
    Dim strLotok As String
    Dim colNum As Collection
    Dim mNum() As Integer
    Dim NumTemp As Variant
    Dim i As Integer
    Dim j As Integer
    Dim UbNum As Long
    
    Set colNum = New Collection
    Set shpVynoska = Connects.FromSheet
    strProvoda = "("
    
    Select Case shpVynoska.Connects.Count 'кол-во соединенных концов у выноски
        Case 0 'Нет соединенных концов - отцепили
                shpVynoska.Cells("Prop.Lotok").FormulaU = """"""
                shpVynoska.Cells("Prop.Provoda").FormulaU = """"""
        Case 1, 2 'С одной стороны
            Set vsoSelection = shpVynoska.ContainingPage.SpatialSearch(shpVynoska.Cells("EndX").Result(0), shpVynoska.Cells("EndY").Result(0), visSpatialTouching, 0.02, 0)
            For Each shpTouchingShapes In vsoSelection
                'Debug.Print shpTouchingShapes.Name
                If shpTouchingShapes.Name Like "w*" Then
                    colNum.Add shpTouchingShapes.Cells("Prop.Number").Result(0)
                ElseIf shpTouchingShapes.Name Like "Lotok*" Then
                    strLotok = shpTouchingShapes.Cells("User.FullName").ResultStr(0)
                End If
            Next

        'Case 2 'С двух сторон - не обрабатываем 2-ю сторону
    End Select
    If colNum.Count = 0 Then Exit Sub
    'из коллекции передаем номера проводов в массив для сортировки
    ReDim mNum(colNum.Count - 1)
    i = 0
    For Each NumTemp In colNum
        mNum(i) = NumTemp
        i = i + 1
    Next
    
    ' "Сортировка вставками" номеров проводов
    '--V--Сортируем по возрастанию номеров проводов
    UbNum = UBound(mNum)
    For j = 1 To UbNum
        NumTemp = mNum(j)
        i = j
        While mNum(i - 1) > NumTemp '>:возрастание, <:убывание
            mNum(i) = mNum(i - 1)
            i = i - 1
            If i <= 0 Then GoTo ExitWhileX
        Wend
ExitWhileX: mNum(i) = NumTemp
    Next
    '--Х--Сортировка по возрастанию номеров проводов
 
    For i = 0 To UbNum
        strProvoda = strProvoda & mNum(i) & ";"
    Next
                    
    strProvoda = Left(strProvoda, Len(strProvoda) - 1)
    If Len(strProvoda) > 1 Then
        strProvoda = strProvoda & ")"
    End If

    shpVynoska.Cells("Prop.Lotok").FormulaU = """" & strLotok & """"
    shpVynoska.Cells("Prop.Provoda").FormulaU = """" & strProvoda & """"
    
End Sub