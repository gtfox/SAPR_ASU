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
    
    
    
    Dim shpSensorFSA As Visio.Shape
    Dim shpSensorFSATemp As Visio.Shape
    Dim shpShortLine As Visio.Shape
    Dim vsoShape As Visio.Shape
    
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
    Dim PageWidth As Double
    Dim PageHeight As Double
    
    Dim i As Integer
    Dim n As Integer
    
    Dim UndoScopeID1 As Long


    Set shpSensorFSA = ActiveWindow.Selection.PrimaryItem
    
    Set colLine = New Collection
    Set colLotok = New Collection
    Set colLineShort = New Collection
    
    Set selSelection = ActiveWindow.Selection
    
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
    AddLotokToCol shpLineUp, selLineUp, colLine, colLotok
    AddLotokToCol shpLineDown, selLineDown, colLine, colLotok
    AddLotokToCol shpLineLeft, selLineLeft, colLine, colLotok
    AddLotokToCol shpLineRight, selLineRight, colLine, colLotok
    If colLotok.Count = 0 Then Exit Sub 'нет лотков - выходим
    
    'Выделяем их
    selSelection.Select shpSensorFSA, visSelect
    For Each vsoShape In colLotok
        selSelection.Select vsoShape, visSelect
    Next
    For Each vsoShape In colLine
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
    'Set selSelectionTemp = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, vsoLayer)
    
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

End Sub

Sub AddLotokToCol(shpLine As Visio.Shape, selLine As Visio.Selection, ByRef colLine As Collection, ByRef colLotok As Collection)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddLotokToCol - Заполняет коллекции лотков и линий
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim vsoShape As Visio.Shape
    Dim shpLotok As Visio.Shape
    
    For Each vsoShape In selLine 'Шейпы в выделении
        If vsoShape.Name Like "Lotok*" Then 'Нашли лоток
            If colLotok.Count = 0 Then 'Первый в коллекции
                colLotok.Add vsoShape
                i = i + 1
            Else
                For Each shpLotok In colLotok 'Лотки в коллекции
                    If vsoShape.Name Like shpLotok.Name Then 'Лоток уже есть в коллекции
                        i = i + 1
                        Exit For
                    Else
                        colLotok.Add vsoShape
                        i = i + 1
                    End If
                Next
            End If
        End If
    Next
    If i > 0 Then 'Линия пересекла лоток
        colLine.Add shpLine
    End If
End Sub


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
                Debug.Print shpTouchingShapes.Name
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

Sub Macro2()
Dim UndoScopeID1 As Long
UndoScopeID1 = Application.BeginUndoScope("Свойства слоя")

    Application.ActiveWindow.Selection.Copy

    

    
    
    Dim vsoLayer1 As Visio.Layer
    Set vsoLayer1 = Application.ActiveWindow.Page.Layers.Add("temp")
    vsoLayer1.NameU = "temp"
    vsoLayer1.CellsC(visLayerColor).FormulaU = "255"
    vsoLayer1.CellsC(visLayerStatus).FormulaU = "0"
    vsoLayer1.CellsC(visLayerVisible).FormulaU = "1"
    vsoLayer1.CellsC(visLayerPrint).FormulaU = "1"
    vsoLayer1.CellsC(visLayerActive).FormulaU = "1"
    vsoLayer1.CellsC(visLayerLock).FormulaU = "0"
    vsoLayer1.CellsC(visLayerSnap).FormulaU = "1"
    vsoLayer1.CellsC(visLayerGlue).FormulaU = "1"
    vsoLayer1.CellsC(visLayerColorTrans).FormulaU = "0%"
    
    Application.ActiveWindow.Page.Paste

    ActiveWindow.DeselectAll

    Dim vsoSelection1 As Visio.Selection
    Set vsoSelection1 = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "temp")
    Application.ActiveWindow.Selection = vsoSelection1
    
    Application.ActiveWindow.Selection.Trim
Application.EndUndoScope UndoScopeID1, True

'Application.Undo
End Sub
Sub Macro4()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Изменить размер объекта")
    Application.ActiveWindow.Page.Shapes.ItemFromID(11).CellsSRC(visSectionObject, visRowXForm1D, vis1DBeginX).FormulaU = "109 mm"
    Application.ActiveWindow.Page.Shapes.ItemFromID(11).CellsSRC(visSectionObject, visRowXForm1D, vis1DBeginY).FormulaU = "173 mm"
    Application.EndUndoScope UndoScopeID1, True

    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(11), visSelect
    Application.ActiveWindow.Selection.SendToBack

End Sub
Sub Macro5()
    Dim sel As Selection
    Set sel = Application.ActiveWindow.Selection
    sel.Offset
    Application.ActiveWindow.Selection.Move 0.143701, 0.177165

End Sub