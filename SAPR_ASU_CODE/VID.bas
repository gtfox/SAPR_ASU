'------------------------------------------------------------------------------------------------------------
' Module        : VID - Чертеж внешнего вида шкафа автоматики
' Author        : gtfox
' Date          : 2021.02.11
' Description   : Выравнивание, распределение, размеры
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------


Sub RaspredelitGorizont() '(selElemets As Visio.Section)
'------------------------------------------------------------------------------------------------------------
' Macros        : RaspredelitGorizont - Распределяет элементы на двери шкафа по горизонтали, вставляет направляющие и размеры

                'Выставляем крайние элементы, между которыми произойдет распределение. Выделяем элементы, запускаем макрос.
                'Для правильной автоматической расстановки размеров для круглых фигур первая точка = срередина
                'Для квадратных фигур первая точка = центр, 2-я = лев край, 3-я = правый, 4-я = верх, 5-я = низ
'------------------------------------------------------------------------------------------------------------
    Dim selElemets As Visio.Selection
    Dim colElemets As Collection
    Dim vsoShape As Visio.Shape
    Dim shpShkaf As Visio.Shape
    Dim shpRazmer As Visio.Shape
    Dim shpNapravl As Visio.Shape
    Dim shpElemet As Visio.Shape
    Dim celRazmer As Visio.Cell
    Dim celElemet As Visio.Cell
    Dim celNapravl As Visio.Cell
    Dim LevKrajDver As String
    Dim NapravlY As String
    Dim RowCount As Integer  'кол-во точек соединения на фигуре
    Dim i As Integer
    
    Set colElemets = New Collection

    'Находим шкаф
    Set shpShkaf = Application.ActivePage.Shapes.ItemFromID(83)
    
    'Суем в коллекцию все кроме направляющих
    For Each vsoShape In Application.ActiveWindow.Selection
        If vsoShape.Type <> visTypeGuide Then
            colElemets.Add vsoShape
        End If
    Next
    
    Application.ActiveWindow.DeselectAll
    Set selElemets = Application.ActiveWindow.Selection

    
    'Выделяем коллекцию
    For Each vsoShape In colElemets
        selElemets.Select vsoShape, visSelect
    Next
    
    'Выравниваем по вертикали (по середине фигуры) + вставка направляющей
    selElemets.Align visHorzAlignNone, visVertAlignMiddle, True
    
    'Находим появившуюся направляющую
    For Each vsoShape In Application.ActiveWindow.Selection
        If vsoShape.Type = visTypeGuide Then
            Set shpNapravl = vsoShape
            Exit For
        End If
    Next

    'Распределяем по горизонтали + вставка направляющих
    selElemets.Distribute visDistHorzSpace, True

    'Находим левый край двери шкафа
    LevKrajDver = Replace(CStr(shpShkaf.Cells("User.DoorLeft").Result("mm")), ",", ".")
    'Находим Y для направляющей
    NapravlY = Replace(CStr(shpNapravl.Cells("PinY").Result("mm")), ",", ".")
    
    'Расставляем размеры
    For Each shpElemet In selElemets
        RowCount = shpElemet.RowCount(visSectionConnectionPts)
        For i = 1 To RowCount
            Select Case i
                Case 1 'Центр
                    If RowCount = 1 Then GoSub DoSub 'Центральный размер проставляется только для круглых элементов
                Case 2, 3 'Лево, право
                    GoSub DoSub
            End Select
        Next
    Next
    
    Application.ActiveWindow.DeselectAll
    
    Exit Sub
     
DoSub:
    Application.ActiveWindow.DeselectAll
    'Вставили размер
    Application.ActiveWindow.Page.Drop Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("Razmer.37"), 0#, 0#
    Set shpRazmer = Application.ActiveWindow.Selection(1)
    'Сдвигаем текст вправо
    shpRazmer.Cells("Controls.X2").FormulaU = "=Scratch.Y12"
    shpRazmer.Cells("Controls.Y2").FormulaU = "=Height"
    'Клеим к фигуре
    Set celRazmer = shpRazmer.CellsU("EndX")
    Set celElemet = shpElemet.CellsSRC(visSectionConnectionPts, i - 1, 0)
    celRazmer.GlueTo celElemet
    'Клеим к направляющей
    Set celRazmer = shpRazmer.CellsU("BeginX")
    Set celNapravl = shpNapravl.CellsSRC(1, 1, 6)
    celRazmer.GlueTo celNapravl
    'Перемещаем ногу размера на край двери
    shpRazmer.CellsSRC(visSectionObject, visRowXForm1D, vis1DBeginX).FormulaU = "INTERSECTX(" & shpNapravl.NameID & "!PinX," & shpNapravl.NameID & "!PinY," & shpNapravl.NameID & "!Angle," & LevKrajDver & " mm," & NapravlY & " mm," & shpNapravl.NameID & "!Angle+90 deg)"
    shpRazmer.CellsSRC(visSectionObject, visRowXForm1D, vis1DBeginY).FormulaU = "INTERSECTY(" & shpNapravl.NameID & "!PinX," & shpNapravl.NameID & "!PinY," & shpNapravl.NameID & "!Angle," & LevKrajDver & " mm," & NapravlY & " mm," & shpNapravl.NameID & "!Angle+90 deg)"
    'Высота размера больше фигуры на 5 мм
    shpRazmer.Cells("Controls.Row_1.Y").Formula = Replace(CStr(shpElemet.Cells("Height").Result(0) * 0.5 + 0.196850393700787), ",", ".")
Return
End Sub

Sub VertRazmery() '(selElemets As Visio.Section)
'------------------------------------------------------------------------------------------------------------
' Macros        : VertRazmery - Вставляет вертикальные размеры для двери

                'Выделяем по одному элементу в каждой "строке" и запускаем макрос.
                'Распределение по вертикали делается выделяя только горизонтальные напрпвляющие, а не сами элементы.
                'Распределение по вертикали можно делать как до так и после вставки размеров
                'Для правильной автоматической расстановки размеров для круглых фигур первая точка = срередина
                'Для квадратных фигур первая точка = центр, 2-я = лев край, 3-я = правый, 4-я = верх, 5-я = низ
'------------------------------------------------------------------------------------------------------------

    Dim selElemets As Visio.Selection
    Dim colElemets As Collection
    Dim vsoShape As Visio.Shape
    Dim shpShkaf As Visio.Shape
    Dim shpRazmer As Visio.Shape
    Dim shpDver As Visio.Shape
    Dim shpElemet As Visio.Shape
    Dim celRazmer As Visio.Cell
    Dim celElemet As Visio.Cell
    Dim celDver As Visio.Cell
    Dim LevKrajDver As String
    Dim VerhDver As String
    Dim RowCount As Integer  'кол-во точек соединения на фигуре
    Dim i As Integer
    
    Set colElemets = New Collection

    'Находим шкаф
    Set shpShkaf = Application.ActivePage.Shapes.ItemFromID(83)
    
    'Суем в коллекцию все кроме направляющих
    For Each vsoShape In Application.ActiveWindow.Selection
        If vsoShape.Type <> visTypeGuide Then
            colElemets.Add vsoShape
        End If
    Next
    
    Application.ActiveWindow.DeselectAll
    Set selElemets = Application.ActiveWindow.Selection

    
    'Выделяем коллекцию
    For Each vsoShape In colElemets
        selElemets.Select vsoShape, visSelect
    Next

    'Находим дверь
    Set shpDver = shpShkaf.Shapes("Dver")
        
    'Находим левый край двери шкафа
    LevKrajDver = shpShkaf.Cells("User.DoorLeft").Result(0)
    'Находим верхний край двери шкафа
    VerhDver = Replace(CStr(shpShkaf.Cells("User.DoorUp").Result(0)), ",", ".")

    'Расставляем размеры
    For Each shpElemet In selElemets
        RowCount = shpElemet.RowCount(visSectionConnectionPts)
        For i = 1 To RowCount
            Select Case i
                Case 1 'Центр
                    If RowCount = 1 Then GoSub DoSub 'Центральный размер проставляется только для круглых элементов
                Case 4, 5 'Верх, низ
                    GoSub DoSub
            End Select
        Next
    Next
    
    Application.ActiveWindow.DeselectAll
    
    Exit Sub
    
DoSub:
    Application.ActiveWindow.DeselectAll
    'Вставили размер
    Application.ActiveWindow.Page.Drop Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("Razmer.37"), 0#, 0#
    Set shpRazmer = Application.ActiveWindow.Selection(1)
    'Сдвигаем текст влево
    shpRazmer.Cells("Controls.X2").FormulaU = "=Scratch.X12"
    shpRazmer.Cells("Controls.Y2").FormulaU = "=Height"
    'Клеим к фигуре
    Set celRazmer = shpRazmer.CellsU("BeginX")
    Set celElemet = shpElemet.CellsSRC(visSectionConnectionPts, i - 1, 0)
    celRazmer.GlueTo celElemet
    'Конец размера на верх двери
    shpRazmer.Cells("EndX").Formula = shpRazmer.Cells("BeginX").Result(0)
    shpRazmer.Cells("EndY").Formula = VerhDver
    'Формула для нахождения точки приклеивания
    shpRazmer.Cells("User.PntToGlue").FormulaU = "PNTX(LOCTOLOC(PNT(EndX,EndY),ThePage!PageWidth," & shpDver.NameID & "!Width))/" & shpDver.NameID & "!Width"
    DoEvents
    'Клеим к верху двери
    Set celRazmer = shpRazmer.CellsU("EndX")
    celRazmer.GlueToPos shpDver, shpRazmer.Cells("User.PntToGlue").Result(0), 1#
    'Высота размера 8 мм от края двери
    shpRazmer.Cells("Controls.Row_1.Y").Formula = Replace(CStr(shpRazmer.Cells("User.PntToGlue").Result(0) * shpDver.Cells("Width").Result(0) + 0.31496062992126), ",", ".")
Return
End Sub

Sub VpisatVList()
'------------------------------------------------------------------------------------------------------------
' Macros        : VpisatVList - "Вписывает чертеж в лист" :) Увеличивает масштаб докумета под размер чертежа

                'Рисуем прямоугольник больше размера чертежа. Размер прямоугольника - это будущий размер листа. Запускаем макрос.
                'Масштаб и размер докумета меняются, прямоугольник удаляется.
'------------------------------------------------------------------------------------------------------------
    Dim vsoShape As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim kW As Double
    Dim kH As Double
    Dim k As Double
    
    Set vsoPage = Application.ActivePage
    Set vsoShape = Application.ActiveWindow.Selection(1)
    
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = "420 mm"
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = "297 mm"
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawingScale).FormulaU = "1 mm"
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawScaleType).FormulaU = "0"
    
    kW = vsoShape.Cells("Width").Result(0) / vsoPage.PageSheet.Cells("PageWidth").Result(0)
    kH = vsoShape.Cells("Height").Result(0) / vsoPage.PageSheet.Cells("PageHeight").Result(0)
    k = IIf(kW > kH, kW, kH)
    With vsoPage.PageSheet
        .CellsSRC(visSectionObject, visRowPage, visPageDrawingScale).FormulaU = Replace(CStr(k), ",", ".") & " mm"
        .CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = Replace(CStr(.CellsSRC(visSectionObject, visRowPage, visPageWidth).Result("mm") * k), ",", ".") & " mm"
        .CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = Replace(CStr(.CellsSRC(visSectionObject, visRowPage, visPageHeight).Result("mm") * k), ",", ".") & " mm"
        .CellsSRC(visSectionObject, visRowPage, visPageDrawScaleType).FormulaU = "3"
    End With
    vsoShape.Delete
End Sub

Sub Macro5()
'Добавление connection points
    Application.ActiveWindow.Shape.AddSection visSectionConnectionPts
    Application.ActiveWindow.Shape.AddRow visSectionConnectionPts, visRowLast, visTagDefault
    Application.ActiveWindow.Shape.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctX).FormulaForceU = "Width*0"
    Application.ActiveWindow.Shape.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctY).FormulaForceU = "Height*0"
    Application.ActiveWindow.Shape.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirX).FormulaForceU = "0 mm"
    Application.ActiveWindow.Shape.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirY).FormulaForceU = "0 mm"
    Application.ActiveWindow.Shape.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctType).FormulaForceU = "0 mm"
    Application.ActiveWindow.Shape.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctAutoGen).FormulaForceU = "0 mm"
    Application.ActiveWindow.Shape.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctType).FormulaU = 2

End Sub
