'------------------------------------------------------------------------------------------------------------
' Module        : VID - Чертеж внешнего вида шкафа автоматики
' Author        : gtfox
' Date          : 2021.02.11
' Description   : Выравнивание, распределение, размеры
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------


Sub Raspredelit() '(selElemets As Visio.Section, Nrazm As Integer)
'------------------------------------------------------------------------------------------------------------
' Macros        : Raspredelit - Распределяет элементы на двери шкафа по горизонтали, вставляет направляющие и размеры

                'Выставляем крайние элементы, между которыми произойдет распределение. Выделяем элементы, запускаем макрос.
'------------------------------------------------------------------------------------------------------------
    'Для правильной автоматической расстановки размеров для круглых фигур первая точка = срередина
    'Для квадратных фигур первая точка = лев край, 2-я = правый, 3-я = верх, 4-я = низ, 5-я = центр
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
    Dim Nrazm As Integer  'кол-во проставляемых размеров для фигуры
    Dim i As Integer
    
    Set colElemets = New Collection
    
    Nrazm = 2

    'Находим шкаф
    Set shpShkaf = Application.ActivePage.Shapes.ItemFromID(181)
    
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
        For i = 1 To Nrazm
            If shpElemet.CellsSRCExists(visSectionConnectionPts, i - 1, 0, False) Then
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
                
            End If
        Next
    Next
    
    Application.ActiveWindow.DeselectAll

End Sub

Sub VertRazmery() '(selElemets As Visio.Section, Nrazm As Integer)
'------------------------------------------------------------------------------------------------------------
' Macros        : VertRazmery - Вставляет вертикальные размеры для двери

                'Выделяем по одному элементу в каждой "строке" и запускаем макрос.
                'Распределение по вертикали делается выделяя только горизонтальные напрпвляющие, а не сами элементы.
                'Распределение по вертикали можно делать как до так и после вставки размеров
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
    Dim Nrazm As Integer  'кол-во проставляемых размеров для фигуры
    Dim i As Integer
    
    Set colElemets = New Collection
    
    Nrazm = 4

    'Находим шкаф
    Set shpShkaf = Application.ActivePage.Shapes.ItemFromID(181)
    
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
    LevKrajDver = Replace(CStr(shpShkaf.Cells("User.DoorLeft").Result("mm")), ",", ".")
    'Находим верхний край двери шкафа
    VerhDver = Replace(CStr(shpShkaf.Cells("User.DoorUp").Result("mm")), ",", ".")
    
    'Расставляем размеры
    For Each shpElemet In selElemets
        For i = 3 To Nrazm
            If shpElemet.CellsSRCExists(visSectionConnectionPts, i - 1, 0, False) Then
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
                'Клеим к верху двери
                Set celRazmer = shpRazmer.CellsU("EndX")
                celRazmer.GlueToPos shpDver, 0.232343, 1#
                
                
                'Перемещаем ногу размера на край двери
                shpRazmer.CellsSRC(visSectionObject, visRowXForm1D, vis1DBeginX).FormulaU = "INTERSECTX(" & shpNapravl.NameID & "!PinX," & shpNapravl.NameID & "!PinY," & shpNapravl.NameID & "!Angle," & LevKrajDver & " mm," & NapravlY & " mm," & shpNapravl.NameID & "!Angle+90 deg)"
                shpRazmer.CellsSRC(visSectionObject, visRowXForm1D, vis1DBeginY).FormulaU = "INTERSECTY(" & shpNapravl.NameID & "!PinX," & shpNapravl.NameID & "!PinY," & shpNapravl.NameID & "!Angle," & LevKrajDver & " mm," & NapravlY & " mm," & shpNapravl.NameID & "!Angle+90 deg)"
                'Высота размера 8 мм
                shpRazmer.Cells("Controls.Row_1.Y").Formula = Replace(CStr(shpElemet.Cells("Height").Result(0) * 0.5 + 0.31496062992126), ",", ".")
            End If
        Next
    Next
    
    Application.ActiveWindow.DeselectAll

End Sub

Sub Macro9()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(83).CellsU("BeginX")
    Set vsoCell2 = Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(7, 2, 0)
    vsoCell1.GlueTo vsoCell2
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell3 As Visio.Cell
    Dim vsoCell4 As Visio.Cell
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(83).CellsU("EndX")
    vsoCell3.GlueToPos Application.ActiveWindow.Page.Shapes.ItemFromID(189), 0.232343, 1#
    Application.EndUndoScope UndoScopeID2, True

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
