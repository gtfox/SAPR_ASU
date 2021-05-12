'------------------------------------------------------------------------------------------------------------
' Module        : VID - Чертеж внешнего вида шкафа автоматики
' Author        : gtfox
' Date          : 2021.02.11
' Description   : Выравнивание, распределение, размеры
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
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
    For Each vsoShape In ActivePage.Shapes
        If ShapeSATypeIs(vsoShape, typeVidShkafaShkaf) Then Set shpShkaf = vsoShape: Exit For
    Next
'    Set shpShkaf = Application.ActivePage.Shapes.ItemFromID(83)
    
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
    Application.ActiveWindow.Page.Drop Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("Razmer"), 0#, 0#
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
    For Each vsoShape In ActivePage.Shapes
        If ShapeSATypeIs(vsoShape, typeVidShkafaShkaf) Then Set shpShkaf = vsoShape: Exit For
    Next
'    Set shpShkaf = Application.ActivePage.Shapes.ItemFromID(83)
    
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
    Application.ActiveWindow.Page.Drop Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("Razmer"), 0#, 0#
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

Public Sub AddElementyCxemyOnVID()
'------------------------------------------------------------------------------------------------------------
' Macros        : AddElementyCxemyOnVID - Вставляет на лист ВИД элементы со СХЕМЫ
                'В соответствии с типом элемента схемы выбирается шейп внешнего вида и добавляется на лист ВИД.
                'Шейп вида связывается со схемой, заполняются имя и описание.
                'Существующие на ВИДе элементы повторно не добавляются.
'------------------------------------------------------------------------------------------------------------
    Dim vsoPageVID As Visio.Page
    Dim vsoPageCxema As Visio.Page
    Dim colPagesCxema As Collection
    Dim shpElementOnCxema As Visio.Shape
    Dim shpElementOnVID As Visio.Shape
    Dim colElementOnVID As Collection
    Dim colElementToVID As Collection
    Dim vsoSelection As Visio.Selection
    Dim VIDvss As Document
    Dim PageParent As String
    Dim NameIdParent As String
    Dim AdrParent As String
    Dim NazvanieShemy As String
    Dim SymName As String
    Dim SAType As Integer
    Dim nCount As Double
    Dim DropX  As Double
    Dim DropY As Double
    
    
    Set colElementOnVID = New Collection
    Set colElementToVID = New Collection
    Set colPagesCxema = New Collection
    Set vsoSelection = ActiveWindow.Selection
    Set vsoPageVID = ActiveDocument.Pages("ВИД")
    Set vsoPageCxema = ActiveDocument.Pages(cListNameCxema)
    Set VIDvss = Application.Documents.Item("SAPR_ASU_VID.vss")
    
    DropX = 0
    
    NazvanieShemy = "Схема1"
    
    'Берем все листы одной схемы
    For Each vsoPageCxema In ActiveDocument.Pages
        If vsoPageCxema.Name Like cListNameCxema & "*" Then
            If vsoPageCxema.CellExists("Prop.SA_NazvanieShemy", 0) Then
                If vsoPageCxema.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then
                    colPagesCxema.Add vsoPageCxema
                End If
            End If
        End If
    Next
    
    'Находим что уже есть на ВИДе
    For Each shpElementOnVID In vsoPageVID.Shapes
        If ShapeSATypeIs(shpElementOnVID, typeVidShkafaDIN) Or ShapeSATypeIs(shpElementOnVID, typeVidShkafaDver) Then
            colElementOnVID.Add shpElementOnVID, shpElementOnVID.Cells("User.Name").ResultStr(0) '& ";" & shpElementOnVID.Cells("User.NameParent").ResultStr(0)
        End If
    Next
    
    'Суем туда же все со СХЕМЫ. Одинаковое не влезает => ошибка. Что не влезло: нам оно то и нужно
    For Each vsoPageCxema In colPagesCxema
        For Each shpElementOnCxema In vsoPageCxema.Shapes
            SAType = ShapeSAType(shpElementOnCxema)
            Select Case SAType
                Case typeCoil, typeParent, typeElement, typeTerm ', typePLCParent
                    nCount = colElementOnVID.Count
                    On Error Resume Next
                    colElementOnVID.Add shpElementOnCxema, shpElementOnCxema.Cells("User.Name").ResultStr(0) '& ";" & shpElementOnCxema.Cells("User.NameParent").ResultStr(0)
                    If colElementOnVID.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                        colElementToVID.Add shpElementOnCxema
                        nCount = colElementOnVID.Count
                    End If
                Case Else
            End Select
        Next
    Next
    
    'Вставляем на ВИД недостающие элементы
    For Each shpElementOnCxema In colElementToVID
        SAType = ShapeSAType(shpElementOnCxema)
        
        PageParent = shpElementOnCxema.ContainingPage.NameU
        NameIdParent = shpElementOnCxema.NameID
        AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
        
        
        
        Select Case SAType
            Case typeCoil, typeParent, typeElement ', typePLCParent
                SymName = shpElementOnVID.Cells("Prop.SymName").ResultStr(0)
                
                Select Case SymName
                    Case "HL" 'HL (Лампа)
                        Set shpElementOnVID = vsoPageVID.Drop(VIDvss.Masters.Item(SymName), DropX, 0)
                        shpElementOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.Name"
                        shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + shpElementOnCxema.ContainingPage.NameU + "/" + shpElementOnCxema.NameID + """"
                    Case "SA" 'SA (Переключатель)
                        Set shpElementOnVID = vsoPageVID.Drop(VIDvss.Masters.Item(SymName), DropX, 0)
                        shpElementOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.Name"
                        shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + shpElementOnCxema.ContainingPage.NameU + "/" + shpElementOnCxema.NameID + """"
                        If shpElementOnCxema.Cells("Prop.3P").Result(0) = 1 Then shpElementOnVID.Cells("Prop.TipPerkluchtelya").Formula = 3
                    Case "SB" 'SB (Кнопка)
                        Set shpElementOnVID = vsoPageVID.Drop(VIDvss.Masters.Item(SymName), DropX, 0)
                        shpElementOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.Name"
                        shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + shpElementOnCxema.ContainingPage.NameU + "/" + shpElementOnCxema.NameID + """"
                        If shpElementOnCxema.Cells("Prop.Alarm").Result(0) = 1 Then shpElementOnVID.Cells("Prop.TipKnopki").Formula = "INDEX(2,Prop.TipKnopki.Format)"
                    Case "SF" 'SF (Автомат 1ф)
                        Set shpElementOnVID = vsoPageVID.Drop(VIDvss.Masters.Item(SymName), DropX, 0)
                        shpElementOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.Name"
                        shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + shpElementOnCxema.ContainingPage.NameU + "/" + shpElementOnCxema.NameID + """"
                        shpElementOnVID.Cells("Prop.TokAvtomata").Formula = shpElementOnCxema.Cells("Prop.TokAvtomata").Result(0)
                    Case "QF" 'QF (Автомат 3ф)

                    Case "QSD" 'QSD (УЗО)
                    
                    Case "QFD" 'QFD (Дифавтомат)
                    
                    Case "QA" 'QA (Автомат защиты двигателя)
                    
                    Case "QS" 'QS (Выключатель нагрузки)
                    
                    Case "FU" 'FU (Предохранитель)
                    
                    Case "RU" 'RU (Варистор)

                    Case "KM" 'KM (Контактор электромагнитный)
                    
                    Case "KL" 'KL (Реле промежуточное)
                    
                    Case "KT" 'KT (Реле времени)
                    
                    Case "KV" 'KV (Реле напряжения)
                    
                    Case "KK" 'KK (Реле тепловое)
                    
                    Case "HA" 'HA (Звонок)
                    
                    Case "UG" 'UG (Блок питания)
                    
                    Case "TV" 'TV (Трансформатор)
                    
                    Case "UZ" 'UZ (Частотник, Твердотельное реле)

                    Case "XS" 'XS (Розетка)

                    Case "DD" 'DD (ТРМ, ПЛК-моноблок)
                    
                    Case Else
                End Select
            Case typeTerm
                'Заполнить коллекцию
                'Сгруппировать по клеммнику
                'Клеить клеммы в один клеммник
                Set shpElementOnVID = vsoPageVID.Drop(VIDvss.Masters.Item("XT"), 0, 0)
                shpElementOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.Name" '?
                
                shpElementOnVID.Cells("Prop.Sechenie").Formula = AdrParent + "!Prop.Sechenie"
                shpElementOnVID.Cells("Prop.SymName").Formula = AdrParent + "!Prop.SymName"
                shpElementOnVID.Cells("Prop.Number").Formula = AdrParent + "!Prop.Number"
                shpElementOnVID.Cells("Prop.NumberKlemmnik").Formula = AdrParent + "!Prop.NumberKlemmnik"
                
            Case Else
        End Select
        DropX = DropX + shpElementOnVID.Cells("Width").Result(0)
'        DropY = DropY + shpElementOnVID.Cells("Height").Result(0)
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
