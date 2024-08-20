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
    Const PodemRazmera = 50 / 25.4
    
    Set colElemets = New Collection

    'Находим шкаф
    For Each vsoShape In ActivePage.Shapes
        If ShapeSATypeIs(vsoShape, typeVidShkafaShkaf) Then Set shpShkaf = vsoShape: Exit For
    Next
    If shpShkaf Is Nothing Then Exit Sub
    
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
                    shpElemet.Shapes("Desc").Cells("Geometry1.NoFill").Formula = 0 'Непрозрачное описание
                    shpElemet.BringToFront 'На передние план
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
    shpRazmer.Cells("Controls.Row_1.Y").Formula = Replace(CStr(shpElemet.Cells("Height").Result(0) * 0.5 + PodemRazmera), ",", ".")
    'Укорачиваем ноги размеру
'    shpRazmer.Cells("Scratch.C4").Formula = "13mm"
'    shpRazmer.Cells("Prop.Kasanie").Formula = 0
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
    If shpShkaf Is Nothing Then Exit Sub
    
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

Public Sub VpisatVList()
'------------------------------------------------------------------------------------------------------------
' Macros        : VpisatVList - Запуск макроса VpisatVListExec
'------------------------------------------------------------------------------------------------------------
    Application.DoCmd visCmdDRRectTool 'Рисование прямоугольника
    ThisDocument.bVpisatVList = True
End Sub

Sub VpisatVListExec(vsoShape As Visio.Shape, iFormat As Long)
'------------------------------------------------------------------------------------------------------------
' Macros        : VpisatVListExec - "Вписывает чертеж в лист" - Увеличивает масштаб докумета под размер чертежа

                'Рисуем прямоугольник больше размера чертежа. Размер прямоугольника - это будущий размер листа. Запускаем макрос.
                'Масштаб и размер докумета меняются, прямоугольник удаляется.
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim kW As Double
    Dim kH As Double
    Dim k As Double
    Dim W As String
    Dim H As String
    
    Set vsoPage = Application.ActivePage
    
    Select Case iFormat
        Case 0
            W = "1189 mm"
            H = "841 mm"
        Case 1
            W = "841 mm"
            H = "594 mm"
        Case 2
            W = "594 mm"
            H = "420 mm"
        Case 3
            W = "420 mm"
            H = "297 mm"
        Case 4
            W = "297 mm"
            H = "210 mm"
    End Select
    
    
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = W
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = H
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawingScale).FormulaU = "1 mm"
    vsoPage.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawScaleType).FormulaU = "0"
    
    kW = vsoShape.Cells("Width").Result(0) / vsoPage.PageSheet.Cells("PageWidth").Result(0)
    kH = vsoShape.Cells("Height").Result(0) / vsoPage.PageSheet.Cells("PageHeight").Result(0)
    k = IIf(kW > kH, kW, kH)
    With vsoPage.PageSheet
        .CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = Replace(CStr(.CellsSRC(visSectionObject, visRowPage, visPageWidth).Result("mm") * k), ",", ".") & " mm"
        .CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = Replace(CStr(.CellsSRC(visSectionObject, visRowPage, visPageHeight).Result("mm") * k), ",", ".") & " mm"
        .CellsSRC(visSectionObject, visRowPage, visPageDrawScaleType).FormulaU = "3"
        .CellsSRC(visSectionObject, visRowPage, visPageDrawingScale).FormulaU = Replace(CStr(k), ",", ".") & " mm"
    End With
    vsoShape.Delete
    Application.DoCmd visCmdSelectionModeRect 'Возврат мыши
End Sub

Public Sub PageVIDAddElementsFrm()
    Load frmPageVIDAddElements
    frmPageVIDAddElements.Show
End Sub


Public Sub AddElementyCxemyOnVID(NazvanieShkafa As String)
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
    Dim shpTermOnCxema As Visio.Shape
    Dim shpTermOnVID As Visio.Shape
    Dim colTermToVID As Collection
    Dim vsoSelection As Visio.Selection
    Dim VIDvss As Document
    Dim cellKlemmnik As Visio.Cell
    Dim cellKlemma As Visio.Cell
    Dim PageParent As String
    Dim NameIdParent As String
    Dim AdrParent As String, GUIDParent As String
'    Dim NazvanieShkafa As String
    Dim SymName As String
    Dim SAType As Integer
    Dim nCount As Double
    Dim DropX  As Double
    Dim DropY As Double
    Dim dX As Double
    Dim dY As Double
    Dim n As Integer
    Dim ElementovVStroke As Integer 'Количество элементов в одной "строке" при вставке на ВИД
    
    If NazvanieShkafa = "" Then
        MsgBox "Нет шкафа для вставки. Название шкафа пустое", vbExclamation, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set colElementOnVID = New Collection
    Set colElementToVID = New Collection
    Set colTermToVID = New Collection
    Set colPagesCxema = New Collection
    Set vsoSelection = ActiveWindow.Selection
    Set vsoPageVID = Application.ActivePage  'Pages("ВИД")
'    Set vsoPageCxema = ActiveDocument.Pages(cListNameCxema)
    Set VIDvss = Application.Documents.Item("SAPR_ASU_VID.vss")
    
    PageName = cListNameCxema
    ElementovVStroke = 10
    
    'Находим что уже есть на ВИДе
    For Each shpElementOnVID In vsoPageVID.Shapes
        If ShapeSATypeIs(shpElementOnVID, typeVidShkafaDIN) Or ShapeSATypeIs(shpElementOnVID, typeVidShkafaDver) Then
            If shpElementOnVID.CellExists("User.NameKlemmnik", 0) Then
                colElementOnVID.Add shpElementOnVID, shpElementOnVID.Cells("User.NameParent").ResultStr(0) & ";" & shpElementOnVID.Cells("User.Shkaf").ResultStr(0)
            Else
                colElementOnVID.Add shpElementOnVID, shpElementOnVID.Cells("User.NameParent").ResultStr(0) '& ";" & shpElementOnVID.Cells("User.NameParent").ResultStr(0)
            End If
        End If
    Next
    
    'Суем туда же все со СХЕМЫ. Одинаковое не влезает => ошибка. Что не влезло: нам оно то и нужно
    For Each vsoPageCxema In ActiveDocument.Pages
        If vsoPageCxema.name Like PageName & "*" Then
            For Each shpElementOnCxema In vsoPageCxema.Shapes
                If shpElementOnCxema.CellExists("User.Shkaf", 0) Then
                    If shpElementOnCxema.Cells("User.Shkaf").ResultStr(0) = NazvanieShkafa Then
                        SAType = ShapeSAType(shpElementOnCxema)
                        Select Case SAType
                            Case typeCxemaCoil, typeCxemaParent, typeCxemaElement, typePLCParent
                                nCount = colElementOnVID.Count
                                On Error Resume Next
                                colElementOnVID.Add shpElementOnCxema, shpElementOnCxema.Cells("User.Name").ResultStr(0) '& ";" & shpElementOnCxema.Cells("User.NameParent").ResultStr(0)
                                err.Clear
                                On Error GoTo 0
                                If colElementOnVID.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                                    colElementToVID.Add shpElementOnCxema
                                End If
                            Case typeCxemaTerm
                                nCount = colElementOnVID.Count
                                On Error Resume Next
                                colElementOnVID.Add shpElementOnCxema, shpElementOnCxema.Cells("User.FullName").ResultStr(0) & ";" & shpElementOnCxema.Cells("User.Shkaf").ResultStr(0)
                                err.Clear
                                On Error GoTo 0
                                If colElementOnVID.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                                    colElementToVID.Add shpElementOnCxema
                                End If
                            Case Else
                        End Select
                    End If
                End If
            Next
        End If
    Next
'-------------------------------------------------------------------------------Клеммы---------------------------------------------------------------------------------------------
    'Клеммы шкафа собираем в отдельную коллекцию
    For Each shpElementOnCxema In colElementToVID
        If ShapeSATypeIs(shpElementOnCxema, typeCxemaTerm) Then
            colTermToVID.Add shpElementOnCxema
        End If
    Next
    
    'Вставляем на ВИД недостающие клеммы группируя по клеммникам
    If colTermToVID.Count > 0 Then
        KlemmnikName = colTermToVID.Item(colTermToVID.Count).Cells("User.KlemmnikName").ResultStr(0)
    End If
    KlemmnikNameOld = ""
    While colTermToVID.Count > 0
        'Находим и вставляем все клеммы текущего клеммника
        For Each shpTermOnCxema In colTermToVID
            If shpTermOnCxema.Cells("User.KlemmnikName").ResultStr(0) = KlemmnikName Then
            
                Set shpTermOnVID = vsoPageVID.Drop(VIDvss.Masters.Item("XT"), 0, DropY)
                
                If KlemmnikName <> KlemmnikNameOld Then 'Первая клемма нового клеммника
                    shpTermOnVID.Cells("Prop.Nachalo").Formula = 1
                    Set cellKlemmnik = shpTermOnVID.CellsSRC(visSectionConnectionPts, visRowConnectionPts + 1, 0)
                    KlemmnikNameOld = KlemmnikName
                End If
                
                'Заполняем данные
                PageParent = shpTermOnCxema.ContainingPage.NameU
                NameIdParent = shpTermOnCxema.NameID
                AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
                shpTermOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.FullName"
                shpTermOnVID.Cells("User.Name").Formula = AdrParent + "!User.Name" '?
                shpTermOnVID.Cells("User.Shkaf").Formula = AdrParent + "!User.Shkaf"
                shpTermOnVID.Cells("User.Mesto").Formula = AdrParent + "!User.Mesto"
                shpTermOnVID.Cells("Prop.Sechenie").Formula = AdrParent + "!Prop.Sechenie"
                shpTermOnVID.Cells("Prop.SymName").Formula = AdrParent + "!Prop.SymName"
                shpTermOnVID.Cells("Prop.Number").Formula = AdrParent + "!Prop.Number"
                shpTermOnVID.Cells("Prop.NumberKlemmnik").Formula = AdrParent + "!Prop.NumberKlemmnik"
                
                'Клеим клемму к клеммнику
                Set cellKlemma = shpTermOnVID.CellsSRC(visSectionConnectionPts, visRowConnectionPts, 0)
'                cellKlemma.GlueTo cellKlemmnik
                'Костыль из-за того, что visio не может таскать длинные цепочки склеенных фигур
                If cellKlemmnik.Shape <> cellKlemma.Shape Then
                    cellKlemma.Shape.Cells("PinX").Formula = cellKlemmnik.Shape.NameID & "!PinX+" & cellKlemmnik.Shape.NameID & "!Width"
                    cellKlemma.Shape.Cells("PinY").Formula = cellKlemmnik.Shape.NameID & "!PinY"
                End If
                Set cellKlemmnik = shpTermOnVID.CellsSRC(visSectionConnectionPts, visRowConnectionPts + 1, 0)
            End If
        Next
        
        'Удаляем из коллекции вставленный клеммник
        For i = colTermToVID.Count To 1 Step -1
            If colTermToVID.Item(i).Cells("User.KlemmnikName").ResultStr(0) = KlemmnikName Then
                colTermToVID.Remove i
            End If
        Next
        If colTermToVID.Count > 0 Then
            'Берем следующий клеммник
            KlemmnikName = colTermToVID.Item(colTermToVID.Count).Cells("User.KlemmnikName").ResultStr(0)
            'Смещаемся ниже
            DropY = DropY - shpTermOnVID.Cells("Height").Result(0) * 2
        End If
    Wend

'-------------------------------------------------------------------------------Элементы---------------------------------------------------------------------------------------------

    'Вставляем на ВИД недостающие элементы
    For Each shpElementOnCxema In colElementToVID
        SAType = ShapeSAType(shpElementOnCxema)
        
        PageParent = shpElementOnCxema.ContainingPage.NameU
        NameIdParent = shpElementOnCxema.NameID
        AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
        GUIDParent = shpElementOnCxema.UniqueID(visGetOrMakeGUID)
        SymName = shpElementOnCxema.Cells("Prop.SymName").ResultStr(0)
        
        Select Case SAType
            Case typeCxemaCoil, typeCxemaParent, typeCxemaElement ', typePLCParent
                
                On Error Resume Next
                Set shpElementOnVID = vsoPageVID.Drop(VIDvss.Masters.Item(SymName & IIf(shpElementOnCxema.NameU Like SymName & "3P*", "3P", "")), DropX, DropY)
                If err.Number <> 0 Then
                    err.Clear
                    On Error GoTo 0
                    MsgBox "Элемент схемы " & shpElementOnCxema.NameU & " не имеет чертёж внешнего вида", vbExclamation + vbOKOnly, "САПР-АСУ: отсутствует чертеж внешнего вида элемента схемы"
                Else
                    shpElementOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.Name"
                    shpElementOnVID.Cells("User.Name").Formula = AdrParent + "!Prop.SymName&" + AdrParent + "!Prop.Number"
                    shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + shpElementOnCxema.ContainingPage.NameU + "/" + shpElementOnCxema.NameID + """"
                    shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkExtraInfo).FormulaU = GUIDParent
                    shpElementOnVID.Shapes("Desc").text = shpElementOnCxema.Shapes("Desc").text 'Здесь не ссылка, т.к. на щите надписи могут отличаться от схемы
                    shpElementOnVID.Cells("Prop.ShowDesc").Formula = 1
                    dX = shpElementOnVID.Cells("Width").Result(0)
                    dY = IIf(shpElementOnVID.Cells("Height").Result(0) > dY, shpElementOnVID.Cells("Height").Result(0), dY)
                    Select Case SymName
                        Case "HL" 'HL (Лампа)
                        
                        Case "SA" 'SA (Переключатель)
                            If shpElementOnCxema.Cells("Prop.3P").Result(0) = 1 Then shpElementOnVID.Cells("Prop.TipPerkluchtelya").Formula = 3
                        Case "SB" 'SB (Кнопка)
                            If shpElementOnCxema.Cells("Prop.Alarm").Result(0) = 1 Then shpElementOnVID.Cells("Prop.TipKnopki").FormulaU = "INDEX(2,Prop.TipKnopki.Format)" ' """Аварийная"""
                        Case "SF" 'SF (Автомат 1ф)
                            shpElementOnVID.Cells("Prop.TokAvtomata").Formula = AdrParent + "!Prop.Tok"
                        Case "QF" 'QF (Автомат 3ф)
                            shpElementOnVID.Cells("Prop.TokAvtomata").Formula = AdrParent + "!Prop.Tok"
                        Case "QSD" 'QSD (УЗО)
                            shpElementOnVID.Cells("Prop.Polusov").Formula = AdrParent + "!Prop.Polusov"
                        Case "QFD" 'QFD (Дифавтомат)
                            shpElementOnVID.Cells("Prop.Polusov").Formula = AdrParent + "!Prop.Polusov"
                        Case "QA" 'QA (Автомат защиты двигателя)
                            shpElementOnVID.Cells("Prop.TipAvtomata").Formula = AdrParent + "!Prop.Harakteristika"
                        Case "QS" 'QS (Выключатель нагрузки)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "FU" 'FU (Предохранитель)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "RU" 'RU (Варистор)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "KM" 'KM (Контактор электромагнитный)
                            shpElementOnVID.Cells("Prop.TokKontaktora").Formula = AdrParent + "!Prop.Tok"
                        Case "KL" 'KL (Реле промежуточное)
                            shpElementOnVID.Cells("Prop.Kontaktov").Formula = AdrParent + "!Prop.Kontaktov"
                        Case "KT" 'KT (Реле времени)
    
                        Case "KV" 'KV (Реле напряжения)
    
                        Case "KK" 'KK (Реле тепловое)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "HA" 'HA (Звонок)
    
                        Case "UG" 'UG (Блок питания)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "TV" 'TV (Трансформатор)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "UZ" 'UZ (Твердотельное реле)
                            shpElementOnVID.Cells("Prop.Polusov").Formula = IIf(shpElementOnCxema.NameU Like SymName & "3P*", 3, 1)
                        Case "UZF" 'UZF (ИБП, Стабилизатор)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "UF" 'UF (Частотник)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
                        Case "XS" 'XS (Розетка)
                            shpElementOnVID.Cells("Prop.Tok").Formula = AdrParent + "!Prop.Tok"
'                        Case "DD" 'DD (ТРМ, ПЛК-моноблок)
'                            shpElementOnVID.Cells("Prop.TPM").Formula = AdrParent + "!Prop.Model"
                        Case Else
                    End Select
                End If
            Case typePLCParent
                On Error Resume Next
                Set shpElementOnVID = vsoPageVID.Drop(VIDvss.Masters.Item("DD"), DropX, DropY)
'                If err.Number <> 0 Then
                    err.Clear
                    On Error GoTo 0
''                    MsgBox "Элемент схемы " & shpElementOnCxema.NameU & " не имеет чертёж внешнего вида", vbExclamation + vbOKOnly, "САПР-АСУ: отсутствует чертеж внешнего вида элемента схемы"
'                Else
                    shpElementOnVID.Cells("User.NameParent").Formula = AdrParent + "!User.Name"
                    shpElementOnVID.Cells("User.Name").Formula = AdrParent + "!Prop.SymName&" + AdrParent + "!Prop.Number"
                    shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + shpElementOnCxema.ContainingPage.NameU + "/" + shpElementOnCxema.NameID + """"
                    shpElementOnVID.CellsSRC(visSectionHyperlink, 0, visHLinkExtraInfo).FormulaU = GUIDParent
                    shpElementOnVID.Shapes("Desc").text = shpElementOnCxema.Shapes("Desc").text
                    shpElementOnVID.Cells("Prop.ShowDesc").Formula = 1
                    dX = shpElementOnVID.Cells("Width").Result(0)
                    dY = IIf(shpElementOnVID.Cells("Height").Result(0) > dY, shpElementOnVID.Cells("Height").Result(0), dY)
                    Select Case SymName
                        Case "DD" 'DD (ПЛК-модульный)
                            shpElementOnVID.Cells("Prop.TPM").Formula = """" & shpElementOnCxema.Cells("Prop.Model").ResultStr(0) & """" 'AdrParent + "!Prop.TPM"
                        Case Else
                    End Select
'                End If
            Case Else
                dX = 0
        End Select
        DropX = DropX + dX * 2
        n = n + 1
        If n = ElementovVStroke Then
            DropY = DropY - dY * 2
            DropX = 0
            dY = 0
            n = 0
        End If
    Next
    If colElementToVID.Count > 0 Then MsgBox "Добавлено " & colElementToVID.Count & " аппаратов из схемы", vbInformation + vbOKOnly, "САПР-АСУ: аппараты из схемы добавлены"
End Sub
