'------------------------------------------------------------------------------------------------------------
' Module        : KabeliSVP - Кабели на схеме внешних проводок (СВП)
' Author        : gtfox
' Date          : 2020.09.21
' Description   : Автосоздание схемы внешних проводок (СВП)
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Public PastePoint As Variant 'Точка вставки следующего датчика

Const NachaloVstavki As Double = 20 / 25.4 'Расстояние + Interval(5mm) от левого края листа куда вставляется первый датчик
Const SVPWireL As Double = 10 / 25.4 'Длина проводов торчащих из шины на СВП
Const Interval As Double = 5 / 25.4 'Расстояние между датчиками на СВП
Const Klemma As Double = 240 / 25.4 'Высота расположения клеммника шкафа на СВП
Const Datchik As Double = 97.5 / 25.4 'Высота расположения датчика на СВП
Const KonecLista As Double = 10 / 25.4 'Расстояние от правого края листа, за которое не дожны заходить фигуры

Sub AddSensorOnSVP(shpSensor As Visio.Shape, vsoPageSVP As Visio.Page, ShinaNumber As Boolean)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddSensorOnSVP - Добавляет датчик, провод, клеммы на СВП
                'После вставки провода и кабель эл. схемы заменяются шейпом кабеля СВП, жилы приклеиваются к клеммам
'------------------------------------------------------------------------------------------------------------
    Dim shpSensorIO As Visio.Shape
    Dim shpTerm As Visio.Shape
    Dim shpCable As Visio.Shape
    Dim shpWire As Visio.Shape
    Dim colCablesOnElSh As Collection
    Dim colCables As Collection
    Dim colWires As Collection
    Dim colTerms As Collection
    Dim colWiresIO As Collection
    Dim vsoSelection As Visio.Selection
    Dim vsoMaster As Visio.Master
    Dim shpKabelSVP As Visio.Shape
    Dim vsoGroup As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim shpSensorSVP As Visio.Shape
    Dim MultiCable As Boolean
    Dim cellKlemmaShkafa As Visio.Cell
    Dim cellKlemmaDatchika As Visio.Cell
    Dim cellWireDown As Visio.Cell
    Dim cellWireUp As Visio.Cell
    Dim NumberKlemmaShkafa As Integer
    Dim NumberKlemmaDatchika As Integer
    Dim WireNumber As Integer
    Dim i As Integer
    Dim UserType As Integer
    Dim WireHeight As Double

    
    Set colCables = New Collection
    Set colWires = New Collection
    Set colTerms = New Collection
    Set colCablesOnElSh = New Collection
    
    ActiveWindow.Page = ActiveDocument.Pages(shpSensor.ContainingPage.name)
    
    Set vsoSelection = ActiveWindow.Selection
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_SVP.vss").Masters.Item("KabelSVP")
    
    MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)

    If MultiCable Then
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeSensorIO) Then
                'Находим подключенные провода и суем их в коллекцию
                Set colWires = FillColWires(shpSensorIO)
                'Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
                Set colTerms = FillColTerms(colWires)
                'Выделяем всех
                vsoSelection.Select colWires.Item(1).Parent, visSelect 'Кабель
                For Each shpTerm In colTerms
                    vsoSelection.Select shpTerm, visSelect 'Клеммы шкафа
                Next
                'Сохраняем кабели с эл.сх. чтобы получить от них по ссылке длину кабеля
                colCablesOnElSh.Add colWires.Item(1).Parent, IIf(colWires.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), shpCable.Cells("Prop.SymName").ResultStr(0) & colWires.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0)))
            End If
        Next
        vsoSelection.Select shpSensor, visSelect 'Датчик
    Else
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeSensorIO) Then
                'Находим подключенные провода на конкретном IO и суем их в коллекцию
                Set colWiresIO = FillColWires(shpSensorIO)
                'Добавляем провода с конкретного входа в общую колекцию проводов датчика
                For Each vsoShape In colWiresIO
                    colWires.Add vsoShape
                Next
            End If
        Next
        'Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
        Set colTerms = FillColTerms(colWires)
        'Выделяем всех
        vsoSelection.Select shpSensor, visSelect 'Датчик
        vsoSelection.Select colWires.Item(1).Parent, visSelect 'Кабель
        For Each shpTerm In colTerms
            vsoSelection.Select shpTerm, visSelect 'Клеммы шкафа
        Next
        'Сохраняем кабели с эл.сх. чтобы получить от них по ссылке длину кабеля
        colCablesOnElSh.Add colWires.Item(1).Parent, IIf(colWires.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), colWires.Item(1).Parent.Cells("Prop.SymName").ResultStr(0) & colWires.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0)))
    End If
    ActiveWindow.Selection = vsoSelection
    Set vsoGroup = ActiveWindow.Selection.Group
    
    'Чистим события перед копированием
    For Each vsoShape In ActiveWindow.Selection.PrimaryItem.Shapes
        With vsoShape
            .Cells("Prop.AutoNum").Formula = 0
            .Cells("EventMultiDrop").Formula = """"""
            .Cells("EventDrop").Formula = """"""
            .Cells("EventDblClick").Formula = """"""
        End With
    Next
    
    'Копируем что насобирали
    ActiveWindow.Selection.Copy

    'Восстанавливаем события после копирования
    For Each vsoShape In ActiveWindow.Selection.PrimaryItem.Shapes
        With vsoShape
            .Cells("Prop.AutoNum").Formula = 1
            .Cells("EventMultiDrop").Formula = "CALLTHIS(""AutoNumber.AutoNum"")"
            If ShapeSATypeIs(vsoShape, typeSensor) Or ShapeSATypeIs(vsoShape, typeActuator) Then
                .Cells("EventDrop").FormulaU = "CALLTHIS(""ThisDocument.EventDropAutoNum"")+SETF(GetRef(PinY),""80 mm/ThePage!PageScale*ThePage!DrawingScale"")+SETF(GetRef(Prop.ShowDesc),""true"")"
                .Cells("EventDblClick").Formula = "CALLTHIS(""CrossReferenceSensor.AddReferenceSensorFrm"")"
            Else
                .Cells("EventDrop").Formula = "CALLTHIS(""ThisDocument.EventDropAutoNum"")"
                .Cells("EventDblClick").Formula = "DOCMD(1312)"
            End If
        End With
    Next

    ActiveWindow.Selection.Ungroup

    ActiveWindow.Page = ActiveDocument.Pages(vsoPageSVP.name)
    'Отключаем события автоматизации (чтобы не перенумеровалось все)
    Application.EventsEnabled = 0
    
    ActivePage.Paste
    'Application.ActiveDocument.Pages(cListNameSVP).Paste

    Set vsoGroup = ActiveWindow.Selection.PrimaryItem

    'Отключаем меню
    For Each vsoShape In vsoGroup.Shapes
        With vsoShape
'            .Cells("Prop.AutoNum").Formula = 0
'            .Cells("EventMultiDrop").Formula = ""
'            .Cells("EventDrop").Formula = ""
'            .Cells("EventDblClick").Formula = ""
            .Cells("Actions.AddDB.Invisible").Formula = 1
            If ShapeSATypeIs(vsoShape, typeSensor) Or ShapeSATypeIs(vsoShape, typeActuator) Then
                .Cells("Actions.Celyj.Invisible").Formula = 1
                .Cells("Actions.Nachalo.Invisible").Formula = 1
                .Cells("Actions.Seredina.Invisible").Formula = 1
                .Cells("Actions.Konec.Invisible").Formula = 1
                .Cells("Actions.Tune.Invisible").Formula = 1
                .Cells("Actions.ShowDesc.Invisible").Formula = 1
                .Cells("Actions.AddReference.Invisible").Formula = 1
                .Cells("Actions.KlemmyProvoda.Invisible").Formula = 1
                .Cells("Actions.KabeliIzProvodov.Invisible").Formula = 1
                .Cells("Actions.KabeliSrazu.Invisible").Formula = 1
            End If
        End With
    Next

    Set colCables = New Collection
    vsoGroup.Cells("PinX").Formula = "(" & PastePoint & "+" & Interval & "+" & vsoGroup.Cells("LocPinX").Result(0) & ")/ThePage!PageScale*ThePage!DrawingScale"
    vsoGroup.Cells("PinY").Formula = Klemma & "-" & vsoGroup.Cells("LocPinY").Result(0)
    
    'Анализируем что вставили
    For Each vsoShape In vsoGroup.Shapes
         Select Case ShapeSAType(vsoShape)
            Case typeSensor, typeActuator
                Set shpSensorSVP = vsoShape
            Case typeCableSH
                colCables.Add vsoShape
         End Select
    Next
    
    'Сохраняем точку вставки следующего датчика
    PastePoint = vsoGroup.Cells("PinX").Result(0) + vsoGroup.Cells("LocPinX").Result(0)
    'Если датчик вылез за границы листа, то удаляем его и выходим из макроса
    If PastePoint > vsoGroup.ContainingPage.PageSheet.Cells("PageWidth").Result(0) - KonecLista Then
        vsoGroup.Delete
        Exit Sub
    End If

    'Разгруппировываем
    vsoGroup.Ungroup
    
    'Двигаем тексты под датчик
    shpSensorSVP.Cells("Controls.DescPos.Y").Formula = "Height*-2"
    shpSensorSVP.Cells("Controls.TextPos").Formula = "Width*0.5"
    shpSensorSVP.Cells("Controls.TextPos.Y").Formula = "Height*-0.2"
    shpSensorSVP.Cells("Controls.FSAPos").Formula = "Width*0.5"
    shpSensorSVP.Cells("Controls.FSAPos.Y").Formula = "Height*-0.5"
    shpSensorSVP.Cells("Controls.NamePos").Formula = "Width*0.5"
    shpSensorSVP.Cells("Controls.NamePos.Y").Formula = "Height*-0.4"
    shpSensorSVP.CellsSRC(visSectionObject, visRowTextXForm, visXFormLocPinX).FormulaU = "TxtWidth * 0.5"
    shpSensorSVP.Shapes("FSA").CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0.5"
    shpSensorSVP.Shapes("Name").CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0.5"
    
    'Ставим на место датчик
    shpSensorSVP.Cells("PinY").Formula = Datchik
    DoEvents 'На*уя тут этот DoEvents?
    
    For Each shpCable In colCables
        'В кабеле находим длину провода
        WireHeight = shpCable.Shapes(1).Cells("Height").Result(0)
        'Вставляем шейп кабеля СВП
        Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, Datchik + WireHeight - SVPWireL)
        shpKabelSVP.Cells("Width").Formula = WireHeight - SVPWireL * 2
        shpKabelSVP.Cells("Prop.Number").Formula = """" & IIf(shpCable.Cells("Prop.BukvOboz").Result(0), shpCable.Cells("Prop.SymName").ResultStr(0) & shpCable.Cells("Prop.Number").Result(0), CStr(shpCable.Cells("Prop.Number").Result(0))) & """"
        
        shpKabelSVP.Cells("Prop.WireCount").Formula = shpCable.Shapes.Count
        'По номеру кабеля СВП находим шейп кабеля на эл.сх.
        Set vsoShape = colCablesOnElSh.Item(IIf(shpCable.Cells("Prop.BukvOboz").Result(0), shpCable.Cells("Prop.SymName").ResultStr(0) & shpCable.Cells("Prop.Number").Result(0), CStr(shpCable.Cells("Prop.Number").Result(0))))
        'Заполняем длину кабеля из эл.схемы (длина кабеля эл.схемы заполняется из плана)
        shpKabelSVP.Cells("Prop.Dlina").FormulaU = "Pages[" + vsoShape.ContainingPage.NameU + "]!" + vsoShape.NameID + "!Prop.Dlina"
        shpKabelSVP.Cells("Prop.Marka").Formula = "Pages[" + vsoShape.ContainingPage.NameU + "]!" + vsoShape.NameID + "!User.Marka" '"""" & shpCable.Cells("User.Marka").ResultStr(0) & """"
        WireNumber = 0
        'Ищем вход в датчике соединенный с текущим кабелем
        For Each shpWire In shpCable.Shapes
            For i = 1 To shpWire.Connects.Count
                If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeTerm) Then
                    Set cellKlemmaShkafa = shpWire.Connects(i).ToCell
                    NumberKlemmaShkafa = shpWire.Connects(i).ToSheet.Cells("Prop.Number").Result(0)
                ElseIf ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeSensorTerm) Then
                    Set cellKlemmaDatchika = shpWire.Connects(i).ToCell
                    NumberKlemmaDatchika = shpWire.Connects(i).ToSheet.Cells("User.Number").Result(0)
                End If
            Next
            If WireNumber < 14 Then
                WireNumber = WireNumber + 1
                Set cellWireDown = shpKabelSVP.Cells("Controls.W" & WireNumber & "1")
                Set cellWireUp = shpKabelSVP.Cells("Controls.W" & WireNumber & "2")
                'Клеим провод
                cellWireDown.GlueTo cellKlemmaDatchika
                shpKabelSVP.Cells("Prop.WIRE" & WireNumber & "1").Formula = IIf(ShinaNumber, NumberKlemmaShkafa, WireNumber)
                cellWireUp.GlueTo cellKlemmaShkafa
                shpKabelSVP.Cells("Prop.WIRE" & WireNumber & "2").Formula = IIf(ShinaNumber, NumberKlemmaDatchika, WireNumber)
            Else
                MsgBox "В кабеле больше " & WireNumber & " проводов", vbOKOnly + vbCritical, "САПР-АСУ: Info"
                Exit For
            End If
            
        Next
        'Круг по середине кабеля
        shpKabelSVP.Cells("Controls.BendPnt").Formula = shpKabelSVP.Cells("Width").Result(0) * 0.5
    Next
    
    'Удаляем кабели эл. схемы
    For Each shpCable In colCables
        shpCable.Delete
    Next

    'Включаем события автоматизации
    Application.EventsEnabled = -1

End Sub

Function FillColWires(shpSensorIO As Visio.Shape) As Collection
'------------------------------------------------------------------------------------------------------------
' Function        : FillColWires - Находим подключенные провода и суем их в коллекцию
'------------------------------------------------------------------------------------------------------------
    Dim colWires As Collection
    Dim shpSensorTerm As Visio.Shape
    
    Set colWires = New Collection
    For Each shpSensorTerm In shpSensorIO.Shapes
        If ShapeSATypeIs(shpSensorTerm, typeSensorTerm) Then
            If shpSensorTerm.FromConnects.Count = 1 Then
                If ShapeSATypeIs(shpSensorTerm.FromConnects.FromSheet, typeWire) Then
                    colWires.Add shpSensorTerm.FromConnects.FromSheet
                End If
            End If
        End If
    Next
    Set FillColWires = colWires
End Function

Function FillColTerms(colWires As Collection) As Collection
'------------------------------------------------------------------------------------------------------------
' Function        : FillColTerms - Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
'------------------------------------------------------------------------------------------------------------
    Dim colTerms As Collection
    Dim shpWire As Visio.Shape
    
    Set colTerms = New Collection
    
    For Each shpWire In colWires
        If ShapeSATypeIs(shpWire, typeWire) Then
            If shpWire.Connects.Count = 2 Then
                For i = 1 To shpWire.Connects.Count
                    If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeTerm) Then
                        colTerms.Add shpWire.Connects(i).ToSheet
                    End If
                Next
            End If
        End If
    Next
    Set FillColTerms = colTerms
End Function

Function FindSensorFromKabel(shpKabel As Visio.Shape) As Visio.Shape
'------------------------------------------------------------------------------------------------------------
' Function        : FindSensorFromKabel - Находим датчик/привод подключенный кабелем
'------------------------------------------------------------------------------------------------------------
    Dim shpWire As Visio.Shape

    For Each shpWire In shpKabel.Shapes
        If ShapeSATypeIs(shpWire, typeWire) Then
            If shpWire.Connects.Count = 2 Then
                For i = 1 To shpWire.Connects.Count
                    If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeSensorTerm) Then
                        Set FindSensorFromKabel = shpWire.Connects(i).ToSheet.Parent.Parent
                        Exit Function
                    End If
                Next
            End If
        End If
    Next
End Function

Public Sub PageSVPAddKabeliFrm()
    Load frmPageSVPAddKabeli
    frmPageSVPAddKabeli.Show
End Sub

Public Sub AddPagesSVP(NazvanieShkafa As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddPagesSVP - Создает листы СВП
                'Заполняет листы СВП датчиками, отсортированными по возрастанию их координаты Х на эл. схеме
'------------------------------------------------------------------------------------------------------------
'    Dim NazvanieShkafa As String
    Dim ThePage As Visio.Shape
    Dim vsoShapeOnPage As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim shpElement As Shape
    Dim Prev As Shape
    Dim colShpPage As Collection
    Dim colShpDoc As Collection
    Dim shpMas() As Shape
    Dim shpTemp As Shape
    Dim Index As Integer
    Dim ShinaNumber As Boolean 'Нумерация проводов кабеля по типу ШИНЫ(Номер=Клемме), или Номер провода кабеля = Порядковому номеру жилы в кабеле
    Dim ss As String
    Dim i As Integer, ii As Integer, j As Integer, n As Integer
    
    ShinaNumber = 1
    
    PastePoint = NachaloVstavki
    
    Set ThePage = ActivePage.PageSheet
    
    Set colShpDoc = New Collection
    
    PageName = cListNameCxema
    'If ThePage.CellExists("Prop.SA_NazvanieShkafa", 0) Then NazvanieShkafa = ThePage.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер.
'    NazvanieShkafa = 4

    'Цикл поиска датчиков и приводов
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If vsoPage.name Like PageName & "*" Then    'Берем те, что содержат "Схема" в имени
            Set colShpPage = New Collection
            For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                If vsoShapeOnPage.CellExists("User.Shkaf", 0) Then
                    If vsoShapeOnPage.Cells("User.Shkaf").ResultStr(0) = NazvanieShkafa Then 'Берем все шкафы с именем того, на который вставляем элемент
                        Select Case ShapeSAType(vsoShapeOnPage) 'Если в шейпе есть тип, то -
                            Case typeSensor, typeActuator
                                'Собираем в коллекцию нужные для сортировки шейпы
                                colShpPage.Add vsoShapeOnPage
                            Case Else
                        End Select
                    End If
                End If
            Next
            
            'Сортируем то что нашли на листе
            
            'из коллекции передаем в массив для сортировки
            If colShpPage.Count > 0 Then
                ReDim shpMas(colShpPage.Count - 1)
                i = 0
                For Each shpElement In colShpPage
                    Set shpMas(i) = shpElement
                    i = i + 1
                Next
            
                ' "Сортировка вставками" массива шейпов по возрастанию коордонаты Х
                '--V--Сортируем по возрастанию коордонаты Х
                UbMas = UBound(shpMas)
                For j = 1 To UbMas
                    Set shpTemp = shpMas(j)
                    i = j
                    'If shpMas(i) Is Nothing Then Exit Sub
                    While shpMas(i - 1).Cells("PinX").Result("mm") > shpTemp.Cells("PinX").Result("mm") '>:возрастание, <:убывание
                        Set shpMas(i) = shpMas(i - 1)
                        i = i - 1
                        If i <= 0 Then GoTo ExitWhileX
                    Wend
ExitWhileX:                  Set shpMas(i) = shpTemp
                Next
                '--Х--Сортировка по возрастанию коордонаты Х
                
                'Собираем отсортированные листы в коллекцию документа
                For i = 0 To UbMas
                    colShpDoc.Add shpMas(i)
                Next
                Set colShpPage = Nothing
            End If
        End If
    Next

    If colShpDoc.Count > 0 Then
        'Берем первую страницу СВП
        Set vsoPage = GetSAPageExist(cListNameSVP) 'ActiveDocument.Pages(cListNameSVP)
        If vsoPage Is Nothing Then Set vsoPage = AddSAPage(cListNameSVP)
        SetPageSVP vsoPage
        'Вставляем на лист СВП найденные и отсортированные датчики/приводы
        For i = 1 To colShpDoc.Count
            AddSensorOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
            'Если лист кончился
            If PastePoint > vsoPage.PageSheet.Cells("PageWidth").Result(0) - KonecLista Then
                'Положение текущей страницы
                Index = vsoPage.Index
                'Создаем новую страницу СВП
                Set vsoPage = AddSAPage(cListNameSVP)
                'Положение новой страницы сразу за текущей
                vsoPage.Index = Index + 1
                PastePoint = NachaloVstavki
                'Вставляем этот же датчик только на следующем листе
                AddSensorOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
            End If
        Next
    End If

    ActiveWindow.DeselectAll
End Sub

Sub SetPageSVP(vsoPage As Visio.Page)
    Dim shpShkaf As Visio.Shape
    'Подвал
    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("PodvalCxemy"), 0, 0
    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).FormulaU = "(25 mm-TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale"
    'Шкаф
    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("ShkafMesto"), 0, 0
    Set shpShkaf = Application.ActiveWindow.Selection(1)
    With shpShkaf
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Formula = NachaloVstavki + Interval
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Klemma - 5 / 25.4
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "37.5 mm"
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = "382.5 mm"
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0"
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).FormulaU = "Height * 0"
        .CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.5 mm"
        .CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "1"
        .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "24 pt"
        .CellsSRC(visSectionObject, visRowTextXForm, visXFormLocPinX).FormulaU = "TxtWidth * 0.5"
        .Cells("Controls.TextPos").FormulaU = "Width * 0.5"
        .AddSection visSectionConnectionPts
        .AddRow visSectionConnectionPts, visRowLast, visTagDefault
        .CellsSRC(visSectionConnectionPts, 0, visCnnctX).FormulaForceU = "Width*1"
        .CellsSRC(visSectionConnectionPts, 0, visCnnctY).FormulaForceU = "Height*0"
        .CellsSRC(visSectionConnectionPts, 0, visCnnctDirX).FormulaForceU = "0 mm"
        .CellsSRC(visSectionConnectionPts, 0, visCnnctDirY).FormulaForceU = "0 mm"
        .CellsSRC(visSectionConnectionPts, 0, visCnnctType).FormulaForceU = "0 mm"
        .CellsSRC(visSectionConnectionPts, 0, visCnnctAutoGen).FormulaForceU = "0 mm"
        .CellsSRC(visSectionConnectionPts, 0, 6).FormulaForceU = ""
    End With
    'Клеммник
    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("klemmnik"), 0, 0
    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Formula = NachaloVstavki + Interval
    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Klemma - 5 / 25.4
    Application.ActiveWindow.Selection(1).Cells("Controls.Line").GlueTo shpShkaf.Cells("Connections.X1")
End Sub