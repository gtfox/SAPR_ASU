'------------------------------------------------------------------------------------------------------------
' Module        : KabeliSVP - Кабели на эл. схеме, на планах и на схеме внешних проводок (СВП)
' Author        : gtfox
' Date          : 2020.09.21
' Description   : Вставка и нумерация кабелей на эл. схеме, на планах и автосоздание схемы внешних проводок (СВП)
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Public PastePoint As Variant 'Точка вставки следующего датчика
Const KonecLista As Double = 10 / 25.4 'Расстояние от правого края листа, за которое не дожны заходить фигуры


Public Sub AddCableOnSensor(shpSensor As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddCableOnSensor - Вставляет кабель для подключенного датчика/привода на эл.схеме
                'Вставляется шейп кабеля для подключенного датчика/привода на эл.схеме
                'группируется с подключенными проводами, нумеруется, связываются ссылками друг на друга
                'Если датчик многокабельный(MultiCable=true), то кабели ссылаются не на датчик, а на конкретные входы в датчике
'------------------------------------------------------------------------------------------------------------
    Dim shpKabel As Visio.Shape
    Dim shpSensorIO As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim colWires As Collection
    Dim colWiresIO As Collection
    Dim vsoMaster As Visio.Master
    Dim MultiCable As Boolean '1 вход = 1 кабель
    Dim NomerShemy As Integer
    Dim PinX As Double
    Dim PinY As Double
    
    PinX = shpSensor.Cells("PinX").Result(0)
    PinY = shpSensor.Cells("PinY").Result(0)
    
    NomerShemy = shpSensor.ContainingPage.PageSheet.Cells("User.NomerShemy").Result(0)
    
    MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)
    Set colWires = New Collection
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_SVP.vss").Masters.Item("Kabel")

    If MultiCable Then
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If shpSensorIO.Name Like "SensorIO*" Then
                'Вставляем шейп кабеля
                Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensorIO.Cells("PinX").Result(0) + PinX, shpSensorIO.Cells("PinY").Result(0) + PinY + 0.196850393700787)
                'Находим подключенные провода и суем их в коллекцию
                Set colWires = FillColWires(shpSensorIO)
                'Добавляем подключенные провода в группу с кабелем
                AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
                'Число проводов в кабеле
                shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
                'Сохраняем к какому шкафу подключен кабель
                If NomerShemy = 0 Then 'если на листе несколько шкафов то...
                    'Определяем к какому шкафу/коробке принадлежит клеммник
                    '-------------Пока не реализовано----------------------
                Else
                    shpKabel.Cells("User.LinkToBox").Formula = NomerShemy
                End If
'                'Кабели ссылаются не на датчик, а на конкретные входы в датчике
'                shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
'                'Связываем входы с кабелями
'                shpSensorIO.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
            End If
        Next
    Else
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If shpSensorIO.Name Like "SensorIO*" Then
                'Находим подключенные провода на конкретном IO и суем их в коллекцию
                Set colWiresIO = FillColWires(shpSensorIO)
                'Добавляем провода с конкретного входа в общую колекцию проводов датчика
                For Each vsoShape In colWiresIO
                    colWires.Add vsoShape
                Next
            End If
        Next
        'Вставляем шейп кабеля
        Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensor.Cells("PinX").Result(0), shpSensor.Cells("PinY").Result(0) + 0.19685)
        'Добавляем подключенные провода в группу с кабелем
        AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
        'Число проводов в кабеле
        shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
        'Сохраняем к какому шкафу подключен кабель
        If NomerShemy = 0 Then 'если на листе несколько шкафов то...
            'Определяем к какому шкафу/коробке принадлежит клеммник
            '-------------Пока не реализовано----------------------
        Else
            shpKabel.Cells("User.LinkToBox").Formula = NomerShemy
        End If
'        'Кабель ссылается не на датчик, а на конкретный вход в датчике
'        shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
'        'Связываем вход с кабелем
'        shpSensorIO.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
    End If

End Sub

Sub AddToGroupCable(shpKabel As Visio.Shape, vsoPage As Visio.Page, colWires As Collection)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddToGroupCable - Добавляем подключенные провода в группу с кабелем
'------------------------------------------------------------------------------------------------------------
    Dim vsoSelection As Visio.Selection
    Dim vsoActivePage As Visio.Page
    Dim shpWire As Visio.Shape
    
    'Добавляем подключенные провода в группу с кабелем
    Set vsoSelection = ActiveWindow.Selection
    ActiveWindow.Page = vsoPage 'ActiveDocument.Pages(i)' активация нужной страницы
    Set vsoActivePage = ActiveWindow.Page
    With vsoSelection
        .DeselectAll
        .Select shpKabel, visSelect
        For Each shpWire In colWires
            'Чистим провода
            shpWire.Cells("Prop.Number").FormulaU = ""
            shpWire.Cells("Prop.SymName").FormulaU = ""
            shpWire.Cells("User.AdrSource").FormulaU = ""
            shpWire.Cells("Prop.AutoNum").FormulaU = False
            shpWire.Cells("Prop.HideNumber").FormulaU = True
            shpWire.Cells("Prop.HideName").FormulaU = True
            .Select shpWire, visSelect
        Next
        .AddToGroup
        'Сдвигаем вверх
        .DeselectAll
        .Select shpKabel, visSelect
        .Move 0#, 0.19685
    End With
End Sub

Sub DeleteCableSH(shpKabel As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteCableSH - Чистим ссылку в подключенном датчике перед удалением кабеля, и удаляем кабель
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
'    Dim shpSensorIO As Visio.Shape
'
'    'Находим датчик по ссылке в кабеле
'    Set shpSensorIO = ShapeByHyperLink(shpKabel.Cells("User.LinkToSensor").ResultStr(0))
'    'Чистим ссылку на кабель в датчике
'    On Error Resume Next
'    shpSensorIO.Cells("User.LinkToCable").FormulaU = ""
    
End Sub

Sub AddSensorOnSVP(shpSensor As Visio.Shape, vsoPageSVP As Visio.Page, ShinaNumber As Boolean)
    
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
    Const SVPWireL As Double = 10 / 25.4 'Длина проводов торчащих из шины на СВП
    Const Interval As Double = 5 / 25.4 'Расстояние между датчиками на СВП
    Const Klemma As Double = 240 / 25.4 'Высота расположения клеммника шкафа на СВП
    Const Datchik As Double = 100 / 25.4 'Высота расположения датчика на СВП
    
    Set colCables = New Collection
    Set colWires = New Collection
    Set colTerms = New Collection
    Set colCablesOnElSh = New Collection
    
    ActiveWindow.Page = ActiveDocument.Pages(shpSensor.ContainingPage.Name)
    
    Set vsoSelection = ActiveWindow.Selection
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_SVP.vss").Masters.Item("KabelSVP")
    
    MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)
    
    If MultiCable Then
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If shpSensorIO.Name Like "SensorIO*" Then
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
                colCablesOnElSh.Add colWires.Item(1).Parent, CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0))
            End If
        Next
        vsoSelection.Select shpSensor, visSelect 'Датчик
    Else
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If shpSensorIO.Name Like "SensorIO*" Then
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
        colCablesOnElSh.Add colWires.Item(1).Parent, CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0))
    End If

    'Копируем что насобирали
    vsoSelection.Copy
    'Отключаем события автоматизации (чтобы не перенумеровалось все)
    Application.EventsEnabled = 0

    ActiveWindow.Page = ActiveDocument.Pages(vsoPageSVP.Name)
    ActivePage.Paste
    'Application.ActiveDocument.Pages("СВП").Paste

    Set vsoGroup = ActiveWindow.Selection.Group
    
    Set colCables = New Collection
    vsoGroup.Cells("PinX").Formula = "(" & PastePoint & "+" & Interval & "+" & vsoGroup.Cells("LocPinX").Result(0) & ")/ThePage!PageScale*ThePage!DrawingScale"
    vsoGroup.Cells("PinY").Formula = Klemma & "-" & vsoGroup.Cells("LocPinY").Result(0)
    
    'Анализируем что вставили
    For Each vsoShape In vsoGroup.Shapes
         Select Case vsoShape.Cells("User.SAType").Result(0)
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
    
    'Ставим на место датчик
    shpSensorSVP.Cells("PinY").Formula = Datchik
    DoEvents 'На-я тут этот DoEvents?
    
    For Each shpCable In colCables
        'В кабеле находим длину провода
        WireHeight = shpCable.Shapes(1).Cells("Height").Result(0)
        'Вставляем шейп кабеля СВП
        Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, Datchik + WireHeight - SVPWireL)
        shpKabelSVP.Cells("Width").Formula = WireHeight - SVPWireL * 2
        shpKabelSVP.Cells("Prop.Number").Formula = shpCable.Cells("Prop.Number").Result(0)
        shpKabelSVP.Cells("Prop.Marka").Formula = """" & shpCable.Cells("User.Marka").ResultStr(0) & """"
        shpKabelSVP.Cells("Prop.WireCount").Formula = shpCable.Shapes.Count
        'По номеру кабеля СВП находим шейп кабеля на эл.сх.
        Set vsoShape = colCablesOnElSh.Item(CStr(shpCable.Cells("Prop.Number").Result(0)))
        'Заполняем длину кабеля из эл.схемы (длина кабеля эл.схемы заполняется из плана)
        shpKabelSVP.Cells("Prop.Dlina").FormulaU = "Pages[" + vsoShape.ContainingPage.NameU + "]!" + vsoShape.NameID + "!Prop.Dlina"
        
        WireNumber = 0
        'Ищем вход в датчике соединенный с текущим кабелем
        For Each shpWire In shpCable.Shapes
            For i = 1 To shpWire.Connects.Count
                If shpWire.Connects(i).ToSheet.Name Like "Term*" Then
                    Set cellKlemmaShkafa = shpWire.Connects(i).ToCell
                    NumberKlemmaShkafa = shpWire.Connects(i).ToSheet.Cells("Prop.Number").Result(0)
                ElseIf shpWire.Connects(i).ToSheet.Name Like "PLCTerm*" Then
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
                MsgBox "В кабеле больше " & WireNumber & " проводов", vbOKOnly + vbCritical, "Info"
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
    Dim shpPLCTerm As Visio.Shape
    
    Set colWires = New Collection
    For Each shpPLCTerm In shpSensorIO.Shapes
        If shpPLCTerm.Name Like "PLCTerm*" Then
            If shpPLCTerm.FromConnects.Count = 1 Then
                If shpPLCTerm.FromConnects.FromSheet.Name Like "w*" Then
                    colWires.Add shpPLCTerm.FromConnects.FromSheet
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
        If shpWire.Name Like "w*" Then
            If shpWire.Connects.Count = 2 Then
                For i = 1 To shpWire.Connects.Count
                    If shpWire.Connects(i).ToSheet.Name Like "Term*" Then
                        colTerms.Add shpWire.Connects(i).ToSheet
                    End If
                Next
            End If
        End If
    Next
    Set FillColTerms = colTerms
End Function

Public Sub AddPagesSVP()
'------------------------------------------------------------------------------------------------------------
' Macros        : AddPagesSVP - Создает листы СВП
                'Заполняет листы СВП датчиками, отсортированными по возрастанию их координаты Х на эл. схеме
'------------------------------------------------------------------------------------------------------------
    Dim NomerShemy As Integer
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
    Dim ShinaNumber As Boolean 'Нумерация проводов кабеля по типу ШИНЫ(Номер=Клемме), или Номер провода кабеля = Порядковому нореру
    Dim ss As String
    Dim i As Integer, ii As Integer, j As Integer, N As Integer
    
    ShinaNumber = False
    
    PastePoint = "25 mm - TheDoc!User.OffsetFrame"
    
    Set ThePage = ActivePage.PageSheet
    
    Set colShpDoc = New Collection
    
    PageName = "Схема"  'Имена листов где возможна нумерация
    'If ThePage.CellExists("User.NomerShemy", 0) Then NomerShemy = ThePage.Cells("User.NomerShemy").Result(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер.
    NomerShemy = 4

    'Цикл поиска датчиков и приводов
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If InStr(1, vsoPage.Name, PageName) > 0 Then    'Берем те, что содержат "Схема" в имени
            If vsoPage.PageSheet.Cells("User.NomerShemy").Result(0) = NomerShemy Then    'Берем все схемы с номером той, на которую вставляем элемент
                Set colShpPage = New Collection
                For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                    If vsoShapeOnPage.CellExists("User.SAType", 0) Then   'Если в шейпе есть тип, то -
                        Select Case vsoShapeOnPage.Cells("User.SAType").Result(0)
                            Case typeSensor, typeActuator
                                'Собираем в коллекцию нужные для сортировки шейпы
                                colShpPage.Add vsoShapeOnPage
                        End Select
                    End If
                Next
                
                'Сортируем то что нашли на листе
                
                'из коллекции передаем в массив для сортировки
                If colShpPage.Count > 0 Then
                    ReDim shpMas(colShpPage.Count - 1)
                Else
                    Exit Sub
                End If
                
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
ExitWhileX:              Set shpMas(i) = shpTemp
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

    'Берем первую страницу СВП
    Set vsoPage = SAPageExist("СВП") 'ActiveDocument.Pages("СВП")
    If vsoPage Is Nothing Then Set vsoPage = AddSAPage("СВП")
    
    'Вставляем на лист СВП найденные и отсортированные датчики/приводы
    For i = 1 To colShpDoc.Count
        AddSensorOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
        'Если лист кончился
        If PastePoint > vsoPage.PageSheet.Cells("PageWidth").Result(0) - KonecLista Then
            'Положение текущей страницы
            Index = vsoPage.Index
            'Создаем новую страницу СВП
            Set vsoPage = AddSAPage("СВП")
            'Положение новой страницы сразу за текущей
            vsoPage.Index = Index + 1
            PastePoint = "25 mm - TheDoc!User.OffsetFrame"
            'Вставляем этот же датчик только на следующем листе
            AddSensorOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
        End If
    Next
    ActiveWindow.DeselectAll
End Sub