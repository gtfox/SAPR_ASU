'------------------------------------------------------------------------------------------------------------
' Module        : KabeliSVP - Кабели на схеме внешних проводок (СВП)
' Author        : gtfox
' Date          : 2020.09.21/2023.03.29/2024.08.03
' Description   : Автосоздание схемы внешних проводок (СВП)
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Public PastePoint As Variant 'Точка вставки следующего датчика
Public colSensorsAdded As Collection

Const NachaloVstavki As Double = 20 / 25.4 'Расстояние + Interval(5mm) от левого края листа куда вставляется первый датчик
Const SVPWireL As Double = 10 / 25.4 'Длина проводов торчащих из шины на СВП
Const Interval As Double = 15 / 25.4 'Расстояние между датчиками на СВП 5
Const Klemma As Double = 240 / 25.4 'Высота расположения клеммника шкафа на СВП
Const Datchik15 As Double = 97.5 / 25.4 'Высота расположения датчика на СВП для малого штампа 15
Const KonecLista As Double = 10 / 25.4 'Расстояние от правого края листа, за которое не дожны заходить фигуры
Const Datchik55 As Double = 137.5 / 25.4 'Высота расположения датчика на СВП для большого штампа 55
Const KrugKabelya As Double = 0.6 'Круг по середине кабеля 0,5


Public Sub PageSVPAddKabeliFrm()
    Load frmPageSVPAddKabeli
    frmPageSVPAddKabeli.Show
End Sub

Public Sub AddPagesSVP(NazvanieShkafa As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddPagesSVP - Создает листы СВП
                'Заполняет листы СВП датчиками/приводами, клеммами и кабелями, отсортированными по возрастанию их координаты Х на эл. схеме
'------------------------------------------------------------------------------------------------------------
'    Dim NazvanieShkafa As String
    Dim ThePage As Visio.Shape
    Dim vsoShapeOnPage As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim shpElement As Visio.Shape
    Dim Prev As Visio.Shape
    Dim colShpPage As Collection
    Dim colShpDoc As Collection
    Dim shpMas() As Visio.Shape
    Dim shpTemp As Visio.Shape
    Dim shpShkafUp As Visio.Shape
    Dim Index As Integer
    Dim ShinaNumber As Boolean 'Нумерация проводов кабеля по типу ШИНЫ(Номер=Клемме на другом конце), или Номер провода кабеля = Порядковому номеру жилы в кабеле
    Dim ss As String
    Dim i As Integer, ii As Integer, j As Integer, n As Integer
    
    ShinaNumber = 0
    
    PastePoint = NachaloVstavki
    
    Set ThePage = ActivePage.PageSheet
    
    Set colShpDoc = New Collection
    
    Set colSensorsAdded = New Collection
    
    PageName = cListNameCxema

    'Цикл поиска кабелей
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If vsoPage.name Like PageName & "*" Then    'Берем те, что содержат "Схема" в имени
            Set colShpPage = New Collection
            For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                If vsoShapeOnPage.CellExists("User.LinkToBox", 0) Then
                    If GetNazvanie(vsoShapeOnPage.Cells("User.LinkToBox").ResultStr(0), 2) = NazvanieShkafa Then 'Берем все шкафы с именем того, на который вставляем элемент
                        Select Case ShapeSAType(vsoShapeOnPage) 'Если в шейпе есть тип, то -
                            Case typeCxemaCable
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
            
                ' "Сортировка вставками" массива шейпов по возрастанию номера кабеля
                '--V--Сортируем по возрастанию
                UbMas = UBound(shpMas)
                For j = 1 To UbMas
                    Set shpTemp = shpMas(j)
                    i = j
                    'If shpMas(i) Is Nothing Then Exit Sub
                    While shpMas(i - 1).Cells("Prop.Number").Result(0) > shpTemp.Cells("Prop.Number").Result(0) '>:возрастание, <:убывание
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
        Set vsoPage = ActivePage 'GetSAPageExist(cListNameSVP) 'ActiveDocument.Pages(cListNameSVP)
'        If vsoPage Is Nothing Then Set vsoPage = AddSAPage(cListNameSVP)
        Set shpShkafUp = GetShkafUp(vsoPage)
        'Вставляем на лист СВП найденные и отсортированные датчики/приводы
        For i = 1 To colShpDoc.Count
            AddCableOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
            'Если лист кончился
            If PastePoint > vsoPage.PageSheet.Cells("PageWidth").Result(0) - KonecLista Then
                'Положение текущей страницы
                Index = vsoPage.Index
                'Создаем новую страницу СВП
                Set vsoPage = AddSAPage(cListNameSVP)
                Set shpShkafUp = GetShkafUp(vsoPage)
                'Положение новой страницы сразу за текущей
                vsoPage.Index = Index + 1
                PastePoint = NachaloVstavki
                'Вставляем этот же датчик только на следующем листе
                AddCableOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
            End If
        Next
        shpShkafUp.Cells("Prop.SA_NazvanieShkafa").Formula = """" & GetNazvanie(colShpDoc.Item(1).Cells("User.LinkToBox.Prompt").ResultStr(0), 2) & """"
        shpShkafUp.Cells("Prop.SA_NazvanieMesta").Formula = """" & GetNazvanie(colShpDoc.Item(1).Cells("User.LinkToBox.Prompt").ResultStr(0), 3) & """"
    End If

    ResetLocalShkafMesto ActivePage
    ActiveWindow.DeselectAll
End Sub

Sub AddCableOnSVP(shpCable As Visio.Shape, vsoPageSVP As Visio.Page, ShinaNumber As Boolean)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddCableOnSVP - Добавляет датчик, провод, клеммы на СВП
                'После вставки провода и кабель эл. схемы заменяются шейпом кабеля СВП, жилы приклеиваются к клеммам
'------------------------------------------------------------------------------------------------------------
    Dim shpSensorIO As Visio.Shape
    Dim shpTerm As Visio.Shape
    Dim shpSensor As Visio.Shape
    Dim shpWire As Visio.Shape
    Dim shpShkafDown As Visio.Shape
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
    Dim BukvOboz As Boolean
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
    Dim bNacaloKlemmnika As Boolean

    Set colCables = New Collection
    Set colWires = New Collection
    Set colTerms = New Collection
    Set colCablesOnElSh = New Collection
    
    
    'Если кабелем подключён датчик/привод то используем AddSensorOnSVP
    Set shpSensor = FindSensorFromKabel(shpCable)
    If Not shpSensor Is Nothing Then
        On Error GoTo err1
        colSensorsAdded.Add shpSensor, CStr(shpSensor.Cells("Prop.SymName").ResultStr(0) & shpSensor.Cells("Prop.Number").Result(0))
        AddSensorOnSVP shpSensor, vsoPageSVP, ShinaNumber
err1:
        Exit Sub
    End If
    
    'На обоих концах кабеля находтся шкафы/коробки

    ActiveWindow.Page = ActiveDocument.Pages(shpCable.ContainingPage.name)
    
    Set vsoSelection = ActiveWindow.Selection
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("KabelSVP")
    
    BukvOboz = shpCable.Cells("Prop.BukvOboz").Result(0)
    
    'собираем провода и клеммы
    For Each shpWire In shpCable.Shapes
        If ShapeSATypeIs(shpWire, typeCxemaWire) Then
            If shpWire.Connects.Count = 2 Then
                colWires.Add shpWire, shpWire.NameID
                For i = 1 To shpWire.Connects.Count
                    If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeCxemaTerm) Then
                        colTerms.Add shpWire.Connects(i).ToSheet
                    End If
                Next
            End If
        End If
    Next
    
    'При соединении кабелем двух шкафов: Кто выше тот и шкаф :)
'    shpCable.Cells("User.LinkToBox.Prompt").Formula = """" & shpCable.Cells("User.LinkToBox").ResultStr(0) & """"
'    shpCable.Cells("User.LinkToSensor.Prompt").Formula = """" & shpCable.Cells("User.LinkToSensor").ResultStr(0) & """"
    If shpCable.Shapes(1).Connects(1).ToSheet.Cells("PinY").Result(0) > shpCable.Shapes(1).Connects(2).ToSheet.Cells("PinY").Result(0) Then
        shpCable.Cells("User.LinkToBox.Prompt").Formula = """" & shpCable.Shapes(1).Connects(1).ToSheet.Cells("User.FullName.Prompt").ResultStr(0) & """"
        shpCable.Cells("User.LinkToSensor.Prompt").Formula = """" & shpCable.Shapes(1).Connects(2).ToSheet.Cells("User.FullName.Prompt").ResultStr(0) & """"
    Else
        shpCable.Cells("User.LinkToBox.Prompt").Formula = """" & shpCable.Shapes(1).Connects(2).ToSheet.Cells("User.FullName.Prompt").ResultStr(0) & """"
        shpCable.Cells("User.LinkToSensor.Prompt").Formula = """" & shpCable.Shapes(1).Connects(1).ToSheet.Cells("User.FullName.Prompt").ResultStr(0) & """"
    End If

    'Выделяем всех
    vsoSelection.Select shpCable, visSelect 'Кабель
    For Each shpTerm In colTerms
        vsoSelection.Select shpTerm, visSelect 'Клеммы шкафа
        If shpTerm.Cells("Prop.Nachalo").Result(0) = 1 Then bNacaloKlemmnika = True
    Next
    'Сохраняем кабели с эл.сх. чтобы получить от них по ссылке длину кабеля
    colCablesOnElSh.Add shpCable, IIf(shpCable.Cells("Prop.BukvOboz").Result(0), shpCable.Cells("Prop.SymName").ResultStr(0) & shpCable.Cells("Prop.Number").Result(0), CStr(shpCable.Cells("Prop.Number").Result(0)))

    ActiveWindow.Selection = vsoSelection
    Set vsoGroup = ActiveWindow.Selection.Group
    
    'Чистим события перед копированием
    For Each vsoShape In ActiveWindow.Selection.PrimaryItem.Shapes
        With vsoShape
'            .Cells("Prop.AutoNum").Formula = 0
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
'            .Cells("Prop.AutoNum").Formula = 1
            .Cells("EventMultiDrop").Formula = "CALLTHIS(""AutoNumber.AutoNum"")"
            .Cells("EventDrop").Formula = "CALLTHIS(""ThisDocument.EventDropAutoNum"")"
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
        End With
    Next

    Set colCables = New Collection
    vsoGroup.Cells("PinX").Formula = "(" & PastePoint & "+" & IIf(bNacaloKlemmnika, Interval * 2, Interval) & "+" & vsoGroup.Cells("LocPinX").Result(0) & ")/ThePage!PageScale*ThePage!DrawingScale"
    vsoGroup.Cells("PinY").Formula = Klemma & "-" & vsoGroup.Cells("LocPinY").Result(0)
    bNacaloKlemmnika = False
    
    'Анализируем что вставили
    For Each vsoShape In vsoGroup.Shapes
         Select Case ShapeSAType(vsoShape)
            Case typeCxemaCable
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

    Set shpShkafDown = GetShkafDown(vsoGroup)
    shpShkafDown.Cells("Prop.SA_NazvanieShkafa").Formula = """" & GetNazvanie(shpCable.Cells("User.LinkToSensor.Prompt").ResultStr(0), 2) & """"
    shpShkafDown.Cells("Prop.SA_NazvanieMesta").Formula = """" & GetNazvanie(shpCable.Cells("User.LinkToSensor.Prompt").ResultStr(0), 3) & """"

    'Разгруппировываем
    vsoGroup.Ungroup

    'Ставим на место датчик
    For Each vsoShape In ActiveWindow.Selection
        If ShapeSATypeIs(vsoShape, typeCxemaTerm) Then
            If vsoShape.Cells("PinY").Result(0) < Klemma - vsoShape.Cells("LocPinY").Result(0) Then
                If ActivePage.Shapes("Рамка").Cells("User.Height").Result("mm") > 15 Then
                     vsoShape.Cells("PinY").Formula = Datchik55 - vsoShape.Cells("LocPinY").Result(0)
                Else
                     vsoShape.Cells("PinY").Formula = Datchik15 - vsoShape.Cells("LocPinY").Result(0)
                End If
                'vsoShape.Cells("PinY").Formula = IIf(ActivePage.Shapes("Рамка").Cells("Prop.CHAPTER").Result(0) = 0, Datchik55, Datchik15) - vsoShape.Cells("LocPinY").Result(0)
            End If
        End If
    Next

    DoEvents 'На*уя тут этот DoEvents?
    
    For Each shpCable In colCables
        'В кабеле находим длину провода
        WireHeight = shpCable.Shapes(1).Cells("Height").Result(0)
        'Вставляем шейп кабеля СВП
        If ActivePage.Shapes("Рамка").Cells("User.Height").Result("mm") > 15 Then
             Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, Datchik55 + WireHeight - SVPWireL)
        Else
             Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, Datchik15 + WireHeight - SVPWireL)
        End If
        'Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, IIf(ActivePage.Shapes("Рамка").Cells("Prop.CHAPTER").Result(0) = 0, Datchik55, Datchik15) + WireHeight - SVPWireL)
        shpKabelSVP.Cells("Width").Formula = WireHeight - SVPWireL * 2
        shpKabelSVP.Cells("Prop.Number").Formula = """" & IIf(shpCable.Cells("Prop.BukvOboz").Result(0), shpCable.Cells("Prop.SymName").ResultStr(0) & shpCable.Cells("Prop.Number").Result(0), CStr(shpCable.Cells("Prop.Number").Result(0))) & """"
        
        shpKabelSVP.Cells("Prop.WireCount").Formula = shpCable.Shapes.Count
        'По номеру кабеля СВП находим шейп кабеля на эл.сх.
        Set vsoShape = colCablesOnElSh.Item(IIf(shpCable.Cells("Prop.BukvOboz").Result(0), shpCable.Cells("Prop.SymName").ResultStr(0) & shpCable.Cells("Prop.Number").Result(0), CStr(shpCable.Cells("Prop.Number").Result(0))))
        'Заполняем длину кабеля из эл.схемы (длина кабеля эл.схемы заполняется из плана)
        shpKabelSVP.Cells("Prop.Dlina").FormulaU = "Pages[" + vsoShape.ContainingPage.NameU + "]!" + vsoShape.NameID + "!Prop.Dlina"
        shpKabelSVP.Cells("Prop.Marka").Formula = "Pages[" + vsoShape.ContainingPage.NameU + "]!" + vsoShape.NameID + "!User.Marka" '"""" & shpCable.Cells("User.Marka").ResultStr(0) & """"
        WireNumber = 0

        For Each shpWire In shpCable.Shapes
            'При соединении кабелем двух шкафов: Кто выше тот и шкаф :)
            If shpWire.Connects(1).ToSheet.Cells("PinY").Result(0) > shpWire.Connects(2).ToSheet.Cells("PinY").Result(0) Then
                Set cellKlemmaShkafa = shpWire.Connects(1).ToCell
                NumberKlemmaShkafa = shpWire.Connects(1).ToSheet.Cells("Prop.Number").Result(0)
                Set cellKlemmaDatchika = shpWire.Connects(2).ToCell
                NumberKlemmaDatchika = shpWire.Connects(2).ToSheet.Cells("Prop.Number").Result(0)
            Else
                Set cellKlemmaShkafa = shpWire.Connects(2).ToCell
                NumberKlemmaShkafa = shpWire.Connects(2).ToSheet.Cells("Prop.Number").Result(0)
                Set cellKlemmaDatchika = shpWire.Connects(1).ToCell
                NumberKlemmaDatchika = shpWire.Connects(1).ToSheet.Cells("Prop.Number").Result(0)
            End If

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
        shpKabelSVP.Cells("Controls.BendPnt").Formula = shpKabelSVP.Cells("Width").Result(0) * KrugKabelya
    Next
    
    'Удаляем кабели эл. схемы
    For Each shpCable In colCables
        shpCable.Delete
    Next

    'Включаем события автоматизации
    Application.EventsEnabled = -1

End Sub

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
    Dim bNacaloKlemmnika As Boolean


    Set colCables = New Collection
    Set colWires = New Collection
    Set colTerms = New Collection
    Set colCablesOnElSh = New Collection

    ActiveWindow.Page = ActiveDocument.Pages(shpSensor.ContainingPage.name)

    Set vsoSelection = ActiveWindow.Selection
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("KabelSVP")

    MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)

    shpSensor.Cells("User.Shkaf.Prompt").Formula = """" & shpSensor.Cells("User.Shkaf").ResultStr(0) & """"
'    shpSensor.Cells("User.Mesto.Prompt").Formula = """" & shpSensor.Cells("User.Mesto").ResultStr(0) & """"
    
    If MultiCable Then
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeCxemaSensorIO) Then
                'Находим подключенные провода и суем их в коллекцию
                Set colWires = FillColWires(shpSensorIO)
                'Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
                Set colTerms = FillColTerms(colWires)
                'Выделяем всех
                For Each shpTerm In colWires
                    vsoSelection.Select shpTerm.Parent, visSelect 'Кабель
                Next
                For Each shpTerm In colTerms
                    vsoSelection.Select shpTerm, visSelect 'Клеммы шкафа
                Next
                'Сохраняем кабели с эл.сх. чтобы получить от них по ссылке длину кабеля
'                colCablesOnElSh.Add colWires.Item(1).Parent, IIf(colWires.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), colWires.Item(1).Parent.Cells("Prop.SymName").ResultStr(0) & colWires.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0)))
            End If
        Next
        Set colCablesOnElSh = FillColCables(shpSensor)
        vsoSelection.Select shpSensor, visSelect 'Датчик
        
    Else
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeCxemaSensorIO) Then
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
            If shpTerm.Cells("Prop.Nachalo").Result(0) = 1 Then bNacaloKlemmnika = True
        Next
        'Сохраняем кабели с эл.сх. чтобы получить от них по ссылке длину кабеля
        colCablesOnElSh.Add colWires.Item(1).Parent, IIf(colWires.Item(1).Parent.Cells("Prop.BukvOboz").Result(0), colWires.Item(1).Parent.Cells("Prop.SymName").ResultStr(0) & colWires.Item(1).Parent.Cells("Prop.Number").Result(0), CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0)))
        colWires.Item(1).Parent.Cells("User.LinkToBox.Prompt").Formula = """" & colWires.Item(1).Parent.Cells("User.LinkToBox").ResultStr(0) & """"
        colWires.Item(1).Parent.Cells("User.LinkToSensor.Prompt").Formula = """" & colWires.Item(1).Parent.Cells("User.LinkToSensor").ResultStr(0) & """"
    End If
    ActiveWindow.Selection = vsoSelection
    Set vsoGroup = ActiveWindow.Selection.Group

    'Чистим события перед копированием
    For Each vsoShape In ActiveWindow.Selection.PrimaryItem.Shapes
        With vsoShape
'            .Cells("Prop.AutoNum").Formula = 0
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
'            .Cells("Prop.AutoNum").Formula = 1
            .Cells("EventMultiDrop").Formula = "CALLTHIS(""AutoNumber.AutoNum"")"
            If ShapeSATypeIs(vsoShape, typeCxemaSensor) Or ShapeSATypeIs(vsoShape, typeCxemaActuator) Then
                .Cells("EventDrop").FormulaU = "CALLTHIS(""ThisDocument.EventDropAutoNum"")+SETF(GetRef(PinY),""80 mm/ThePage!PageScale*ThePage!DrawingScale"")+SETF(GetRef(Prop.ShowDesc),""true"")"
                .Cells("EventDblClick").Formula = "CALLTHIS(""CrossReferenceSensor.AddReferenceSensorFrm"")"
            Else
                .Cells("EventDrop").Formula = "CALLTHIS(""ThisDocument.EventDropAutoNum"")"
'                .Cells("EventDblClick").Formula = "DOCMD(1312)"
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
            If ShapeSATypeIs(vsoShape, typeCxemaSensor) Or ShapeSATypeIs(vsoShape, typeCxemaActuator) Then
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
                .Cells("Prop.PerenosOboz").Formula = 0 'Отключаем перенос обозначений
            End If
        End With
    Next

    Set colCables = New Collection
    vsoGroup.Cells("PinX").Formula = "(" & PastePoint & "+" & IIf(bNacaloKlemmnika, Interval * 2, Interval) & "+" & vsoGroup.Cells("LocPinX").Result(0) & ")/ThePage!PageScale*ThePage!DrawingScale"
    vsoGroup.Cells("PinY").Formula = Klemma & "-" & vsoGroup.Cells("LocPinY").Result(0)
    bNacaloKlemmnika = False

    'Анализируем что вставили
    For Each vsoShape In vsoGroup.Shapes
         Select Case ShapeSAType(vsoShape)
            Case typeCxemaSensor, typeCxemaActuator
                Set shpSensorSVP = vsoShape
            Case typeCxemaCable
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
    shpSensorSVP.Cells("Controls.FSAPos.Y").Formula = "Height*-0.75"
    shpSensorSVP.Cells("Controls.NamePos").Formula = "Width*0.5"
    shpSensorSVP.Cells("Controls.NamePos.Y").Formula = "Height*-0.4"
    shpSensorSVP.CellsSRC(visSectionObject, visRowTextXForm, visXFormLocPinX).FormulaU = "TxtWidth * 0.5"
    shpSensorSVP.Shapes("FSA").CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0.5"
    shpSensorSVP.Shapes("Name").CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0.5"

    'Ставим на место датчик
    If ActivePage.Shapes("Рамка").Cells("User.Height").Result("mm") > 15 Then
         shpSensorSVP.Cells("PinY").Formula = Datchik55
    Else
         shpSensorSVP.Cells("PinY").Formula = Datchik15
    End If
    'shpSensorSVP.Cells("PinY").Formula = IIf(ActivePage.Shapes("Рамка").Cells("Prop.CHAPTER").Result(0) = 0, Datchik55, Datchik15)
    DoEvents 'На*уя тут этот DoEvents?

    For Each shpCable In colCables
        'В кабеле находим длину провода
        WireHeight = shpCable.Shapes(1).Cells("Height").Result(0)
        'Вставляем шейп кабеля СВП
        If ActivePage.Shapes("Рамка").Cells("User.Height").Result("mm") > 15 Then
             Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, Datchik55 + WireHeight - SVPWireL)
        Else
             Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, Datchik15 + WireHeight - SVPWireL)
        End If
        'Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, IIf(ActivePage.Shapes("Рамка").Cells("Prop.CHAPTER").Result(0) = 0, Datchik55, Datchik15) + WireHeight - SVPWireL)
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
                If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeCxemaTerm) Then
                    Set cellKlemmaShkafa = shpWire.Connects(i).ToCell
                    NumberKlemmaShkafa = shpWire.Connects(i).ToSheet.Cells("Prop.Number").Result(0)
                ElseIf ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeCxemaSensorTerm) Then
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
        shpKabelSVP.Cells("Controls.BendPnt").Formula = shpKabelSVP.Cells("Width").Result(0) * KrugKabelya
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
        If ShapeSATypeIs(shpSensorTerm, typeCxemaSensorTerm) Then
            If shpSensorTerm.FromConnects.Count = 1 Then
                If ShapeSATypeIs(shpSensorTerm.FromConnects.FromSheet, typeCxemaWire) Then
                    colWires.Add shpSensorTerm.FromConnects.FromSheet, shpSensorTerm.FromConnects.FromSheet.name
                End If
            End If
        End If
    Next
    Set FillColWires = colWires
End Function

Function FillColWiresOnPage(shpSensorIO As Visio.Shape) As Collection
'------------------------------------------------------------------------------------------------------------
' Function        : FillColWiresOnPage - Находим подключенные провода находящиеся на листе (не в группе кабеля) и суем их в коллекцию
'------------------------------------------------------------------------------------------------------------
    Dim colWires As Collection
    Dim shpSensorTerm As Visio.Shape
    
    Set colWires = New Collection
    For Each shpSensorTerm In shpSensorIO.Shapes
        If ShapeSATypeIs(shpSensorTerm, typeCxemaSensorTerm) Then
            If shpSensorTerm.FromConnects.Count = 1 Then
                If ShapeSATypeIs(shpSensorTerm.FromConnects.FromSheet, typeCxemaWire) Then
                    If Not shpSensorTerm.FromConnects.FromSheet.Parent.Type = visTypeGroup Then
                        colWires.Add shpSensorTerm.FromConnects.FromSheet, shpSensorTerm.FromConnects.FromSheet.name
                    End If
                End If
            End If
        End If
    Next
    Set FillColWiresOnPage = colWires
End Function

Function FillColTerms(colWires As Collection) As Collection
'------------------------------------------------------------------------------------------------------------
' Function        : FillColTerms - Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
'------------------------------------------------------------------------------------------------------------
    Dim colTerms As Collection
    Dim shpWire As Visio.Shape
    
    Set colTerms = New Collection
    
    For Each shpWire In colWires
        If ShapeSATypeIs(shpWire, typeCxemaWire) Then
            If shpWire.Connects.Count = 2 Then
                For i = 1 To shpWire.Connects.Count
                    If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeCxemaTerm) Then
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
        If ShapeSATypeIs(shpWire, typeCxemaWire) Then
            If shpWire.Connects.Count = 2 Then
                For i = 1 To shpWire.Connects.Count
                    If ShapeSATypeIs(shpWire.Connects(i).ToSheet, typeCxemaSensorTerm) Then
                        Set FindSensorFromKabel = shpWire.Connects(i).ToSheet.Parent.Parent
                        Exit Function
                    End If
                Next
                Set FindSensorFromKabel = Nothing
            End If
        End If
    Next
End Function

Function GetShkafUp(vsoPage As Visio.Page) As Visio.Shape
    Dim shpShkaf As Visio.Shape
    'Подвал
    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("PodvalCxemy"), 0, 0
    
    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).FormulaU = "(25 mm-TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale"
    If ActivePage.Shapes("Рамка").Cells("User.Height").Result("mm") > 15 Then
        Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).FormulaU = "(55 mm+1.15mm+TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale"
    Else
        Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).FormulaU = "(15 mm+1.15mm+TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale"
    End If

    
    'Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).FormulaForceU = IIf(ActivePage.Shapes("Рамка").Cells("Prop.CHAPTER").Result(0) = 0, "(55 mm+1.15mm+TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale", "(15 mm+1.15mm+TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale")

    'Шкаф
    vsoPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("ShkafMesto"), 0, 0
    Set shpShkaf = Application.ActiveWindow.Selection(1)
    With shpShkaf
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Formula = NachaloVstavki + Interval
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Klemma - 5 / 25.4
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "37.5 mm"
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = "(ThePage!PageWidth-TheDoc!User.SA_FR_OffsetFrame-PinX-5mm)/ThePage!PageScale*ThePage!DrawingScale"
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
    Set GetShkafUp = shpShkaf
End Function

Function GetShkafDown(vsoGroup As Visio.Shape) As Visio.Shape
    Dim shpShkaf As Visio.Shape
    'Шкаф
    vsoGroup.ContainingPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("ShkafMesto"), 0, 0
    Set shpShkaf = Application.ActiveWindow.Selection(1)
    With shpShkaf
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Formula = vsoGroup.Cells("PinX").Result(0) - vsoGroup.Cells("LocPinX").Result(0)
        If ActivePage.Shapes("Рамка").Cells("User.Height").Result("mm") > 15 Then
             .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Datchik55
        Else
             .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Datchik15
        End If
        '.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = IIf(ActivePage.Shapes("Рамка").Cells("Prop.CHAPTER").Result(0) = 0, Datchik55, Datchik15)
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0"
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).FormulaU = "Height * 1"
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "25 mm"
        .CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Formula = vsoGroup.Cells("Width").Result(0)
        .CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.5 mm"
        .CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "1"
'        .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "24 pt"
        .CellsSRC(visSectionObject, visRowTextXForm, visXFormLocPinX).FormulaU = "TxtWidth * 0.5"
        .Cells("Controls.TextPos").FormulaU = "Width * 0.5"
        .Cells("Prop.PerenosOboz").FormulaU = 1
        .Cells("Controls.TextPos.Y").FormulaU = "Height * 0"
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
'    'Клеммник
    vsoGroup.ContainingPage.Drop Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.ItemU("klemmnik"), 0, 0
    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).Formula = vsoGroup.Cells("PinX").Result(0) - vsoGroup.Cells("LocPinX").Result(0)
    If ActivePage.Shapes("Рамка").Cells("User.Height").Result("mm") > 15 Then
         Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Datchik55 - 5 / 25.4
    Else
         Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = Datchik15 - 5 / 25.4
    End If
    'Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).Formula = IIf(ActivePage.Shapes("Рамка").Cells("Prop.CHAPTER").Result(0) = 0, Datchik55, Datchik15) - 5 / 25.4
    Application.ActiveWindow.Selection(1).Cells("Controls.Line").GlueTo shpShkaf.Cells("Connections.X1")
    Set GetShkafDown = shpShkaf
End Function

Public Sub svpDEL()
'------------------------------------------------------------------------------------------------------------
' Macros        : svpDEL - Удаляет листы схемы внешних проводок
'------------------------------------------------------------------------------------------------------------
    If MsgBox("Удалить листы схемы внешних проводок?", vbQuestion + vbOKCancel, "САПР-АСУ: Удалить листы СВП") = vbOK Then
        del_pages cListNameSVP
        'MsgBox "Старая версия спецификации удалена", vbInformation
    End If
End Sub

Function GetNazvanie(ShkafMesto As String, sel As Integer) As String
'sel=1  Элемент
'sel=2  Шкаф
'sel=3  Место
    Dim mMesto() As String
    Dim mShkaf() As String
    Dim mElement() As String
    
    If sel >= 1 Then
        mElement = Split(ShkafMesto, ActiveDocument.DocumentSheet.Cells("User.SA_PrefElement").ResultStr(0))
        If UBound(mElement) > 0 Then
            GetNazvanie = mElement(1)
        Else
            GetNazvanie = ""
        End If
    End If
    If sel >= 2 Then
        mShkaf = Split(mElement(0), ActiveDocument.DocumentSheet.Cells("User.SA_PrefShkaf").ResultStr(0))
        If UBound(mShkaf) > 0 Then
            GetNazvanie = mShkaf(1)
        Else
            GetNazvanie = ""
        End If
    End If
    If sel = 3 Then
        mMesto = Split(mShkaf(0), ActiveDocument.DocumentSheet.Cells("User.SA_PrefMesto").ResultStr(0))
        If UBound(mMesto) > 0 Then
            GetNazvanie = mMesto(1)
        Else
            GetNazvanie = ""
        End If
    End If
End Function
