'------------------------------------------------------------------------------------------------------------
' Module        : KabeliSVP - Кабели на эл. схеме, на планах и на схеме внешних проводок (СВП)
' Author        : gtfox
' Date          : 2020.09.21
' Description   : Вставка и нумерация кабелей на эл. схеме, на планах и автосоздание схемы внешних проводок (СВП)
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Public PastePoint As Variant 'Точка вставки следующего датчика
Const KonecLista As Double = 10 / 25.4 'Расстояние от правого края листа, за которое не дожны заходить фигуры
Const DyKlemma As Double = 12.5 / 25.4 'Высота расположения клеммы шкафа относительно датчика на Схеме


Public Sub AddCableOnSensor(shpSensor As Visio.Shape, Optional iOptions As Integer = 0)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddCableOnSensor - Вставляет кабель для подключенного датчика/привода на эл.схеме
                'Вставляется шейп кабеля для подключенного датчика/привода на эл.схеме
                'группируется с подключенными проводами, нумеруется, связываются ссылками друг на друга
                'Если датчик многокабельный(MultiCable=true), то кабели ссылаются не на датчик, а на конкретные входы в датчике
                'iOptions - 1=Клеммы и провода 2=Клеммы+Кабели 3=Кабели из проводов
'------------------------------------------------------------------------------------------------------------
    Dim shpKabel As Visio.Shape
    Dim shpSensorIO As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim colWires As Collection
    Dim colWiresIO As Collection
    Dim vsoMaster As Visio.Master
    Dim MultiCable As Boolean '1 вход = 1 кабель
    Dim NazvanieShemy As String
    Dim PinX As Double
    Dim PinY As Double
    
    PinX = shpSensor.Cells("PinX").Result(0)
    PinY = shpSensor.Cells("PinY").Result(0)
    
    NazvanieShemy = shpSensor.ContainingPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0)
    
    MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)
    Set colWires = New Collection
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_SVP.vss").Masters.Item("Kabel")

    If MultiCable Then
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeSensorIO) Then
                'Добавляем клеммы и провода
                If iOptions <= 2 Then AddKlemmyIProvoda shpSensorIO '1=Клеммы и провода
                If iOptions >= 2 Then '3=Кабели из проводов
                    'Вставляем шейп кабеля
                    Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensorIO.Cells("PinX").Result(0) + PinX, shpSensorIO.Cells("PinY").Result(0) + PinY + 0.196850393700787)
                    'Находим подключенные провода и суем их в коллекцию
                    Set colWires = FillColWires(shpSensorIO)
                    'Добавляем подключенные провода в группу с кабелем
                    AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
                    'Число проводов в кабеле
                    shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
                    'Сохраняем к какому шкафу подключен кабель
                    If NazvanieShemy = "" Then 'если на листе несколько шкафов то...
                        'Определяем к какому шкафу/коробке принадлежит клеммник
                        '-------------Пока не реализовано----------------------
                    Else
                        shpKabel.Cells("User.LinkToBox").Formula = """" & NazvanieShemy & """"
                    End If
'                    'Кабели ссылаются не на датчик, а на конкретные входы в датчике
'                    shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
'                    'Связываем входы с кабелями
'                    shpSensorIO.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
                End If
            End If
        Next
    Else
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeSensorIO) Then
                'Добавляем клеммы и провода
                If iOptions <= 2 Then AddKlemmyIProvoda shpSensorIO 'Клеммы
                If iOptions >= 2 Then 'Кабели
                    'Находим подключенные провода на конкретном IO и суем их в коллекцию
                    Set colWiresIO = FillColWires(shpSensorIO)
                    'Добавляем провода с конкретного входа в общую колекцию проводов датчика
                    For Each vsoShape In colWiresIO
                        colWires.Add vsoShape
                    Next
                End If
            End If
        Next
        If iOptions >= 2 Then 'Кабели
            'Вставляем шейп кабеля
            Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensor.Cells("PinX").Result(0), shpSensor.Cells("PinY").Result(0) + 0.19685)
            'Добавляем подключенные провода в группу с кабелем
            AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
            'Число проводов в кабеле
            shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
            'Сохраняем к какому шкафу подключен кабель
            If NazvanieShemy = "" Then 'если на листе несколько шкафов то...
                'Определяем к какому шкафу/коробке принадлежит клеммник
                '-------------Пока не реализовано----------------------
            Else
                shpKabel.Cells("User.LinkToBox").Formula = """" & NazvanieShemy & """"
            End If
'            'Кабель ссылается не на датчик, а на конкретный вход в датчике
'            shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
'            'Связываем вход с кабелем
'            shpSensorIO.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
        End If
    End If
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
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
            shpWire.CellsU("EventDrop").FormulaU = """"""
            shpWire.CellsU("EventMultiDrop").FormulaU = """"""
            .Select shpWire, visSelect
        Next
        .AddToGroup
        'Сдвигаем вверх
        .DeselectAll
        .Select shpKabel, visSelect
        .Move 0#, 0.19685
    End With
End Sub

Sub AddKlemmyIProvoda(shpSensorIO As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddKlemmyIProvoda - Вставляем клеммы шкафа и подключаем провода к датчику
'------------------------------------------------------------------------------------------------------------
'    Const DyKlemma As Double = 22.5 / 25.4 'Высота расположения клеммы шкафа относительно датчика на Схеме
    Dim vsoPage As Visio.Page
    Dim vsoMasterKlemma As Visio.Master
    Dim vsoMasterProvod As Visio.Master
    Dim shpKlemma As Visio.Shape
    Dim shpProvod As Visio.Shape
    Dim shpSensorTerm As Visio.Shape
    Dim cellKlemmaShkafa As Visio.Cell
    Dim cellKlemmaDatchika As Visio.Cell
    Dim cellProvodDown As Visio.Cell
    Dim cellProvodUp As Visio.Cell
    Dim AbsPinX As Double
    Dim AbsPinY As Double
    
    Set vsoPage = ActivePage
    Set vsoMasterKlemma = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("Term")
    Set vsoMasterProvod = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("w1")

    For Each shpSensorTerm In shpSensorIO.Shapes
        If ShapeSATypeIs(shpSensorTerm, typeSensorTerm) Then
            AbsPinX = shpSensorTerm.Cells("User.AbsPinX").Result(0)
            AbsPinY = shpSensorTerm.Cells("User.AbsPinY").Result(0)
            'Вставляем клеммму
            Set shpKlemma = vsoPage.Drop(vsoMasterKlemma, AbsPinX, AbsPinY + DyKlemma)
            'Вставляем провод
            Set shpProvod = vsoPage.Drop(vsoMasterProvod, AbsPinX, AbsPinY)
            'Клеим провод
            Set cellKlemmaDatchika = shpSensorTerm.CellsSRC(visSectionConnectionPts, visRowConnectionPts, 0)
            Set cellKlemmaShkafa = shpKlemma.CellsSRC(visSectionConnectionPts, visRowConnectionPts + 1, 0)
            Set cellProvodDown = shpProvod.CellsU("BeginX")
            Set cellProvodUp = shpProvod.CellsU("EndX")
            cellProvodDown.GlueTo cellKlemmaDatchika
            cellProvodUp.GlueTo cellKlemmaShkafa
        End If
    Next
    ActiveWindow.DeselectAll
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
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
                colCablesOnElSh.Add colWires.Item(1).Parent, CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0))
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
        colCablesOnElSh.Add colWires.Item(1).Parent, CStr(colWires.Item(1).Parent.Cells("Prop.Number").Result(0))
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

    ActiveWindow.Page = ActiveDocument.Pages(vsoPageSVP.Name)
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
    
    'Ставим на место датчик
    shpSensorSVP.Cells("PinY").Formula = Datchik
    DoEvents 'На*уя тут этот DoEvents?
    
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

Public Sub PageSVPAddKabeliFrm()
    Load frmPageSVPAddKabeli
    frmPageSVPAddKabeli.Show
End Sub

Public Sub AddPagesSVP(NazvanieShemy As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddPagesSVP - Создает листы СВП
                'Заполняет листы СВП датчиками, отсортированными по возрастанию их координаты Х на эл. схеме
'------------------------------------------------------------------------------------------------------------
'    Dim NazvanieShemy As String
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
    Dim i As Integer, ii As Integer, j As Integer, N As Integer
    
    ShinaNumber = 1
    
    PastePoint = "25 mm - TheDoc!User.SA_FR_OffsetFrame"
    
    Set ThePage = ActivePage.PageSheet
    
    Set colShpDoc = New Collection
    
    PageName = cListNameCxema
    'If ThePage.CellExists("Prop.SA_NazvanieShemy", 0) Then NazvanieShemy = ThePage.Cells("Prop.SA_NazvanieShemy").ResultStr(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер.
'    NazvanieShemy = 4

    'Цикл поиска датчиков и приводов
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If InStr(1, vsoPage.Name, PageName) > 0 Then    'Берем те, что содержат "Схема" в имени
            If vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then    'Берем все схемы с номером той, на которую вставляем элемент
                Set colShpPage = New Collection
                For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                    Select Case ShapeSAType(vsoShapeOnPage) 'Если в шейпе есть тип, то -
                        Case typeSensor, typeActuator
                            'Собираем в коллекцию нужные для сортировки шейпы
                            colShpPage.Add vsoShapeOnPage
                        Case Else
                    End Select
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
        End If
    Next

    If colShpDoc.Count > 0 Then
        'Берем первую страницу СВП
        Set vsoPage = GetSAPageExist(cListNameSVP) 'ActiveDocument.Pages(cListNameSVP)
        If vsoPage Is Nothing Then Set vsoPage = AddSAPage(cListNameSVP)
        
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
                PastePoint = "25 mm - TheDoc!User.SA_FR_OffsetFrame"
                'Вставляем этот же датчик только на следующем листе
                AddSensorOnSVP colShpDoc.Item(i), vsoPage, ShinaNumber
            End If
        Next
    End If

    ActiveWindow.DeselectAll
End Sub