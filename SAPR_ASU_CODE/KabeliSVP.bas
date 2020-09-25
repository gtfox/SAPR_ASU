'------------------------------------------------------------------------------------------------------------
' Module        : KabeliSVP - Кабели на эл. схеме, на планах и на схеме внешних проводок (СВП)
' Author        : gtfox
' Date          : 2020.09.21
' Description   : Вставка и нумерация кабелей на эл. схеме, на планах и автосоздание схемы внешних проводок (СВП)
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------



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
    Dim colWiresTemp As Collection
    Dim vsoMaster As Visio.Master
    Dim MultiCable As Boolean '1 вход = 1 кабель
    Dim PinX As Double
    Dim PinY As Double
    
    PinX = shpSensor.Cells("PinX").Result(0)
    PinY = shpSensor.Cells("PinY").Result(0)
    
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
                'Кабели ссылаются не на датчик, а на конкретные входы в датчике
                shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
                'Связываем входы с кабелями
                shpSensorIO.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
            End If
        Next
    Else
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If shpSensorIO.Name Like "SensorIO*" Then
                'Находим подключенные провода и суем их в коллекцию
                Set colWiresTemp = FillColWires(shpSensorIO)
                'Сращиваем все коллекции в одну
                For Each vsoShape In colWiresTemp
                    colWires.Add vsoShape
                Next
            End If
        Next
        'Вставляем шейп кабеля
        Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensor.Cells("PinX").Result(0), shpSensor.Cells("PinY").Result(0) + 0.19685)
        'Добавляем подключенные провода в группу с кабелем
        AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
        'Чмсло проводов в кабеле
        shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
        'Кабель ссылается не на датчик, а на конкретный вход в датчике
        shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
        'Связываем вход с кабелем
        shpSensorIO.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
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

Sub DeleteCableSH(shpKabel As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteCableSH - Чистим ссылку в подключенном датчике перед удалением кабеля, и удаляем кабель
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpSensorIO As Visio.Shape
    
    'Находим датчик по ссылке в кабеле
    Set shpSensorIO = HyperLinkToShape(shpKabel.Cells("User.LinkToSensor").ResultStr(0))
    'Чистим ссылку на кабель в датчике
    On Error Resume Next
    shpSensorIO.Cells("User.LinkToCable").FormulaU = ""
    
End Sub

Sub AddSensorOnSVP()
    
    Dim shpSensorIO As Visio.Shape
    Dim shpSensor As Visio.Shape
    Dim shpTerm As Visio.Shape
    Dim shpCable As Visio.Shape
    Dim shpWire As Visio.Shape
    Dim colCables As Collection
    Dim colWires As Collection
    Dim colTerms As Collection
    Dim vsoSelection As Visio.Selection
    Dim vsoMaster As Visio.Master
    Dim shpKabelSVP As Visio.Shape
    Dim vsoGroup As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim shpSensorSVP As Visio.Shape
    Dim MultiCable As Boolean
    Dim PastePoint As Double
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
    Dim SensorSVPPinY As Double
    Const SVPWireL As Double = 10 / 25.4 'Длина проводов торчащих из шины на СВП
    Const Interval As Double = 5 / 25.4 'Расстояние между датчиками на СВП
    Const Klemma As Double = 240 / 25.4 'Высота расположения клеммника шкафа на СВП
    Const Datchik As Double = 100 / 25.4 'Высота расположения датчика на СВП
    
    Set colCables = New Collection
    Set colWires = New Collection
    Set colTerms = New Collection
    Set vsoSelection = ActiveWindow.Selection
    Set shpSensor = ActiveWindow.Selection.PrimaryItem 'ActiveDocument.Pages("Схема").Shapes("Sensor.582")
    MultiCable = shpSensor.Cells("Prop.MultiCable").Result(0)
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_SVP.vss").Masters.Item("KabelSVP")
    
    If MultiCable Then
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If shpSensorIO.Name Like "SensorIO*" Then
                'Находим подключенные провода и суем их в коллекцию
                Set colWires = FillColWires(shpSensorIO)
                'Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
                Set colTerms = FillColTerms(colWires)
                'Выделяем всех
                vsoSelection.Select shpSensor, visSelect 'Датчик
                vsoSelection.Select colWires.Item(1).Parent, visSelect 'Кабель
                For Each shpTerm In colTerms
                    vsoSelection.Select shpTerm, visSelect 'Клеммы шкафа
                Next

            End If
        Next
    Else
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If shpSensorIO.Name Like "SensorIO*" Then
                'Находим подключенные провода и суем их в коллекцию
                Set colWiresTemp = FillColWires(shpSensorIO)
                'Сращиваем все коллекции в одну
                For Each vsoShape In colWiresTemp
                    colWires.Add vsoShape
                Next
                'Добавляем текущий кабель в коллекцию кабелей
                colCables.Add colWiresTemp.Item(1).Parent
            End If
        Next
        'Находим подключенные к проводам клеммы шкафа и суем их в коллекцию
        Set colTerms = FillColTerms(colWires)
        'Выделяем всех
        vsoSelection.Select shpSensor, visSelect 'Датчик
        For Each shpCable In colCables
            vsoSelection.Select shpCable, visSelect 'Кабели
        Next
        For Each shpTerm In colTerms
            vsoSelection.Select shpTerm, visSelect 'Клеммы шкафа
        Next
    End If

    'Копируем что насобирали
    vsoSelection.Copy
    'Отключаем события автоматизации (чтобы не перенумеровалось все)
    Application.EventsEnabled = 0

    ActiveWindow.Page = ActiveDocument.Pages("СВП")
    ActivePage.Paste
    'Application.ActiveDocument.Pages("СВП").Paste



    Set vsoGroup = ActiveWindow.Selection.Group
    
    Set colCables = New Collection
    vsoGroup.Cells("PinX").Formula = "(25 mm+" & PastePoint & "+" & Interval & "+" & vsoGroup.Cells("LocPinX").Result(0) & "-TheDoc!User.OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale"
    vsoGroup.Cells("PinY").Formula = Klemma & "-" & vsoGroup.Cells("LocPinY").Result(0)
    For Each vsoShape In vsoGroup.Shapes
         Select Case vsoShape.Cells("User.SAType").Result(0)
            Case typeSensor, typeActuator
                Set shpSensorSVP = vsoShape
            Case typeCableSH
                colCables.Add vsoShape
         End Select
    Next
    
    vsoGroup.Ungroup
    
        'Включаем события автоматизации
    Application.EventsEnabled = -1
    
    shpSensorSVP.Cells("PinY").Formula = Datchik
    SensorSVPPinY = shpSensorSVP.Cells("PinY").Result(0)
    For Each shpCable In colCables
        'В кабеле находим длину провода
        DoEvents 'На-я тут этот DoEvents?
        WireHeight = shpCable.Shapes(1).Cells("Height").Result(0)
        'Вставляем шейп кабеля СВП
        Set shpKabelSVP = shpCable.ContainingPage.Drop(vsoMaster, shpCable.Cells("PinX").Result(0) + shpCable.Cells("Width").Result(0) * 0.5, SensorSVPPinY + WireHeight - SVPWireL)
        shpKabelSVP.Cells("Width").Formula = WireHeight - SVPWireL * 2
        shpKabelSVP.Cells("Prop.Number").Formula = shpCable.Cells("Prop.Number").Result(0)
        shpKabelSVP.Cells("Prop.Marka").Formula = """" & shpCable.Cells("User.Marka").ResultStr(0) & """"
        shpKabelSVP.Cells("Prop.WireCount").Formula = shpCable.Shapes.Count
        'Ищем вход в датчике соединенный с текущим кабелем
        For Each shpWire In shpCable.Shapes
            For i = 1 To shpWire.Connects.Count
                If shpWire.Connects(i).ToSheet.Name Like "Term*" Then
                    Set cellKlemmaShkafa = shpWire.Connects(i).ToCell
                    NumberKlemmaShkafa = shpWire.Connects(i).ToSheet.Cells("Prop.Number").Result(0)
                ElseIf shpWire.Connects(i).ToSheet.Name Like "PLCTerm*" Then
                    Set cellKlemmaDatchika = shpWire.Connects(i).ToCell
                    NumberKlemmaDatchika = shpWire.Connects(i).ToSheet.Cells("User.Number").Result(0)
                    WireNumber = CInt(Right(shpWire.Connects(i).ToSheet.Name, 1))
                End If
            Next
            Set cellWireDown = shpKabelSVP.Cells("Controls.W" & WireNumber & "1")
            Set cellWireUp = shpKabelSVP.Cells("Controls.W" & WireNumber & "2")
            'Клеим провод
            cellWireDown.GlueTo cellKlemmaDatchika
            shpKabelSVP.Cells("Prop.WIRE" & WireNumber & "1").Formula = NumberKlemmaShkafa
            cellWireUp.GlueTo cellKlemmaShkafa
            shpKabelSVP.Cells("Prop.WIRE" & WireNumber & "2").Formula = NumberKlemmaDatchika
            
        Next
        shpKabelSVP.Cells("Controls.BendPnt").Formula = shpKabelSVP.Cells("Width").Result(0) * 0.5
        
    Next
    
    'Удаляем кабели эл. схемы
    For Each shpCable In colCables
        shpCable.Delete
    Next
    
    


End Sub

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