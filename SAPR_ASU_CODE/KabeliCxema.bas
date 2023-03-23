'------------------------------------------------------------------------------------------------------------
' Module        : KabeliSVP - Кабели на эл. схеме
' Author        : gtfox
' Date          : 2020.09.21
' Description   : Вставка и нумерация кабелей на эл. схеме
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

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
    Dim colWiresSelected As Collection
    Dim vsoMaster As Visio.Master
    Dim MultiCable As Boolean '1 вход = 1 кабель
    Dim NazvanieShkafa As String
    Dim PinX As Double
    Dim PinY As Double
    Dim oldCount As Integer
    
    PinX = shpSensor.Cells("PinX").Result(0)
    PinY = shpSensor.Cells("PinY").Result(0)
    
    NazvanieShkafa = shpSensor.Cells("User.Shkaf").ResultStr(0)
    
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
                    Set colWires = FillColWiresOnPage(shpSensorIO)
                    'Добавляем подключенные провода в группу с кабелем
                    AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
                    'Число проводов в кабеле
                    shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
                    'Сохраняем к какому шкафу подключен кабель
                    If NazvanieShkafa = "" Then 'если на листе несколько шкафов то...
                        'Определяем к какому шкафу/коробке принадлежит клеммник
                        '-------------Пока не реализовано----------------------
                    Else
                        shpKabel.Cells("User.LinkToBox").Formula = """" & NazvanieShkafa & """"
                    End If
'                    'Кабели ссылаются не на датчик, а на конкретные входы в датчике
'                    shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
'                    'Связываем входы с кабелями
'                    shpSensorIO.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
                End If
            End If
        Next
    Else
        'Собираем провода со всех входов в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeSensorIO) Then
                'Добавляем клеммы и провода
                If iOptions <= 2 Then AddKlemmyIProvoda shpSensorIO 'Клеммы
                If iOptions >= 2 Then 'Кабели
                    'Находим подключенные провода на конкретном IO и суем их в коллекцию
                    Set colWiresIO = FillColWiresOnPage(shpSensorIO)
                    'Добавляем провода с конкретного входа в общую колекцию проводов датчика
                    For Each vsoShape In colWiresIO
                        colWires.Add vsoShape
                    Next
                End If
            End If
        Next
        
        'Собираем выделенные провода со всех входов в датчике
        Set colWiresSelected = New Collection
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeSensorIO) Then
                If iOptions >= 2 Then 'Кабели
                    If ActiveWindow.Selection.Count > 2 Then
                        oldCount = colWires.Count
                        For Each vsoShape In ActiveWindow.Selection 'Кабель из выделенных проводов
                            If ShapeSATypeIs(vsoShape, typeWire) Then
                                On Error Resume Next
                                colWires.Add vsoShape, vsoShape.name
                                err.Clear
                                On Error GoTo 0
                                If colWires.Count = oldCount Then
                                    colWiresSelected.Add vsoShape, vsoShape.name 'Собираем те, что выделенны подсоединены к датчику
                                Else
                                    oldCount = colWires.Count
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Next
        
        If iOptions >= 2 Then 'Кабели
            If colWiresSelected.Count <= 1 And ActiveWindow.Selection.Count > 2 Then
                MsgBox "Выделите минимум 2 провода для создания кабеля", vbInformation + vbOKOnly, "САПР-АСУ: Экспорт для GitHub"
                Exit Sub
            End If
            'Вставляем шейп кабеля
            Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensor.Cells("PinX").Result(0), shpSensor.Cells("PinY").Result(0) + 0.19685)
            'Добавляем подключенные провода в группу с кабелем
            If colWiresSelected.Count > 1 Then
                AddToGroupCable shpKabel, shpKabel.ContainingPage, colWiresSelected
                'Число проводов в кабеле
                shpKabel.Cells("Prop.WireCount").FormulaU = colWiresSelected.Count
            Else
                AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
                'Число проводов в кабеле
                shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
            End If

            'Сохраняем к какому шкафу подключен кабель
            If NazvanieShkafa = "" Then 'если на листе несколько шкафов то...
                'Определяем к какому шкафу/коробке принадлежит клеммник
                '-------------Пока не реализовано----------------------
            Else
                shpKabel.Cells("User.LinkToBox").Formula = """" & NazvanieShkafa & """"
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
