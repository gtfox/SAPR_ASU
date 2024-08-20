'------------------------------------------------------------------------------------------------------------
' Module        : KabeliSVP - Кабели на эл. схеме
' Author        : gtfox
' Date          : 2020.09.21
' Description   : Вставка и нумерация кабелей на эл. схеме
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Const DyKlemma As Double = 12.5 / 25.4 'Высота расположения клеммы шкафа относительно датчика на Схеме
Const DyKabel As Double = 5 / 25.4 'Высота подъёма кабеля относительно датчика на Схеме


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
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("Kabel")

    If MultiCable Then
        'Перебираем все входы в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeCxemaSensorIO) Then
                'Добавляем клеммы и провода
                If iOptions <= 2 Then AddKlemmyIProvoda shpSensorIO '1=Клеммы и провода
                If iOptions >= 2 Then '3=Кабели из проводов
                    'Вставляем шейп кабеля
                    Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensorIO.Cells("PinX").Result(0) + PinX, shpSensorIO.Cells("PinY").Result(0) + PinY + DyKabel)
                    'Находим подключенные провода и суем их в коллекцию
                    Set colWires = FillColWiresOnPage(shpSensorIO)
                    'Добавляем подключенные провода в группу с кабелем
                    AddToGroupCable shpKabel, shpKabel.ContainingPage, colWires
                    'Число проводов в кабеле
                    shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
                    'Сохраняем к какому шкафу подключен кабель
                    If ShapeSATypeIs(colWires.Item(1).Connects(1).ToSheet, typeCxemaTerm) Then 'клемма шкафа
                        shpKabel.Cells("User.LinkToBox").Formula = colWires.Item(1).Connects(1).ToSheet.NameID & "!User.FullName.Prompt"
                    ElseIf ShapeSATypeIs(colWires.Item(1).Connects(1).ToSheet, typeCxemaSensorTerm) Then 'клемма датчика
                        shpKabel.Cells("User.LinkToBox").Formula = colWires.Item(1).Connects(2).ToSheet.NameID & "!User.FullName.Prompt"
                    End If
                    shpKabel.Cells("User.LinkToSensor").Formula = shpSensor.NameID & "!User.Name"
                End If
            End If
        Next
        


    Else
    
        'Собираем провода со всех входов в датчике
        For Each shpSensorIO In shpSensor.Shapes
            If ShapeSATypeIs(shpSensorIO, typeCxemaSensorIO) Then
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
            If ShapeSATypeIs(shpSensorIO, typeCxemaSensorIO) Then
                If iOptions >= 2 Then 'Кабели
                    If ActiveWindow.Selection.Count > 2 Then
                        oldCount = colWires.Count
                        For Each vsoShape In ActiveWindow.Selection 'Кабель из выделенных проводов
                            If ShapeSATypeIs(vsoShape, typeCxemaWire) Then
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
                MsgBox "Выделите минимум 2 провода для создания кабеля", vbInformation + vbOKOnly, "САПР-АСУ: Создание кабеля"
                Exit Sub
            End If
            'Вставляем шейп кабеля
            Set shpKabel = shpSensor.ContainingPage.Drop(vsoMaster, shpSensor.Cells("PinX").Result(0), shpSensor.Cells("PinY").Result(0) + DyKabel)
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
            If ShapeSATypeIs(shpKabel.Shapes(1).Connects(1).ToSheet, typeCxemaTerm) Then 'клемма шкафа
                shpKabel.Cells("User.LinkToBox").Formula = shpKabel.Shapes(1).Connects(1).ToSheet.NameID & "!User.FullName.Prompt"
            ElseIf ShapeSATypeIs(shpKabel.Shapes(1).Connects(1).ToSheet, typeCxemaSensorTerm) Then 'клемма датчика
                shpKabel.Cells("User.LinkToBox").Formula = shpKabel.Shapes(1).Connects(2).ToSheet.NameID & "!User.FullName.Prompt"
            End If
            shpKabel.Cells("User.LinkToSensor").Formula = shpSensor.NameID & "!User.Name"
        
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
        .Move 0#, DyKabel
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
    Dim NPin As Integer
    
    Set vsoPage = ActivePage
    Set vsoMasterKlemma = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("Term")
    Set vsoMasterProvod = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("w1")

    For Each shpSensorTerm In shpSensorIO.Shapes
        If ShapeSATypeIs(shpSensorTerm, typeCxemaSensorTerm) Then
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
            NPin = NPin + 1
            If NPin = shpSensorIO.Cells("Prop.NPin").Result(0) Then Exit For
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

Sub ClearCableSH(shpKabel As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearCableSH - Чистим ссылку в кабеле на кабель на плане
'------------------------------------------------------------------------------------------------------------
    shpKabel.CellsU("Hyperlink.Kabel.SubAddress").FormulaForceU = """"""
    shpKabel.CellsU("Hyperlink.Kabel.Frame").FormulaForceU = """"""
    shpKabel.CellsU("Hyperlink.Kabel.ExtraInfo").FormulaForceU = """"""
End Sub

Public Sub AddCableFromWires(shpProvod As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddCableFromWires - Вставляет кабель для выделенных проводов на эл.схеме
                'Вставляется шейп кабеля для подключенных проводов на эл.схеме
                'группируется с подключенными проводами, нумеруется, связываются ссылками друг на друга
                'Провода должны быть подключены к клеммам шкафа или датчика/привода
'------------------------------------------------------------------------------------------------------------
    Dim shpKabel As Visio.Shape
    Dim shpWire As Visio.Shape
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
    Dim TermType1 As Integer
    Dim TermType2 As Integer
    Dim MaxNumber As Double
    Dim MinNumber As Double
    Dim i As Integer

    PinY = shpProvod.Cells("PinY").Result(0)

    Set colWiresSelected = New Collection
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("Kabel")
    
    'Находим подключенные провода в выделении
    For Each shpWire In ActiveWindow.Selection
        If ShapeSATypeIs(shpWire, typeCxemaWire) Then
            If shpWire.Connects.Count = 0 Then
                MsgBox "Провод " & shpWire.Cells("Prop.Number").Result(0) & " не подключен", vbExclamation + vbOKOnly, "САПР-АСУ: Создание кабеля"
                Exit Sub
            ElseIf shpWire.Connects.Count = 1 Then
                MsgBox "Провод " & shpWire.Cells("Prop.Number").Result(0) & " не подключен одним концом", vbExclamation + vbOKOnly, "САПР-АСУ: Создание кабеля"
                Exit Sub
            ElseIf shpWire.Connects.Count = 2 Then
                TermType1 = ShapeSAType(shpWire.Connects(1).ToSheet)
                TermType2 = ShapeSAType(shpWire.Connects(2).ToSheet)
                If TermType1 = typeCxemaSensorTerm Or TermType1 = typeCxemaTerm Then 'Or TermType1 = typePLCTerm
                    If TermType2 = typeCxemaSensorTerm Or TermType2 = typeCxemaTerm Then 'Or TermType2 = typePLCTerm
                        colWiresSelected.Add shpWire
                    Else
                        MsgBox "Провод " & shpWire.Cells("Prop.Number").Result(0) & " должен быть подключен к клеммам шкафа или датчика/привода", vbExclamation + vbOKOnly, "САПР-АСУ: Создание кабеля"
                        Exit Sub
                    End If
                Else
                    MsgBox "Провод " & shpWire.Cells("Prop.Number").Result(0) & " должен быть подключен к клеммам шкафа или датчика/привода", vbExclamation + vbOKOnly, "САПР-АСУ: Создание кабеля"
                    Exit Sub
                End If
            End If
        End If
    Next
    
    If colWiresSelected.Count <= 1 Then
        MsgBox "Выделите минимум 2 провода для создания кабеля", vbExclamation + vbOKOnly, "САПР-АСУ: Создание кабеля"
        Exit Sub
    End If

    'Находим позицию вставки шейпа кабеля
    MinNumber = 1.79769313486231E+308
    For i = 1 To colWiresSelected.Count
        PinX = colWiresSelected(i).Cells("PinX").Result(0)
        If PinX < MinNumber Then MinNumber = PinX
        If PinX > MaxNumber Then MaxNumber = PinX
    Next

    'Вставляем шейп кабеля
    Set shpKabel = shpProvod.ContainingPage.Drop(vsoMaster, PinX, PinY)
    'Добавляем подключенные провода в группу с кабелем
    AddToGroupCable shpKabel, shpKabel.ContainingPage, colWiresSelected
    'Число проводов в кабеле
    shpKabel.Cells("Prop.WireCount").FormulaU = colWiresSelected.Count
    shpKabel.Cells("Width").Formula = MaxNumber - MinNumber + DyKabel
            
    Set vsoSelection = ActiveWindow.Selection
    vsoSelection.Select shpKabel, visSelect
    vsoSelection.Move 0#, Abs(shpKabel.Cells("PinY").Result(0) - PinY)

    'Сохраняем к какому шкафу подключен кабель
    If ShapeSATypeIs(shpKabel.Shapes(1).Connects(1).ToSheet, typeCxemaTerm) And ShapeSATypeIs(shpKabel.Shapes(1).Connects(2).ToSheet, typeCxemaTerm) Then 'соединены 2 шкафа
        'При соединении кабелем двух шкафов: Кто выше тот и шкаф :)
    '    shpKabel.Cells("User.LinkToBox.Prompt").Formula = """" & shpKabel.Cells("User.LinkToBox").ResultStr(0) & """"
    '    shpKabel.Cells("User.LinkToSensor.Prompt").Formula = """" & shpKabel.Cells("User.LinkToSensor").ResultStr(0) & """"
        If shpKabel.Shapes(1).Connects(1).ToSheet.Cells("PinY").Result(0) > shpKabel.Shapes(1).Connects(2).ToSheet.Cells("PinY").Result(0) Then
            shpKabel.Cells("User.LinkToBox").Formula = shpKabel.Shapes(1).Connects(1).ToSheet.NameID & "!User.FullName.Prompt"
            shpKabel.Cells("User.LinkToSensor").Formula = shpKabel.Shapes(1).Connects(2).ToSheet.NameID & "!User.FullName.Prompt"
        Else
            shpKabel.Cells("User.LinkToBox").Formula = shpKabel.Shapes(1).Connects(2).ToSheet.NameID & "!User.FullName.Prompt"
            shpKabel.Cells("User.LinkToSensor").Formula = shpKabel.Shapes(1).Connects(1).ToSheet.NameID & "!User.FullName.Prompt"
        End If
    ElseIf ShapeSATypeIs(shpKabel.Shapes(1).Connects(1).ToSheet, typeCxemaTerm) Then 'Соединен шкаф и датчик/привод 'клемма шкафа
        shpKabel.Cells("User.LinkToBox").Formula = shpKabel.Shapes(1).Connects(1).ToSheet.NameID & "!User.FullName.Prompt"
        shpKabel.Cells("User.LinkToSensor").Formula = shpKabel.Shapes(1).Connects(2).ToSheet.NameID & "!User.Name"
    ElseIf ShapeSATypeIs(shpKabel.Shapes(1).Connects(1).ToSheet, typeCxemaSensorTerm) Then 'клемма датчика
        shpKabel.Cells("User.LinkToBox").Formula = shpKabel.Shapes(1).Connects(2).ToSheet.NameID & "!User.FullName.Prompt"
        shpKabel.Cells("User.LinkToSensor").Formula = shpKabel.Shapes(1).Connects(1).ToSheet.NameID & "!User.Name"
    End If

    Application.EventsEnabled = -1
    ThisDocument.InitEvent
End Sub

Function FindKabelFromSensor(shpSensor As Visio.Shape) As Visio.Shape
'------------------------------------------------------------------------------------------------------------
' Function        : FindKabelFromSensor - Находим кабель подключенный к датчику/приводу
'------------------------------------------------------------------------------------------------------------
    Dim shpSensorIO As Visio.Shape
    Dim shpSensorTerm As Visio.Shape

    For Each shpSensorIO In shpSensor.Shapes
        If ShapeSATypeIs(shpSensorIO, typeCxemaSensorIO) Then
            For Each shpSensorTerm In shpSensorIO.Shapes
                If ShapeSATypeIs(shpSensorTerm, typeCxemaSensorTerm) Then
                    If shpSensorTerm.FromConnects.Count = 1 Then
                        If ShapeSATypeIs(shpSensorTerm.FromConnects.FromSheet, typeCxemaWire) Then
                            If shpSensorTerm.FromConnects.FromSheet.Parent.Type = visTypeGroup Then
                                Set FindKabelFromSensor = shpSensorTerm.FromConnects.FromSheet.Parent
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    Set FindKabelFromSensor = Nothing
End Function