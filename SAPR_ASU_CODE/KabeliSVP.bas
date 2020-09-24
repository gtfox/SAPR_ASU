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
                'Чмсло проводов в кабеле
                shpKabel.Cells("Prop.WireCount").FormulaU = colWires.Count
                'Кабели ссылаются не на датчик, а на конкретные входы в датчике
                shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensorIO.ContainingPage.NameU + "/" + shpSensorIO.NameID + """"
                'Связываем входы с проводами
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
        'Кабель ссылается на датчик
        shpKabel.Cells("User.LinkToSensor").FormulaU = """" + shpSensor.ContainingPage.NameU + "/" + shpSensor.NameID + """"
        'Связываем кабель и датчик
        shpSensor.Cells("User.LinkToCable").FormulaU = """" + shpKabel.ContainingPage.NameU + "/" + shpKabel.NameID + """"
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
    Dim shpSensor As Visio.Shape
    Dim shpSensorIO As Visio.Shape
    
    'Находим датчик по ссылке в кабеле
    Set shpSensor = HyperLinkToShape(shpKabel.Cells("User.LinkToSensor").ResultStr(0))
    'Чистим ссылку на кабель в датчике
    On Error Resume Next
    shpSensor.Cells("User.LinkToCable").FormulaU = ""
    
End Sub
