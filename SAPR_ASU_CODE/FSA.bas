'------------------------------------------------------------------------------------------------------------
' Module        : FSA - Функциональная схема автоматизации
' Author        : gtfox
' Date          : 2021.05.25
' Description   : Вставка датчиков и приводов со схемы на ФСА
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
Option Explicit


Public Sub PageFSAAddSensorsFrm()
    Load frmPageFSAAddSensors
    frmPageFSAAddSensors.Show
End Sub

Public Sub AddSensorsOnFSA(NazvanieShemy As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddSensorsOnFSA - Вставляет все датчики со схемы на ФСА
                'Вставка датчиков и приводов со схемы на ФСА, если датчик уже есть, то не вставляет его.
'------------------------------------------------------------------------------------------------------------
    Dim PageParent As String, NameIdParent As String, AdrParent As String
    Dim PageChild  As String, NameIdChild As String, AdrChild As String
    Dim FSAvss As Document
    Dim vsoPageFSA As Visio.Page
    Dim vsoPageCxema As Visio.Page
    Dim colPagesCxema As Collection
    Dim shpSensorOnFSA As Visio.Shape
    Dim shpSensorOnCxema As Visio.Shape
    Dim colSensorOnFSA As Collection
    Dim colSensorToFSA As Collection
    Dim colSensorOnCxema As Collection
    Dim vsoSelection As Visio.Selection
    Dim vsoGroup As Visio.Shape
    Dim nCount As Double
    Dim DropX As Double
    Dim DropY As Double
    
    If NazvanieShemy = "" Then
        MsgBox "Нет схем для вставки", vbExclamation, "Название схемы пустое"
        Exit Sub
    End If
    
    Set colSensorOnFSA = New Collection
    Set colSensorToFSA = New Collection
    Set colPagesCxema = New Collection
    Set colSensorOnCxema = New Collection
    Set vsoSelection = ActiveWindow.Selection
    Set vsoPageFSA = Application.ActivePage  '.Pages("Схема")
    Set vsoPageCxema = ActiveDocument.Pages(cListNameCxema)
    Set FSAvss = Application.Documents.Item("SAPR_ASU_FSA.vss")
    
    DropY = ActivePage.PageSheet.Cells("PageHeight").Result(0)
    
    'Костыль чтобы почистить события в шейпе
    Set shpSensorOnFSA = vsoPageFSA.Drop(FSAvss.Masters.Item("SensorFSA"), 0, 0)
    shpSensorOnFSA.Delete
    ActiveDocument.Masters.ItemU("SensorFSA").Shapes(1).Cells("EventDrop").Formula = "CALLTHIS(""AutoNumber.AutoNumFSA"")"
    ActiveDocument.Masters.ItemU("SensorFSA").Shapes(1).Cells("EventMultiDrop").Formula = """"""

    'Берем все листы одной схемы
    For Each vsoPageCxema In ActiveDocument.Pages
        If vsoPageCxema.Name Like cListNameCxema & "*" Then
            If vsoPageCxema.PageSheet.CellExists("Prop.SA_NazvanieShemy", 0) Then
                If vsoPageCxema.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then
                    colPagesCxema.Add vsoPageCxema
                End If
            End If
        End If
    Next

    'Находим что уже есть на ФСА (связанные датчики)
    For Each shpSensorOnFSA In vsoPageFSA.Shapes
        If ShapeSATypeIs(shpSensorOnFSA, typeFSASensor) Then
            colSensorOnFSA.Add shpSensorOnFSA, shpSensorOnFSA.Cells("User.NameParent").ResultStr(0)
        End If
    Next
    
    'Суем туда же все из СХЕМЫ. Одинаковое не влезает => ошибка. Что не влезло: нам оно то и нужно
    For Each vsoPageCxema In colPagesCxema
        For Each shpSensorOnCxema In vsoPageCxema.Shapes
            If ShapeSATypeIs(shpSensorOnCxema, typeSensor) Or ShapeSATypeIs(shpSensorOnCxema, typeActuator) Then
                nCount = colSensorOnFSA.Count
                On Error Resume Next
                colSensorOnFSA.Add shpSensorOnCxema, shpSensorOnCxema.Cells("User.Name").ResultStr(0)
                If colSensorOnFSA.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                    colSensorToFSA.Add shpSensorOnCxema
                    nCount = colSensorOnFSA.Count
                End If
            End If
        Next
    Next

    'Очищаем коллекцию для вставляемых датчиков
    Set colSensorOnFSA = New Collection
    
    'Вставляем недостающие датчики на ФСА
    For Each shpSensorOnCxema In colSensorToFSA
        Select Case ShapeSAType(shpSensorOnCxema)
            Case typeSensor
                Set shpSensorOnFSA = vsoPageFSA.Drop(ActiveDocument.Masters.ItemU("SensorFSA"), DropX, DropY)
                DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2
                'Связываем датчик на ФСА и датчик наэл. схеме
                AddReferenceSensor shpSensorOnFSA, shpSensorOnCxema
            Case typeActuator
'                Set shpSensorOnFSA = vsoPageFSA.Drop(FSAvss.Masters.Item("ActuatorFSA"), DropX, DropY)
'                DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2
            Case Else
                
        End Select
'        'Собираем в коллецию вставленные датчики
'        colSensorOnFSA.Add shpSensorOnFSA
    Next

'    'Выделяем вставленные датчики
'    ActiveWindow.DeselectAll
'    For Each shpSensorOnFSA In colSensorOnFSA
'        ActiveWindow.Select shpSensorOnFSA, visSelect
'    Next
'    With ActiveWindow.Selection
'        'Выравниваем по горизонтали
'        .Align visHorzAlignNone, visVertAlignMiddle, False
'        'Распределяем по горизонтали
'        .Distribute visDistHorzSpace, False
'        DoEvents
'        'Поднимаем вверх
'        .Move 0, ActivePage.PageSheet.Cells("PageHeight").Result(0) - .PrimaryItem.Cells("PinY").Result(0)
'    End With
    ActiveWindow.DeselectAll
    
End Sub
