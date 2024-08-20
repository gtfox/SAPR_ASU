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

Public Sub AddSensorsOnFSA(NazvanieShkafa As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddSensorsOnFSA - Вставляет все датчики со схемы на ФСА
                'Вставка датчиков и приводов со схемы на ФСА, если датчик уже есть, то не вставляет его.
'------------------------------------------------------------------------------------------------------------
    Dim PageParent As String, NameIdParent As String, AdrParent As String
    Dim PageChild  As String, NameIdChild As String, AdrChild As String
    Dim PageName As String
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
    
    PageName = cListNameCxema
    
    If NazvanieShkafa = "" Then
        MsgBox "Нет шкафа для вставки. Название шкафа пустое", vbExclamation, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set colSensorOnFSA = New Collection
    Set colSensorToFSA = New Collection
    Set colPagesCxema = New Collection
    Set colSensorOnCxema = New Collection
    Set vsoSelection = ActiveWindow.Selection
    Set vsoPageFSA = Application.ActivePage  '.Pages("Схема")
'    Set vsoPageCxema = ActiveDocument.Pages(cListNameCxema)
    Set FSAvss = Application.Documents.Item("SAPR_ASU_FSA.vss")
    
    DropY = ActivePage.PageSheet.Cells("PageHeight").Result(0)
    
    'Костыль чтобы почистить события в шейпе
    Set shpSensorOnFSA = vsoPageFSA.Drop(FSAvss.Masters.Item("SensorFSA"), 0, 0)
    shpSensorOnFSA.Delete
    ActiveDocument.Masters.ItemU("SensorFSA").Shapes(1).Cells("EventDrop").Formula = "CALLTHIS(""AutoNumber.AutoNumFSA"")"
    ActiveDocument.Masters.ItemU("SensorFSA").Shapes(1).Cells("EventMultiDrop").Formula = """"""
    Set shpSensorOnFSA = vsoPageFSA.Drop(FSAvss.Masters.Item("MotorFSA"), 0, 0)
    shpSensorOnFSA.Delete
    ActiveDocument.Masters.ItemU("MotorFSA").Shapes(1).Cells("EventDrop").Formula = "CALLTHIS(""AutoNumber.AutoNumFSA"")"
    ActiveDocument.Masters.ItemU("MotorFSA").Shapes(1).Cells("EventMultiDrop").Formula = """"""
    Set shpSensorOnFSA = vsoPageFSA.Drop(FSAvss.Masters.Item("ValveFSA"), 0, 0)
    shpSensorOnFSA.Delete
    ActiveDocument.Masters.ItemU("ValveFSA").Shapes(1).Cells("EventDrop").Formula = "CALLTHIS(""AutoNumber.AutoNumFSA"")"
    ActiveDocument.Masters.ItemU("ValveFSA").Shapes(1).Cells("EventMultiDrop").Formula = """"""
    
    'Находим что уже есть на ФСА (связанные датчики)
    For Each shpSensorOnFSA In vsoPageFSA.Shapes
        If ShapeSATypeIs(shpSensorOnFSA, typeFSASensor) Or ShapeSATypeIs(shpSensorOnFSA, typeFSAActuator) Then
            colSensorOnFSA.Add shpSensorOnFSA, shpSensorOnFSA.Cells("User.NameParent").ResultStr(0)
        End If
    Next
    
    'Суем туда же все из СХЕМЫ. Одинаковое не влезает => ошибка. Что не влезло: нам оно то и нужно
    For Each vsoPageCxema In ActiveDocument.Pages
        If vsoPageCxema.name Like PageName & "*" Then
            For Each shpSensorOnCxema In vsoPageCxema.Shapes
                If ShapeSATypeIs(shpSensorOnCxema, typeCxemaSensor) Or ShapeSATypeIs(shpSensorOnCxema, typeCxemaActuator) Then
                    If GetNazvanie(FindKabelFromSensor(shpSensorOnCxema).Cells("User.LinkToBox").ResultStr(0), 2) = NazvanieShkafa Then
                        nCount = colSensorOnFSA.Count
                        On Error Resume Next
                        colSensorOnFSA.Add shpSensorOnCxema, shpSensorOnCxema.Cells("User.Name").ResultStr(0)
                        err.Clear
                        On Error GoTo 0
                        If colSensorOnFSA.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - берем его себе
                            colSensorToFSA.Add shpSensorOnCxema
                        End If
                    End If
                End If
            Next
        End If
    Next

    'Очищаем коллекцию для вставляемых датчиков
    Set colSensorOnFSA = New Collection
    
    'Вставляем недостающие датчики на ФСА
    For Each shpSensorOnCxema In colSensorToFSA
        Select Case ShapeSAType(shpSensorOnCxema)
            Case typeCxemaSensor 'Датчик
                Set shpSensorOnFSA = vsoPageFSA.Drop(ActiveDocument.Masters.ItemU("SensorFSA"), DropX, DropY)
                DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2

                If shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "RK" Or shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "TC" Then 'Датчик температуры/Термопара TE
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """TE"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "BP" Then 'Датчик давления PT
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """PT"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "SP" Then 'Реле давления PS
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """PS"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "SL" Then 'Реле уровня LS
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """LS"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "BL" Then 'Датчик пламени BE
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """BE"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "SQ" Then 'Концевик GS
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """GS"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "SK" Then 'Термостат TS
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """TS"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "UZ" Then 'Частотник NY,UZ
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """NY"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "BN" Then 'Сигнализатор загазованности QN
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """QN"""
                Else
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """XX"""
                End If

                'Связываем датчик на ФСА и датчик наэл. схеме
                AddReferenceSensor shpSensorOnFSA, shpSensorOnCxema
            Case typeCxemaActuator 'Привод
                If shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "M" Then 'Насос, Вентилятор FG
                    Set shpSensorOnFSA = vsoPageFSA.Drop(ActiveDocument.Masters.ItemU("MotorFSA"), DropX, DropY)
                    DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """FG"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "B" Then 'Горелка FB
                    Set shpSensorOnFSA = vsoPageFSA.Drop(ActiveDocument.Masters.ItemU("MotorFSA"), DropX, DropY)
                    DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """FB"""
                    shpSensorOnFSA.Cells("Prop.Tip").FormulaU = "INDEX(2,Prop.Tip.Format)"
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "YA" Then 'Клапан электромагнитный FY, 3-х ходовой кран FV
                    Set shpSensorOnFSA = vsoPageFSA.Drop(ActiveDocument.Masters.ItemU("ValveFSA"), DropX, DropY)
                    DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """FV"""
                ElseIf shpSensorOnCxema.Cells("Prop.SymName").ResultStr(0) = "TV" Then 'Трансформатор запальника EZ
                    Set shpSensorOnFSA = vsoPageFSA.Drop(ActiveDocument.Masters.ItemU("MotorFSA"), DropX, DropY)
                    DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """EZ"""
                    shpSensorOnFSA.Cells("Prop.Tip").FormulaU = "INDEX(2,Prop.Tip.Format)" 'TODO нарисовать запальник и вписать цифру индекса
                Else
                    Set shpSensorOnFSA = vsoPageFSA.Drop(ActiveDocument.Masters.ItemU("ValveFSA"), DropX, DropY)
                    DropX = DropX + shpSensorOnFSA.Cells("Width").Result(0) * 2
                    shpSensorOnFSA.Cells("Prop.SymName").FormulaU = """XX"""
                End If
                'Связываем привод на ФСА и привод наэл. схеме
                AddReferenceSensor shpSensorOnFSA, shpSensorOnCxema
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
'    ActiveWindow.DeselectAll
    If colSensorToFSA.Count > 0 Then MsgBox "Добавлено " & colSensorToFSA.Count & " датчиков/приводов со схемы", vbInformation + vbOKOnly, "САПР-АСУ: датчики/приводы добавлены"
End Sub



'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'------------------------------------------FSAPodval--------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------



Sub GlueFSAPodval(shpFSAPodval As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : GlueFSAPodval - Приклеивает канал к подвалу и каналы между собой
                'Нужно выделить канал который хотим приклеить,
                'потом выделить канал к которому хотим приклеить,
                'потом выполнить GlueFSAPodval
'------------------------------------------------------------------------------------------------------------
    Dim shpToWhichGlue As Visio.Shape
    Dim vsoSelection As Visio.Selection
    Dim vsoShape As Visio.Shape
    
    Set vsoSelection = Application.ActiveWindow.Selection
    
    'Смотрим где находимся: в группе / на листе
    If shpFSAPodval.Parent.Type = visTypeGroup Then 'внутри группы
        vsoSelection.IterationMode = visSelModeOnlySub 'только внутренние шейпы
    ElseIf shpFSAPodval.Parent.Type = visTypePage Then 'на листе
        vsoSelection.IterationMode = visSelModeSkipSub 'только не внутренние шейпы
    End If
    
    'Выбираем к кому приклеиться
    For Each vsoShape In vsoSelection
        If (vsoShape.name <> shpFSAPodval.name) And (vsoShape.name <> shpFSAPodval.Parent.name) And ((ShapeSATypeIs(vsoShape, typeFSAPodval)) Or (vsoShape.name Like "FSAPodvalTab*")) Then
            Set shpToWhichGlue = vsoShape
        End If
    Next
    
    'Клеимся
    If Not shpToWhichGlue Is Nothing Then
        shpFSAPodval.Cells("PinX").FormulaForceU = "GUARD(" & shpToWhichGlue.NameID & "!PinX+" & shpToWhichGlue.NameID & "!Width)"
        shpFSAPodval.Cells("PinY").FormulaForceU = "GUARD(" & shpToWhichGlue.NameID & "!PinY)"
    End If

End Sub

Sub unGlueFSAPodval(shpFSAPodval As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : GlueFSAPodval - Отклеивает канал от подвала или от канала
'------------------------------------------------------------------------------------------------------------
    
    'Отклеиваем
    shpFSAPodval.Cells("PinX").FormulaForce = shpFSAPodval.Cells("PinX").Result(0) + 0.3
    shpFSAPodval.Cells("PinY").FormulaForce = shpFSAPodval.Cells("PinY").Result(0) - 0.3

End Sub


Public Sub DuplicateFSAPodval(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DuplicateFSAPodval - Дублирует канал подвала ФСА
'------------------------------------------------------------------------------------------------------------
    Dim vsoDouble As Visio.Shape
    vsoShape.Cells("User.Dropped").FormulaU = 0 'чтобы не привязывалась мышь
    Set vsoDouble = vsoShape.Duplicate    'дублируем фигуру

        vsoDouble.Cells("PinX").FormulaForceU = "GUARD(" & vsoShape.NameID & "!PinX+" & vsoShape.NameID & "!Width)"
        vsoDouble.Cells("PinY").FormulaForceU = "GUARD(" & vsoShape.NameID & "!PinY)"
        vsoDouble.Cells("Actions.Glue.Invisible").FormulaForceU = 1
        vsoDouble.Cells("Actions.UnGlue.Invisible").FormulaForceU = 0

End Sub
