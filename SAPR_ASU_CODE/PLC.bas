'Option Explicit
'------------------------------------------------------------------------------------------------------------
' Module        : PLC - Программируемые логические контроллеры (ПЛК)
' Author        : gtfox
' Date          : 2020.09.11
' Description   : ПЛК и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Sub GenModPLC(vsoModParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : GenModPLC - Заполняет модуль ПЛК

                'Выбираем модуль (PLCModParent) находящийся внутри ПЛК (PLCParent)
                'в котором находится первый вход модуля(одностороннего)
                'или первый правый и первый левый входы двустороннего модуля (2 ряда клемм)
                'и генерируем Вх./Вых. модуля на основе выбранного входа (Первого входа модуля)
                'Остальные Вх./Вых. генерируются автоматически по аналогии с первым
                'Нумерация контактов и номеров входов автоматическая
                'После генерации Вх./Вых. их можно отредактировать вручную
'------------------------------------------------------------------------------------------------------------
    Dim shpPLCIOL As Visio.Shape
    Dim shpPLCIOR As Visio.Shape
    Dim shpPLCIO As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim shpID As Long
    Dim IOName As String '"Название I/O"
    Dim IONumber As Integer 'Номер Вх./Вых.
    Dim ModHeight As Long 'Высота модуля заполненного входами
    
    Dim NIO As Integer 'Количество Вх./Вых. в модуле 1-32
    Dim nColumn As Integer 'Число столбцов клемм: 1 или 2
    
    NIO = vsoModParent.Cells("Prop.NIO").Result(0)
    
    For Each vsoShape In vsoModParent.Shapes
        If ShapeSATypeIs(vsoShape, typePLCIOLParent) Then
            Set shpPLCIOL = vsoShape
            shpPLCIOL.Cells("Prop.Autonum").Formula = True
            nColumn = nColumn + 1
        ElseIf ShapeSATypeIs(vsoShape, typePLCIORParent) Then
            Set shpPLCIOR = vsoShape
            shpPLCIOR.Cells("Prop.Autonum").Formula = True
            nColumn = nColumn + 1
        End If
    Next

    If nColumn = 2 Then
        Set shpPLCIO = ColumnCopy(shpPLCIOL, NIO, nColumn, False, shpPLCIO)
        Set shpPLCIO = ColumnCopy(shpPLCIOR, NIO, nColumn, True, shpPLCIO)
        ModHeight = IIf(shpPLCIOL.Cells("Width").Result(visMillimeters) * (NIO / nColumn) > shpPLCIOR.Cells("Width").Result(visMillimeters) * (NIO / nColumn), shpPLCIOL.Cells("Width").Result(visMillimeters) * (NIO / nColumn), shpPLCIOR.Cells("Width").Result(visMillimeters) * (NIO / nColumn))
    ElseIf nColumn = 1 Then
        If shpPLCIOR Is Nothing Then
            Set shpPLCIO = ColumnCopy(shpPLCIOL, NIO, nColumn, False, shpPLCIO)
            ModHeight = shpPLCIOL.Cells("Width").Result(visMillimeters) * NIO
        Else
            Set shpPLCIO = ColumnCopy(shpPLCIOR, NIO, nColumn, False, shpPLCIO)
            ModHeight = shpPLCIOR.Cells("Width").Result(visMillimeters) * NIO
        End If
    End If
    
    vsoModParent.Cells("Height").Formula = ModHeight & " mm"

End Sub

Function ColumnCopy(shpPLCIO As Visio.Shape, NIO As Integer, nColumn As Integer, r As Boolean, shpPLCIOLast As Visio.Shape) As Visio.Shape
'------------------------------------------------------------------------------------------------------------
' Function      : ColumnCopy - Генерит столбец входов(функция для GenModPLC,GenIOPLC)
'------------------------------------------------------------------------------------------------------------
    Dim shpID As Long
    Dim NPin As Integer '"Число клемм 1-4 на 1 вход"
    Dim i As Integer
    'R-генерим правый столбец
    'shpPLCIOLast-Последний шейп из левого столбца
    
    shpID = shpPLCIO.id
    NPin = shpPLCIO.Cells("Prop.NPin").Result(0)
    
    'Если начало второго столбца - то берем номера клемм из последнего входа левого столбца
    If r Then
        If shpPLCIO.Cells("Prop.IOName").ResultStr(0) Like shpPLCIOLast.Cells("Prop.IOName").ResultStr(0) Then 'одинаковые имена входов
            shpPLCIO.Cells("Prop.IONumber").FormulaU = shpPLCIOLast.Cells("Prop.IONumber").Result(0) + 1
        End If
        shpPLCIO.Cells("User.TNumber1").FormulaU = "sheet." & shpPLCIOLast.id & "!User.LaqstNum+1"
    End If

    For i = 2 To NIO / nColumn
        Set shpPLCIOLast = shpPLCIO
        Set shpPLCIO = shpPLCIO.Duplicate
        NPin = shpPLCIO.Cells("Prop.NPin").Result(0)
        
        shpPLCIO.Cells("Prop.IONumber").FormulaU = shpPLCIOLast.Cells("Prop.IONumber").Result(0) + 1
        shpPLCIO.Cells("User.TNumber1").FormulaU = "sheet." & shpID & "!User.LaqstNum+1"

        shpPLCIO.Cells("PinX").FormulaForceU = "=GUARD(sheet." & shpID & "!PinX)"
        shpPLCIO.Cells("PinY").FormulaForceU = "GUARD(sheet." & shpID & "!PinY-sheet." & shpID & "!Width)"
        
        ClearPLCIOParent shpPLCIO
        
        shpID = shpPLCIO.id

    Next
    Set ColumnCopy = shpPLCIO
End Function

'Активация формы генерации входов
Public Sub dofrmGenIO(shpIO As Visio.Shape) 'Получили шейп с листа
    Load frmGenIO
    frmGenIO.run shpIO 'Передали его в форму
End Sub

Sub GenIOPLC(shpIO As Visio.Shape, NIO As Integer)
'------------------------------------------------------------------------------------------------------------
' Macros        : GenModPLC - Создает входы ПЛК (NIO - кол-во входов)
                'Приклеивает вход ко входу снизу в количестве заданном в форме frmGenIO
'------------------------------------------------------------------------------------------------------------
    Dim shpPLCIO As Visio.Shape
    Dim nColumn As Integer 'Число столбцов клемм: 1 или 2
    
    If shpIO.Parent.Shapes.Count - 2 + NIO > shpIO.Parent.Cells("Prop.NIO").Result(0) Then
        MsgBox "Превышено максимальное число входов модуля", vbInformation, "САПР-АСУ: Info"
    End If
    nColumn = 1
    shpIO.Cells("Prop.Autonum").Formula = True

    Set shpPLCIO = ColumnCopy(shpIO, NIO, nColumn, False, shpPLCIO)

End Sub

Sub GlueIO(shpPLCIO As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : GlueIO - Приклеивает вход снизу от выбранного
                'Нужно выделить вход который хотим приклеить,
                'потом выделить вход к которому хотим приклеить,
                'потом выполнить GlueIO
'------------------------------------------------------------------------------------------------------------
    Dim shpToWhichGlue As Visio.Shape
    Dim vsoSelection As Visio.Selection
    Dim vsoShape As Visio.Shape
    
    Set vsoSelection = Application.ActiveWindow.Selection
    
    'Смотрим где находимся: в группе / на листе
    If shpPLCIO.Parent.Type = visTypeGroup Then 'внутри группы
        vsoSelection.IterationMode = visSelModeOnlySub 'только внутренние шейпы
    ElseIf shpPLCIO.Parent.Type = visTypePage Then 'на листе
        vsoSelection.IterationMode = visSelModeSkipSub 'только не внутренние шейпы
    End If
    
    'Выбираем к кому приклеиться
    For Each vsoShape In vsoSelection
        If (vsoShape.name <> shpPLCIO.name) And (vsoShape.name <> shpPLCIO.Parent.name) And (ShapeSATypeIs(vsoShape, typePLCIOLParent) Or ShapeSATypeIs(vsoShape, typePLCIORParent)) Then
            Set shpToWhichGlue = vsoShape
        End If
    Next
    
    'Связываемся и клеимся
    If Not shpToWhichGlue Is Nothing Then
        'Есил разные (лев/прав) то связываем, но не приклеиваем
        shpToWhichGlue.Cells("Prop.Autonum").Formula = True
        shpPLCIO.Cells("Prop.Autonum").Formula = True
        shpPLCIO.Cells("User.TNumber1").FormulaU = shpToWhichGlue.NameID & "!User.LaqstNum+1"
        If (ShapeSATypeIs(shpToWhichGlue, typePLCIOLParent) And ShapeSATypeIs(shpPLCIO, typePLCIOLParent)) Or (ShapeSATypeIs(shpToWhichGlue, typePLCIORParent) And ShapeSATypeIs(shpPLCIO, typePLCIORParent)) Then
             'Есил одинаковые (лев/прав) то связываем и приклеиваем
             shpPLCIO.Cells("PinX").FormulaForceU = "GUARD(" & shpToWhichGlue.NameID & "!PinX)"
             shpPLCIO.Cells("PinY").FormulaForceU = "GUARD(" & shpToWhichGlue.NameID & "!PinY-" & shpToWhichGlue.NameID & "!Width)"
        End If
    End If

End Sub

Public Sub DuplicateInBox(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DuplicateInBox - Дублирует модуль находящийся внутри ПЛК (Когда копируеш модуль формула в PinY портится)
'------------------------------------------------------------------------------------------------------------
    Dim vsoDouble As Visio.Shape
    Set vsoDouble = vsoShape.Duplicate    'дублируем фигуру

    Select Case ShapeSAType(vsoDouble)

        Case typePLCIOChild, typePLCModParent
            If vsoDouble.Parent.Type = visTypeGroup Then
                If ShapeSATypeIs(vsoDouble.Parent, typeSensor) Or ShapeSATypeIs(vsoDouble.Parent, typeActuator) Then
                    vsoDouble.Cells("PinY").FormulaForceU = "GUARD(" & vsoDouble.Parent.NameID & "!Height*1)"
                Else
                    vsoDouble.Cells("PinY").FormulaForceU = "GUARD(" & vsoDouble.Parent.NameID & "!Height*0)"
                End If
            Else
                vsoDouble.Cells("PinY").FormulaForce = vsoShape.Cells("PinY").Result(0)
            End If
        Case typePLCModChild
            If vsoDouble.Parent.Type = visTypeGroup Then
                vsoDouble.Cells("PinY").FormulaForceU = "GUARD(" & vsoDouble.Parent.NameID & "!Height*Scratch.X1)"
            Else
                vsoDouble.Cells("PinY").FormulaForce = vsoShape.Cells("PinY").Result(0)
            End If
        Case typePLCTerm
            vsoDouble.Cells("PinY").FormulaForce = vsoShape.Cells("PinY").Result(0)
        Case Else
            vsoDouble.Cells("PinY").FormulaForce = vsoShape.Cells("PinY").Result(0)
    End Select
    ActiveWindow.Select vsoDouble, visSubSelect
End Sub
