'------------------------------------------------------------------------------------------------------------
' Module        : MISC - Макросы не относящиеся к другим категориям
' Author        : gtfox
' Date          : 2020.05.05
' Description   : Разные вспомогательные макросы применяющиеся в разных модулях
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------


Dim i As Integer


 
Function GetSAPageExist(PageName As String) As Visio.Page
'------------------------------------------------------------------------------------------------------------
' Function        : GetSAPageExist - Проверяет существование страницы и возвращает ее
'------------------------------------------------------------------------------------------------------------
    On Error GoTo ER
    Set GetSAPageExist = ActiveDocument.Pages.Item(PageName)
    Exit Function
ER:
    Set GetSAPageExist = Nothing
End Function

Function GetSAShapeExist(Container As Object, ShapeName As String) As Visio.Shape
'------------------------------------------------------------------------------------------------------------
' Function        : GetSAShapeExist - Проверяет существование шейпа и возвращает его
'------------------------------------------------------------------------------------------------------------
    On Error GoTo ER
    Set GetSAShapeExist = Container.Shapes(ShapeName)
    Exit Function
ER:
    Set GetSAShapeExist = Nothing
End Function

Function ShapeSAType(vsoShape As Visio.Shape) As Integer
'------------------------------------------------------------------------------------------------------------
' Function        : ShapeSAType - Проверяет существование параметра User.SAType и возвращает его значение
'------------------------------------------------------------------------------------------------------------
    If vsoShape.CellExists("User.SAType", 0) Then   'Если в шейпе есть тип, то -
        ShapeSAType = vsoShape.Cells("User.SAType").Result(0) 'возвращаем его значение
    Else
        ShapeSAType = 0
    End If
End Function

Function ShapeSATypeIs(vsoShape As Visio.Shape, SAType As Integer) As Boolean
'------------------------------------------------------------------------------------------------------------
' Function        : ShapeSATypeIs - Проверяет существование параметра User.SAType и возвращает его соответствие переданному
'------------------------------------------------------------------------------------------------------------
        ShapeSATypeIs = IIf(ShapeSAType(vsoShape) = SAType, True, False)
End Function

Public Function ShapeByHyperLink(HyperLinkToShape As String) As Visio.Shape
'------------------------------------------------------------------------------------------------------------
' Function      : ShapeByHyperLink - Преобразует строку в шейп
                'Строка типа "Схема.3/Sheet.4" разбивается на имя листа и имя шейпа
                'и выдается в качестве объекта-шейпа
                'Если нет ссылки или шейпа на выход идет Nothing
'------------------------------------------------------------------------------------------------------------
    Dim mstrAdrToShape() As String 'массив строк имя страницы и имя шейпа
    
    If HyperLinkToShape <> "" Then 'Если ссылка есть
        'Находим шейп разбивая HyperLinkToShape на имя страницы и имя шейпа
        mstrAdrToShape = Split(HyperLinkToShape, "/")
        On Error GoTo net_takogo_shejpa
        Set ShapeByHyperLink = ActiveDocument.Pages.ItemU(mstrAdrToShape(0)).Shapes(mstrAdrToShape(1))
        Exit Function
    End If
        
net_takogo_shejpa:

    Set ShapeByHyperLink = Nothing
    
End Function

Public Function ShapeByGUID(GUIDToShape As String) As Visio.Shape
'------------------------------------------------------------------------------------------------------------
' Function      : ShapeByGUID - По GUID находит шейп
                'Строка типа "{2287DC42-B167-11CE-88E9-0020AFDDD917}", по ней ищется шейп на всех листах
                'и выдается в качестве объекта-шейпа
                'Если нет ссылки или шейпа на выход идет Nothing
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    
    Set ShapeByGUID = Nothing
    If GUIDToShape <> "" Then 'Если GUID есть
        'Перебираем все листы
        For Each vsoPage In ActiveDocument.Pages
            On Error Resume Next
            Set ShapeByGUID = vsoPage.Shapes.Item("*" & GUIDToShape)
            If Not ShapeByGUID Is Nothing Then Exit Function
        Next
    End If
End Function

Sub SetLocalShkafMesto(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : SetLocalShkafMesto - Задает имя шкафа и место для фигур внутри шейпа "шкаф/место"
'------------------------------------------------------------------------------------------------------------
    Dim selSelection As Visio.Selection
    Dim vsoShp As Visio.Shape
    Dim SAType As Integer
    
    Set selSelection = vsoShape.SpatialNeighbors(visSpatialOverlap + visSpatialTouching + visSpatialContainedIn + visSpatialContain, 0, 0)
    For Each vsoShp In selSelection
        SAType = ShapeSAType(vsoShp)
        If SAType > 1 Then
            Select Case SAType
                Case typeCxemaCoil, typeCxemaParent, typeCxemaElement, typePLCParent, typePLCModParent, typeCxemaTerm, typeCxemaWire
                    vsoShp.Cells("User.Shkaf").FormulaU = "Pages[" & vsoShape.ContainingPage.NameU & "]!" & vsoShape.NameID & "!Prop.SA_NazvanieShkafa"
                    vsoShp.Cells("User.Mesto").FormulaU = "Pages[" & vsoShape.ContainingPage.NameU & "]!" & vsoShape.NameID & "!Prop.SA_NazvanieMesta"
                Case typeCxemaActuator, typeCxemaSensor
                    vsoShp.Cells("User.Shkaf").FormulaU = """"""
                    vsoShp.Cells("User.Mesto").FormulaU = "Pages[" & vsoShape.ContainingPage.NameU & "]!" & vsoShape.NameID & "!Prop.SA_NazvanieMesta"
                Case typeCxemaCable
                    vsoShp.Cells("User.Shkaf").FormulaU = """"""
                    vsoShp.Cells("User.Mesto").FormulaU = """"""
            End Select
        End If
    Next
End Sub

Sub DeleteShkafMesto(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteShkafMesto - Возвращает имя шкафа и место для фигур на схеме
'------------------------------------------------------------------------------------------------------------
    Dim selSelection As Visio.Selection
    Dim vsoShp As Visio.Shape
    Dim SAType As Integer
    
    Set selSelection = vsoShape.SpatialNeighbors(visSpatialOverlap + visSpatialTouching + visSpatialContainedIn + visSpatialContain, 0, 0)
    For Each vsoShp In selSelection
        ClearShkafMesto vsoShp
    Next
End Sub

Sub ClearShkafMesto(vsoShp As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearShkafMesto - "Чистит" имя шкафа и места
'------------------------------------------------------------------------------------------------------------
    Dim SAType As Integer
    SAType = ShapeSAType(vsoShp)
    If SAType > 1 Then
        Select Case SAType
            Case typeCxemaCoil, typeCxemaParent, typeCxemaElement, typePLCParent, typePLCModParent, typeCxemaTerm, typeCxemaWire
                vsoShp.Cells("User.Shkaf").FormulaU = "ThePage!Prop.SA_NazvanieShkafa"
                vsoShp.Cells("User.Mesto").FormulaU = "ThePage!Prop.SA_NazvanieMesta"
            Case typeCxemaActuator, typeCxemaSensor
                vsoShp.Cells("User.Shkaf").FormulaU = """"""
                vsoShp.Cells("User.Mesto").FormulaU = "ThePage!Prop.SA_NazvanieMesta"
            Case typeCxemaCable
                vsoShp.Cells("User.Shkaf").FormulaU = """"""
                vsoShp.Cells("User.Mesto").FormulaU = """"""
        End Select
    End If
End Sub

Sub ResetLocalShkafMesto(vsoObject As Object)
'------------------------------------------------------------------------------------------------------------
' Macros        : ResetLocalShkafMesto - Обновляет имя шкафа и место для фигур внутри шейпов "шкаф/место" и снаружи от них
'------------------------------------------------------------------------------------------------------------
    Dim selSelection As Visio.Selection
    Dim vsoShp As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim SAType As Integer
    Dim colElementyShemy As Collection
    Dim colShkafyMesta As Collection
    
    Set colElementyShemy = New Collection
    Set colShkafyMesta = New Collection
    
    If 1 Then 'vsoObject.Type = visTypePage Then
        'Заполняем коллекции эелементов и шкафов для всего проекта
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.name Like cListNameCxema & "*" Then
                For Each vsoShp In vsoPage.Shapes
                    SAType = ShapeSAType(vsoShp)
                    Select Case SAType
                        Case typeCxemaCoil, typeCxemaParent, typeCxemaElement, typePLCParent, typeCxemaTerm, typeCxemaActuator, typeCxemaSensor
                            colElementyShemy.Add vsoShp
                        Case typeCxemaShkafMesto
                            colShkafyMesta.Add vsoShp
                    End Select
                Next
            End If
        Next
    Else
        'Заполняем коллекции эелементов и шкафов для страницы
        For Each vsoShp In ActivePage.Shapes
            SAType = ShapeSAType(vsoShp)
            Select Case SAType
                Case typeCxemaCoil, typeCxemaParent, typeCxemaElement, typePLCParent, typeCxemaTerm, typeCxemaActuator, typeCxemaSensor
                    colElementyShemy.Add vsoShp
                Case typeCxemaShkafMesto
                    colShkafyMesta.Add vsoShp
            End Select
        Next
    End If

    'Чистим все элементы
    For Each vsoShp In colElementyShemy
        ClearShkafMesto vsoShp
    Next
    
    'Обновляем все шкафы
    For Each vsoShp In colShkafyMesta
        SetLocalShkafMesto vsoShp
    Next
    
'    ActiveWindow.SelectAll
'    Load frmMenuReNumber
'    frmMenuReNumber.ReNumberShemy
'    ActiveWindow.DeselectAll

    ActiveWindow.DeselectAll
    Load frmMenuReNumber
    frmMenuReNumber.obVseCx = True
    frmMenuReNumber.obVseTipObCx = True
    frmMenuReNumber.ReNumberShemy
    
    UpdateNazvanieShkafa
End Sub

Public Sub ObjInfo()
'------------------------------------------------------------------------------------------------------------
' Macros        : ObjInfo - Показывает информацию о выделенном шейпе, субшейпе или странице на форме frmMenuObjInfo
                'Вызывается кнопкой на панели инструментов САПР АСУ
'------------------------------------------------------------------------------------------------------------
    Dim vsoSelection As Visio.Selection
   
    Set vsoSelection = Application.ActiveWindow.Selection
    Load frmMenuObjInfo
    If ActiveWindow.Selection.Count > 1 Then frmMenuObjInfo.Caption = "Info " + "(выделено " + CStr(ActiveWindow.Selection.Count) + ")"
    If vsoSelection.PrimaryItem Is Nothing Then
        vsoSelection.IterationMode = visSelModeOnlySub
        'For Each sh In vsoSelection
            If vsoSelection.PrimaryItem Is Nothing Then
                frmMenuObjInfo.run ActivePage
            Else
                frmMenuObjInfo.run vsoSelection.PrimaryItem
            End If
        'Next
    Else
        frmMenuObjInfo.run vsoSelection.PrimaryItem
    End If
End Sub

Sub UngroupNumSelect()
    Dim vsoSelection As Visio.Selection
    Application.AlertResponse = 1
    ActiveWindow.Selection.Ungroup
    Set vsoSelection = ActiveWindow.Selection
    Application.AlertResponse = 0
    ResetLocalShkafMesto ActivePage
    ActiveWindow.Selection = vsoSelection
End Sub

Sub Duplicate()
    ActiveWindow.Selection.Copy visCopyPasteNoTranslate
    ActivePage.Paste visCopyPasteNoTranslate
End Sub

Sub OnlyGroup()
    Dim vsoShape As Visio.Shape
    If ActiveWindow.Selection.Count > 0 Then
        For Each vsoShape In ActiveWindow.Selection
            vsoShape.CellsSRC(visSectionObject, visRowGroup, visGroupSelectMode).FormulaU = "0"
        Next
    End If
End Sub

Sub BeginGroup()
    Dim vsoShape As Visio.Shape
    If ActiveWindow.Selection.Count > 0 Then
        For Each vsoShape In ActiveWindow.Selection
            vsoShape.CellsSRC(visSectionObject, visRowGroup, visGroupSelectMode).FormulaU = "1"
        Next
    End If
End Sub

Sub MenuAddToStencilFrm()
    Load frmMenuAddToStencil
    frmMenuAddToStencil.Show
End Sub

Sub AddCxemaToStencil(NameStencil As String, NameMaster As String)
    Dim vsoDocument As Visio.Document
    Dim vsoStencilCopy As Visio.Master
    Dim vsoLayer1 As Visio.Layer
    Dim vsoShape As Visio.Shape
    Dim vsoMaster As Visio.Master
    Dim iPos As Integer
    
    If ActiveWindow.Selection.Count > 1 Then
    
        On Error GoTo AddNewStencil
        Set vsoDocument = Application.Documents.Item(NameStencil)
        err.Clear
        On Error GoTo 0
        vsoDocument.Save
        vsoDocument.Close
        Set vsoDocument = Application.Documents.OpenEx(ActiveDocument.path & NameStencil, visOpenRW + visOpenDocked)
        ActiveWindow.Selection.Copy visCopyPasteNoTranslate
        Set vsoLayer1 = Application.ActiveWindow.Page.Layers.Add("temp")
        vsoLayer1.CellsC(visLayerActive).FormulaU = "1"
        DoEvents
        Application.EventsEnabled = 0
        ActivePage.Paste visCopyPasteNoTranslate
        Set vsoShape = ActiveWindow.Selection.Group
        vsoShape.Cells("EventDrop").FormulaU = "SETF(GetRef(PinY)," & Replace(CStr(vsoShape.Cells("PinY").Result(0)), ",", ".") & "/ThePage!PageScale*ThePage!DrawingScale) + SETF(GetRef(PinX)," & Replace(CStr(vsoShape.Cells("PinX").Result(0)), ",", ".") & "/ThePage!PageScale*ThePage!DrawingScale) + RunMacro(""UngroupNumSelect"")"
        Set vsoMaster = vsoDocument.Drop(vsoShape, 0, 0)
        If NameMaster <> "" Then
            vsoMaster.name = NameMaster
        End If
        vsoDocument.Save
        vsoShape.Delete
        vsoLayer1.Delete 1
        Application.EventsEnabled = -1
'            MsgBox "Не выбран набор элементов", vbOKOnly + vbInformation, "САПР-АСУ: Info"
    Else
        MsgBox "Выделите больше одного элемента схемы", vbOKOnly + vbInformation, "САПР-АСУ: Info"
    End If
    
    Exit Sub
    
AddNewStencil:
    Set vsoDocument = Application.Documents.AddEx("vss", visMSMetric, visAddDocked + visAddStencil, 1033)
    If NameStencil <> "" Then
        iPos = InStrRev(NameStencil, ".")
        If iPos > 0 Then
            NameStencil = Left(NameStencil, iPos - 1)
        End If
        NameStencil = NameStencil & ".vss"
        vsoDocument.SaveAs ActiveDocument.path & NameStencil
    Else
        NameStencil = vsoDocument & ".vss"
        vsoDocument.SaveAs ActiveDocument.path & NameStencil
    End If
Resume

End Sub

Sub SetUserSAType(SAType As Integer)
    Dim vsoShape As Visio.Shape
    For Each vsoShape In ActiveWindow.Selection 'ActiveSelection
        vsoShape.Cells("User.SAType").Formula = SAType
    Next
End Sub

Sub SetUserSAType_0()
    SetUserSAType 0
End Sub

Sub SetUserSAType_132()
    SetUserSAType 132
End Sub

'------------------------------------------------------------------------------------------------------------
' Macros        : ExtractOboz - Функция определения неизменяемой части обозначения
' Author        : Shishok
' Date          : 2014.12.01
' Description   : Определения неизменяемой части обозначения Например: 1, ГР1, р, Гр1.1, ППР1-1, Выкл, П122.1 или типа того
' Link          : https://visio.getbb.ru/viewtopic.php?p=5904#p5904, https://github.com/shishok, https://disk.yandex.ru/d/qbpj9WI9d2eqF
'------------------------------------------------------------------------------------------------------------
Function ExtractOboz(Oboz) ' Функция определения неизменяемой части обозначения

Dim ObozF As String, i As Integer, Flag As Boolean
Flag = Oboz Like "*[-.,/\]*"

For i = 1 To Len(Oboz)
    If Not Flag And Mid(Oboz, i, 1) Like "[a-zA-Zа-яА-Я ]" Then GoSub AddChar
    If Flag And Mid(Oboz, i, 1) Like "[a-zA-Zа-яА-Я0-9 ]" Then GoSub AddChar
    If Flag And Mid(Oboz, i, 1) Like "[-.,/\]" Then GoSub AddChar
Next
    
ExtractOboz = ObozF
Exit Function

AddChar:
    ObozF = ObozF + Mid(Oboz, i, 1)
Return
End Function


'    ReDim arrRowValue(10, 1)
'    arrRowValue = [{"1", "2";"11", "22";"111", "222"}]
'    UBarrCellName = UBound(arrRowValue)

'Sub ЦБР()
'    Dim str As String
'    Dim xmDoc As Object
'
'    Set xmDoc = CreateObject("msxml2.DOMDocument")
'    xmDoc.async = 0
'    xmDoc.Load ("http://www.cbr.ru/scripts/XML_daily.asp")
'    With xmDoc.SelectSingleNode("*/Valute[CharCode='USD']")
'        str = CDbl(.ChildNodes(4).Text) / Val(.ChildNodes(2).Text)
'    End With
'    Set xmDoc = Nothing
'End Sub



'    For Each vsoShape In ActivePage.Shapes
'        n = vsoShape.LayerCount
'        If n > 0 Then
'            For i = 1 To n
'                Set vsoShapeLayer = vsoShape.Layer(i)
'                If vsoShapeLayer.Name = vsoLayer.Name Then
'
'                End If
'            Next
'        End If
'    Next


'
'Private Sub mcr1() 'добавление панельки
'
'Set cbar1 = Application.CommandBars.Add(Name:="Custom1", Position:=msoBarFloating)
'cbar1.Visible = True
'
'
'Set myControl = cbar1.Controls _
'    .Add(Type:=msoControlComboBox, Before:=1)
'With myControl
'    .AddItem Text:="First Item", Index:=1
'    .AddItem Text:="Second Item", Index:=2
'    .DropDownLines = 3
'    .DropDownWidth = 75
'    .ListHeaderCount = 0
'    .OnAction = "SAPR_ASU.LockTitleBlock"
'
'End With
'
'End Sub

'Private Sub ttt()
'Dim List As tList
'List = A4m
'
'    Select Case List
'        Case tList.A4m
'            ' Process.
'            List = A3b1
'        Case tList.A4b
'            ' Process.
'        Case tList.A3b1
'            ' Process.
'        Case Else
'
'    End Select
'
'End Sub





'Sub ReadCopyRight()
'    MsgBox ActiveWindow.Selection(1).Cells("Copyright").FormulaU
'End Sub
'Sub RegCopyright()
'    On Error GoTo EMSG
'    ActiveWindow.Selection(1).Cells("Copyright").FormulaU = Chr(34) & "Copyright (C) 2009 Visio Guys" & Chr(34)
'    Exit Sub
'EMSG:
'    MsgBox err.Description
'End Sub
'Sub RegAllCopyright()
'    Dim shp As Visio.Shape
'    On Error GoTo EMSG
'    For Each shp In ActivePage.Shapes
'        shp.Cells("Copyright").FormulaU = Chr(34) & "Copyright (C) 2009 Visio Guys" & Chr(34)
'    Next
'    Exit Sub
'EMSG:
'    MsgBox err.Description
'End Sub


'Sub CopyEventsDisabled()
'
'    Application.ActiveWindow.Selection.Copy
'    Application.EventsEnabled = 0
'    Application.ActivePage.Paste
'    DoEvents
'    Application.EventsEnabled = -1
'End Sub

'Public Enum tList
'    A4m = 1
'    A4b = 2
'    A3m1 = 3
'    A3m2 = 4
'    A3b1 = 5
'    A3b2 = 6
'End Enum
