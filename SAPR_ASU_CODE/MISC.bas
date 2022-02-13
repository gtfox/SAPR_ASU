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



Sub ObjInfo()
'------------------------------------------------------------------------------------------------------------
' Macros        : ObjInfo - Показывает информацию о выделенном шейпе, субшейпе или странице на форме frmObjInfo
                'Вызывается кнопкой на панели инструментов САПР АСУ
'------------------------------------------------------------------------------------------------------------
    Dim vsoSelection As Visio.Selection
   
    Set vsoSelection = Application.ActiveWindow.Selection
    Load frmObjInfo
    If ActiveWindow.Selection.Count > 1 Then frmObjInfo.Caption = "Info " + "(выделено " + CStr(ActiveWindow.Selection.Count) + ")"
    If vsoSelection.PrimaryItem Is Nothing Then
        vsoSelection.IterationMode = visSelModeOnlySub
        'For Each sh In vsoSelection
            If vsoSelection.PrimaryItem Is Nothing Then
                frmObjInfo.run ActivePage
            Else
                frmObjInfo.run vsoSelection.PrimaryItem
            End If
        'Next
    Else
        frmObjInfo.run vsoSelection.PrimaryItem
    End If
End Sub

Private Sub Tune_Stencils() 'переделка шаблонов электры под гост (перед выполнением макроса надо окрыть шаблоны и сделать их редактируемыми)

    Dim appdoc As Document
    Dim appcol As Collection
    Set appcol = New Collection
    Dim mast As Master
    Dim ss As String
        
    'выбираем нужные шаблоны для измениния
    For Each appdoc In Application.Documents
        If (appdoc.Creator = "Electra" Or appdoc.Creator = "Pneumata" Or appdoc.Creator = "Hydraula") And Not (appdoc.Title = "Electra" Or appdoc.Title = "Layout" Or appdoc.Title = "Layout 3D" Or appdoc.Title = "Reports" Or appdoc.Title = "IEC Parts" Or appdoc.Title = "Title Blocks") Then
            appcol.Add appdoc
        End If
    Next
    
    For Each appdoc In appcol
        For Each mast In appdoc.Masters
            If InStr(1, mast.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageScale).FormulaU, "in") Then 'не трогаем элемент если он в мм (значит он уже был изменён)
                
                'масштаб под гост
                mast.Shapes(1).Cells("Width").FormulaForceU = "guard(" & str(mast.Shapes(1).Cells("Width").Result(visInches) * 1.181102362) & ")"
                mast.Shapes(1).Cells("Height").FormulaForceU = "guard(" & str(mast.Shapes(1).Cells("Height").Result(visInches) * 1.181102362) & ")"
                
                If mast.Shapes(1).Shapes.Count > 0 Then
                    'скрываем описание
                    On Error Resume Next
                    mast.Shapes(1).Shapes("Desc").CellsU("HideText").FormulaU = "TRUE"
                    'поворот фигур
                    mast.Shapes(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormAngle).FormulaU = "=IF(Actions.Row_2.Action,-90 deg,0 deg)"
                    mast.Shapes(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormFlipX).FormulaU = 0
                    'только группа
                    mast.Shapes(1).CellsSRC(visSectionObject, visRowGroup, visGroupSelectMode).FormulaU = "0"
                End If
                
                'страница в милиметрах чтобы электра не запускала конвертацию in->mm
                mast.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageScale).FormulaU = "1 mm"
                mast.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawingScale).FormulaU = "1 mm"
                
            End If
        Next mast
        appdoc.Save
    Next appdoc

End Sub

'-----------------------------Переделка таблицы спецификации под универсальную---------------------------------
Sub TuneTable_1()
    Dim shpRow As Visio.Shape
    Dim shpCel As Visio.Shape
    For i = 1 To 30
        Set shpRow = ActivePage.Shapes("Спецификация").Shapes("row" & i)
        shpRow.Shapes(i & "." & 1).Cells("Width").FormulaU = "=Sheet.65!Width"
        shpRow.Shapes(i & "." & 1).Cells("PinX").FormulaU = "=Sheet.65!PinX"
        shpRow.Shapes(i & "." & 2).Cells("Width").FormulaU = "=Sheet.57!Width"
        shpRow.Shapes(i & "." & 2).Cells("PinX").FormulaU = "=Sheet.57!PinX"
        shpRow.Shapes(i & "." & 3).Cells("Width").FormulaU = "=Sheet.64!Width"
        shpRow.Shapes(i & "." & 3).Cells("PinX").FormulaU = "=Sheet.64!PinX"
        shpRow.Shapes(i & "." & 4).Cells("Width").FormulaU = "=Sheet.62!Width"
        shpRow.Shapes(i & "." & 4).Cells("PinX").FormulaU = "=Sheet.62!PinX"
        shpRow.Shapes(i & "." & 5).Cells("Width").FormulaU = "=Sheet.61!Width"
        shpRow.Shapes(i & "." & 5).Cells("PinX").FormulaU = "=Sheet.61!PinX"
        shpRow.Shapes(i & "." & 6).Cells("Width").FormulaU = "=Sheet.60!Width"
        shpRow.Shapes(i & "." & 6).Cells("PinX").FormulaU = "=Sheet.60!PinX"
        shpRow.Shapes(i & "." & 7).Cells("Width").FormulaU = "=Sheet.63!Width"
        shpRow.Shapes(i & "." & 7).Cells("PinX").FormulaU = "=Sheet.63!PinX"
        shpRow.Shapes(i & "." & 8).Cells("Width").FormulaU = "=Sheet.59!Width"
        shpRow.Shapes(i & "." & 8).Cells("PinX").FormulaU = "=Sheet.59!PinX"
        shpRow.Shapes(i & "." & 9).Cells("Width").FormulaU = "=Sheet.58!Width"
        shpRow.Shapes(i & "." & 9).Cells("PinX").FormulaU = "=Sheet.58!PinX"
        shpRow.Shapes(i & "." & 10).Cells("Width").FormulaU = "=Sheet.367!Width"
        shpRow.Shapes(i & "." & 10).Cells("PinX").FormulaU = "=Sheet.367!PinX"
        For j = 1 To 10
            Set shpCel = shpRow.Shapes(i & "." & j)
            shpCel.Cells("PinY").FormulaU = shpRow.NameID & "!Height*0"
            shpCel.Cells("LocPinX").FormulaU = "=Width*0"
            shpCel.Cells("LocPinY").FormulaU = "=Height*0"
        Next
    Next
End Sub

Sub TuneTable_2()
    Dim shpRow As Visio.Shape
    Dim shpCel As Visio.Shape
    For i = 1 To 30
        Set shpRow = ActivePage.Shapes("Спецификация").Shapes("row" & i)
        shpRow.Cells("Height").FormulaForceU = Replace(shpRow.Cells("Height").FormulaU, "))", "," & shpRow.Shapes(i & ".10").NameID & "!User.Row_1))")
    Next
End Sub

Sub TuneTable_3()
    Dim shpCel As Visio.Shape
    For i = 1 To 10
        If i < 10 Then
            Set shpCel = ActivePage.Shapes("Спецификация").Shapes("Head").Shapes("0" & i)
        Else
            Set shpCel = ActivePage.Shapes("Спецификация").Shapes("Head").Shapes("10")
        End If
        With shpCel
            .AddSection visSectionFirstComponent
            .AddRow visSectionFirstComponent, visRowComponent, visTagComponent
            .AddRow visSectionFirstComponent, visRowVertex, visTagLineTo
            .AddRow visSectionFirstComponent, visRowVertex, visTagMoveTo
            .CellsSRC(visSectionFirstComponent, 0, 0).FormulaForceU = "TRUE"
            .CellsSRC(visSectionFirstComponent, 0, 1).FormulaForceU = "FALSE"
            .CellsSRC(visSectionFirstComponent, 0, 2).FormulaForceU = "FALSE"
            .CellsSRC(visSectionFirstComponent, 0, 3).FormulaForceU = "FALSE"
            .CellsSRC(visSectionFirstComponent, 1, 0).FormulaU = "Width*0"
            .CellsSRC(visSectionFirstComponent, 1, 1).FormulaU = "Height*0"
            .CellsSRC(visSectionFirstComponent, 2, 0).FormulaU = "Width*1"
            .CellsSRC(visSectionFirstComponent, 2, 1).FormulaU = "Height*0"
            .AddRow visSectionFirstComponent, 3, visTagLineTo
            .CellsSRC(visSectionFirstComponent, 3, 0).FormulaU = "Width*1"
            .CellsSRC(visSectionFirstComponent, 3, 1).FormulaU = "Height * 1"
            .AddRow visSectionFirstComponent, 4, visTagLineTo
            .CellsSRC(visSectionFirstComponent, 4, 0).FormulaU = "Width*0"
            .CellsSRC(visSectionFirstComponent, 4, 1).FormulaU = "Height*1"
            .AddRow visSectionFirstComponent, 5, visTagLineTo
            .CellsSRC(visSectionFirstComponent, 5, 0).FormulaU = "Width*0"
            .CellsSRC(visSectionFirstComponent, 5, 1).FormulaU = "Geometry1.Y1"
        End With
    Next
End Sub

Sub TuneTable_4()
    Dim shpCel As Visio.Shape
    For i = 1 To 10
        If i < 10 Then
            Set shpCel = ActivePage.Shapes("Спецификация").Shapes("Head").Shapes("0" & i)
        Else
            Set shpCel = ActivePage.Shapes("Спецификация").Shapes("Head").Shapes("10")
        End If
        shpCel.Cells("Width").FormulaU = "=Sheet.47!Width*Sheet.45!Prop.S" & i & "/Sheet.45!Prop.Width"
    Next
End Sub

'-----------------------------------------------------------------------------------------------


Public Sub dl()
Dim sel As Selection
Dim snap1 As Shape
Set sel = ActiveWindow.Selection
If sel.Count <> 1 Then ' если не выделено ничего или больше одного будет сообщение
        MsgBox "Нужно выделить лишь одну линию!"
Exit Sub
End If
Set snap1 = sel.Item(1)
Dim dl As Double
dl = CableLength(snap1)
MsgBox ("длина линии " & dl & " м")
End Sub

Sub UngroupThis(shpObj As Visio.Shape)
'Автоматическая разгруппировка фигур при вбросе
'http://visguy.com/vgforum/index.php?topic=26.0
'CALLTHIS("UngroupThis")
'DOCMD(1052) разгруппирует фигуру
On Error GoTo A
'Respond OK to all messages
Application.AlertResponse = 1
'Ungroup the shape
shpObj.Ungroup
A:
'Stop auto responding to messages
'When macro fails settings will be restored to Visio default
Application.AlertResponse = 0
End Sub

Sub test_vss()
Dim vsoShape As Visio.Master
For Each vsoShape In Application.Documents.Item("SAPR_ASU_vid.vss").Masters
q = vsoShape.Name
w = vsoShape.NameU
'ActivePage.Drop vsoShape, 0, 0
'vsoShape.Delete
'vsoShape.Shapes("DIN").CellsU("FillPattern").Formula = "USE(""Dinrejka"")"
Next
End Sub


Sub edit_vss()
Dim vsoMaster As Visio.Master
Dim vsoShape As Visio.Shape
Dim nameShape As String
nameShape = "PanelMAX"
'nameShape = "Panel"
'For Each vsoMaster In Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters(nameShape).Shapes(nameShape).Shapes
For Each vsoMaster In Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters
On Error GoTo err1
Set vsoShape = vsoMaster.Shapes(vsoMaster.Name)
'q = vsoShape.Name
'w = vsoShape.NameU
'ActivePage.Drop vsoShape, 0, 0
'If (vsoShape.Name Like "DIN*") Or (vsoShape.Name Like "KabKan*") Then
'vsoShape.CellsU("Prop.Dlina").FormulaU = "FORMAT(Width,""0u"")"
N = 3
q = vsoShape.CellsSRC(visSectionAction, N, visActionMenu).ResultStr(0)
vsoShape.CellsSRC(visSectionAction, N, visActionMenu).RowNameU = "AddDB"
'End If
err1:
Set vsoMaster = Nothing
Set vsoShape = Nothing
q = ""
Next
End Sub




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
