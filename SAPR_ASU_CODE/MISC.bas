Dim i As Integer
Sub CopyEventsDisabled()

    Application.ActiveWindow.Selection.Copy
    Application.EventsEnabled = 0
    Application.ActivePage.Paste
    DoEvents
    Application.EventsEnabled = -1
End Sub


Public Sub WireInfoPlan(Connects As IVConnects)
'------------------------------------------------------------------------------------------------------------
' Macros        : WireInfoPlan - Заполняет выноску на плане (лоток и кабели)
                'Клеим выноску на лоток, по которому проложены кабели
                'Выноска заполняется названием лотка и номерами кабелей
'------------------------------------------------------------------------------------------------------------
    Dim shpVynoska As Visio.Shape
    Dim shpTouchingShapes As Visio.Shape
    Dim vsoSelection As Visio.Selection
    Dim strProvoda As String
    Dim strLotok As String
    Dim colNum As Collection
    Dim mNum() As Integer
    Dim NumTemp As Variant
    Dim i As Integer
    Dim j As Integer
    Dim UbNum As Long
    
    Set colNum = New Collection
    Set shpVynoska = Connects.FromSheet
    strProvoda = "("
    
    Select Case shpVynoska.Connects.Count 'кол-во соединенных концов у выноски
        Case 0 'Нет соединенных концов - отцепили
                shpVynoska.Cells("Prop.Lotok").FormulaU = """"""
                shpVynoska.Cells("Prop.Provoda").FormulaU = """"""
        Case 1, 2 'С одной стороны
            Set vsoSelection = shpVynoska.ContainingPage.SpatialSearch(shpVynoska.Cells("EndX").Result(0), shpVynoska.Cells("EndY").Result(0), visSpatialTouching, 0.02, 0)
            For Each shpTouchingShapes In vsoSelection
                Debug.Print shpTouchingShapes.Name
                If shpTouchingShapes.Name Like "w*" Then
                    colNum.Add shpTouchingShapes.Cells("Prop.Number").Result(0)
                ElseIf shpTouchingShapes.Name Like "Lotok*" Then
                    strLotok = shpTouchingShapes.Cells("User.FullName").ResultStr(0)
                End If
            Next

        'Case 2 'С двух сторон - не обрабатываем 2-ю сторону
    End Select
    If colNum.Count = 0 Then Exit Sub
    'из коллекции передаем номера проводов в массив для сортировки
    ReDim mNum(colNum.Count - 1)
    i = 0
    For Each NumTemp In colNum
        mNum(i) = NumTemp
        i = i + 1
    Next
    
    ' "Сортировка вставками" номеров проводов
    '--V--Сортируем по возрастанию номеров проводов
    UbNum = UBound(mNum)
    For j = 1 To UbNum
        NumTemp = mNum(j)
        i = j
        While mNum(i - 1) > NumTemp '>:возрастание, <:убывание
            mNum(i) = mNum(i - 1)
            i = i - 1
            If i <= 0 Then GoTo ExitWhileX
        Wend
ExitWhileX: mNum(i) = NumTemp
    Next
    '--Х--Сортировка по возрастанию номеров проводов
 
    For i = 0 To UbNum
        strProvoda = strProvoda & mNum(i) & ";"
    Next
                    
    strProvoda = Left(strProvoda, Len(strProvoda) - 1)
    If Len(strProvoda) > 1 Then
        strProvoda = strProvoda & ")"
    End If

    shpVynoska.Cells("Prop.Lotok").FormulaU = """" & strLotok & """"
    shpVynoska.Cells("Prop.Provoda").FormulaU = """" & strProvoda & """"

End Sub

'Public Enum tList
'    A4m = 1
'    A4b = 2
'    A3m1 = 3
'    A3m2 = 4
'    A3b1 = 5
'    A3b2 = 6
'End Enum



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
'    Debug.Print ActiveWindow.Selection(1).Cells("Copyright").FormulaU
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


