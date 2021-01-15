

Option Explicit
'------------------------------------------------------------------------------------------------------------
' Module        : frmAddReferencePLCMod - Форма создания связей (перекрестных ссылок) модулей внутри PLC
' Author        : gtfox на основе Shishok::Form_Find
' Date          : 2020.09.14
' Description   : Дочерний элемент выбирает себе родителя через форму
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе этого:
                '------------------------------------------------------------------------------------------------------------
                ' Module    : Form_Find поиск и выделение
                ' Author    : Shishok
                ' Date      : 11.06.2018
                ' Purpose   : Поиск и выделение шейпов по критерию(текст). Для Windows 7 x 32 или типа того
                ' Links     : https://github.com/Shishok/, https://yadi.sk/d/qbpj9WI9d2eqF
                '------------------------------------------------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#End If
Private Const LVM_FIRST As Long = &H1000   ' 4096
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)   ' 4126
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Dim shpChild As Visio.Shape 'шейп из модуля CrossReferencePLCMod
Dim shpParent As Visio.Shape 'шейп выбанный в форме lstvParent. нужен для создания связи
Dim colShapes As Collection
Dim colPages As Collection
Dim FindType As Integer 'Кто запустил создание связи (родитль/дочерний)
Public pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна перед созданием связи
Dim HyperLinkToParentPLC As String
Dim HyperLinkToParentPLCMod As String
Dim mstrAdrParentPLC() As String
Dim mstrAdrParentPLCMod() As String
Dim shpParentPLC As Visio.Shape 'Родительский плк с модулями
Dim shpParentPLCMod As Visio.Shape 'Родительский модуль со входами внутри  родительского плк
Dim vsoShp As Visio.Shape
Dim bError As Boolean

Sub Run(vsoShape As Visio.Shape) 'Приняли шейп из модуля CrossReferencePLCMod
    Set shpChild = vsoShape 'И определили его в форме frmAddReferencePLCMod
    
    FindType = shpChild.Cells("User.SAType").Result(0)

    FillCollection shpChild
    
    Select Case FindType
        Case typePLCModChild 'Если макрос активировался дочерним PLCModChild - значит искали родителей PLCModParent
            lstvParent.ColumnHeaders.Add , , "Модуль" ' добавить ColumnHeaders
            lstvParent.ColumnHeaders.Add , , "Назв." ' SubItems(1)
            lstvParent.ColumnHeaders.Add , , "Связ.мод." ' SubItems(2)
            lstvParent.ColumnHeaders.Add , , "Макс./Сущ.вх." ' SubItems(3)
            lstvParent.ColumnHeaders.Add , , "Связ./Подкл.вх." ' SubItems(4)

        Case typePLCIOChild  'Если макрос активировался дочерним PLCIO - значит искали PLCIO
            lstvParent.ColumnHeaders.Add , , "Входы" ' добавить ColumnHeaders
            lstvParent.ColumnHeaders.Add , , "Связи" ' добавить ColumnHeaders
            lstvParent.ColumnHeaders.Add , , "Провода" ' добавить ColumnHeaders
    End Select
    
            lstvChild.Visible = False
            lstvPages.Visible = False
            'lblResult.Left = 230
            'btnClose.Left = 230
            'Me.Width = 340
    
    
    Fill_lstvParent

    Call lblHeaders_Click ' выровнять ширину столбцов по заголовкам
    
    lblResult.Caption = "Найдено фигур: " & colShapes.Count
    
    ReSize
    
    If bError Then
        Unload Me
    Else
        frmAddReferencePLCMod.Show
    End If
    
End Sub

Private Sub FillCollection(vsoShape As Visio.Shape)
        
        Select Case FindType 'Определяемся в соответствии с типом вызвавшего макрос шейпа
            
            Case typePLCModChild 'Если макрос активировался дочерним PLCModChild - значит искали родителей PLCModParent
                
                HyperLinkToParentPLC = vsoShape.Parent.CellsU("Hyperlink.PLC.SubAddress").ResultStr(0)
                If HyperLinkToParentPLC <> "" Then 'Если ссылка есть
                    'Находим родителя разбивая HyperLink на имя страницы и имя шейпа
                    mstrAdrParentPLC = Split(HyperLinkToParentPLC, "/")
                    'On Error GoTo netu_roditelya 'вдруг его уже удалили и ссылку забыли почистить
                    Set shpParentPLC = ActiveDocument.Pages.ItemU(mstrAdrParentPLC(0)).Shapes(mstrAdrParentPLC(1))
                    
                    lblPLC.Caption = "ПЛК: " & shpParentPLC.CellsU("User.Name").ResultStr(0) & "   Модель: " & shpParentPLC.CellsU("Prop.Model").ResultStr(0)

                    For Each vsoShp In shpParentPLC.Shapes
                        If vsoShp.Name Like "PLCModParent*" Then
                            colShapes.Add vsoShp.ID
                            colPages.Add vsoShp.ContainingPage.ID
                        End If
                    Next
                Else
                    MsgBox "Не привязан ПЛК", vbOKOnly + vbExclamation, "Info"
                    bError = True
                End If

            Case typePLCIOChild 'Если макрос активировался дочерним PLCIOChild - значит искали PLCIOParent
                HyperLinkToParentPLCMod = vsoShape.Parent.CellsU("Hyperlink.PLCMod.SubAddress").ResultStr(0)
                If HyperLinkToParentPLCMod <> "" Then 'Если ссылка есть
                    'Находим родителя разбивая HyperLink на имя страницы и имя шейпа
                    mstrAdrParentPLCMod = Split(HyperLinkToParentPLCMod, "/")
                    'On Error GoTo netu_roditelya 'вдруг его уже удалили и ссылку забыли почистить
                    Set shpParentPLCMod = ActiveDocument.Pages.ItemU(mstrAdrParentPLCMod(0)).Shapes(mstrAdrParentPLCMod(1))
                    
                    lblPLC.Caption = "ПЛК: " & shpParentPLCMod.Parent.CellsU("User.Name").ResultStr(0) & "   Модель: " & shpParentPLCMod.Parent.CellsU("Prop.Model").ResultStr(0)
                    lblPLCMod.Caption = "Модуль: " & shpParentPLCMod.CellsU("User.Name").ResultStr(0) & "   Модель: " & shpParentPLCMod.CellsU("Prop.Model").ResultStr(0)
                    
                    For Each vsoShp In shpParentPLCMod.Shapes
                        If vsoShp.Name Like "PLCIO*" Then
                            colShapes.Add vsoShp.ID
                            colPages.Add vsoShp.ContainingPage.ID
                        End If
                    Next
                Else
                    MsgBox "Не привязан модуль ПЛК", vbOKOnly + vbExclamation, "Info"
                    bError = True
                End If

        End Select

End Sub

Private Sub ReSize() ' изменение высоты формы. Зависит от количества элементов в listbox
    Dim H As Single
    
    H = lstvParent.ListItems.Count
  
    H = H * 12 + 12
    If H < 48 Then H = 48
    If H > 328 Then H = 328
    
    Me.Height = lstvPages.Top + H + 26

    lstvParent.Height = H
    
End Sub

Private Sub lstvParent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку

    With lstvParent
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
    
End Sub

Sub Fill_lstvParent() ' заполнение списка родительских элементов схемы
    Dim i, j, x, y, n, k As Integer
    Dim itmx As ListItem
    Dim wires As String
    Dim vsoShape As Visio.Shape
    lstvParent.ListItems.Clear
    
    Select Case FindType
        Case typePLCModChild  'Если макрос активировался дочерним PLCModChild - значит искали родителей PLCModParent
            For i = 1 To colShapes.Count  ' добавить N ListItem в коллекцию ListItems
                With ActiveDocument.Pages.ItemFromID(colPages.Item(i)).Shapes.ItemFromID(colShapes.Item(i))
                    Set itmx = lstvParent.ListItems.Add(, colPages.Item(i) & "/" & colShapes.Item(i), .Cells("User.Name").ResultStr(0)) '.Cells("TheText").ResultStr("")
                    itmx.SubItems(1) = .Cells("Prop.Model").ResultStr(0)
                    'подсчет кол-ва связей модуля
                    k = 0
                    For n = 1 To .Section(visSectionHyperlink).Count
                        k = k + IIf(.CellsU("Hyperlink." & n & ".SubAddress").ResultStr(0) = "", 0, 1)
                    Next
                    itmx.SubItems(2) = k
                    itmx.SubItems(3) = .Cells("Prop.NIO").Result(0) & "  |  " & .Shapes.Count - 1
                    x = 0
                    y = 0
                    For Each vsoShape In .Shapes
                        If (vsoShape.Name Like "PLCIOL*") Or (vsoShape.Name Like "PLCIOR*") Then
                            'подсчет кол-ва связанных входов
                            x = x + IIf(vsoShape.CellsU("Hyperlink.IO.SubAddress").ResultStr(0) <> "", 1, 0)
                            'подсчет кол-ва подключенных входов
                            For n = 1 To 4
                                If vsoShape.Cells("User.w" & n).Result(0) <> 0 Then
                                    y = y + 1
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                    itmx.SubItems(4) = x & "  |  " & y
                    
              End With
            Next i
        Case typePLCIOChild 'Если макрос активировался дочерним PLCIO - значит искали PLCIO
            For i = 1 To colShapes.Count  ' добавить N ListItem в коллекцию ListItems
                With ActiveDocument.Pages.ItemFromID(colPages.Item(i)).Shapes.ItemFromID(colShapes.Item(i))
                    Set itmx = lstvParent.ListItems.Add(, colPages.Item(i) & "/" & colShapes.Item(i), .Cells("User.Name").ResultStr(0)) '.Cells("TheText").ResultStr("")
                    itmx.SubItems(1) = .CellsU("Hyperlink.IO.ExtraInfo").ResultStr(0)
                    wires = IIf(.Cells("User.w1").Result(0) <> 0, .Cells("User.w1").Result(0), "")
                    For j = 2 To 4
                        wires = IIf(.Cells("User.w" & j).Result(0) <> 0, wires & ", " & .Cells("User.w" & j).Result(0), wires & "")
                    Next j
                    itmx.SubItems(2) = wires
                    wires = ""
                End With
            Next i
    End Select

End Sub

Private Sub lstvParent_DblClick()

    Select Case FindType
        Case typePLCModChild  'Если макрос активировался дочерним - значит искали родителей
            'Создаем связь как и было задумано
            AddReferencePLCMod shpChild, shpParent
        Case typePLCIOChild 'Если макрос активировался родителем - значит искали дочерних
            'Меняем местами родителя/дочернего, т.к. в переменной shpChild содержится родитель, а в shpParent дочерний
            AddReferencePLCIO shpChild, shpParent
    End Select

    'Активация событий. Они чета сомодезактивируются xD
    'Set vsoPagesEvent = ActiveDocument.Pages
    
    btnClose_Click
    
End Sub

Private Sub lstvParent_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim vsoShape As Visio.Shape
    Dim ShapeID As String
    Dim PageID As String
    Dim mstrShPgID() As String
    
    'lblCurParent.Caption = Item.Text
    
    mstrShPgID = Split(Item.Key, "/")
    PageID = mstrShPgID(0)   ' ID страницы
    ShapeID = mstrShPgID(1)   ' ID шейпа

    With ActiveWindow
        .Page = ActiveDocument.Pages.ItemFromID(PageID) ' активация нужной страницы
        Set vsoShape = ActivePage.Shapes.ItemFromID(ShapeID)
        If vsoShape.Parent.Type = visTypeGroup Then
            .Select vsoShape, visDeselectAll + visSubSelect  ' выделение субшейпа
            '.CenterViewOnShape ActivePage.Shapes(shName), visCenterViewSelectShape '2010+
        Else
            .Select vsoShape, visDeselectAll + visSelect     ' выделение шейпа
            '.CenterViewOnShape ActivePage.Shapes(shName) , visCenterViewSelectShape '2010+
            .SetViewRect vsoShape.Cells("PinX") - pinWidth / 2, vsoShape.Cells("PinY") + pinHeight / 2, pinWidth, pinHeight
            '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
        End If
    End With

    If vsoShape.CellExistsU("User.Location", 0) Then
        'lblCurParent.Caption = Item.Text + "  " + vsoShape.Cells("User.Location").ResultStr(0)
    End If
    

    ReSize
    
    Set shpParent = vsoShape 'передаем родителя для создания связи
    
End Sub

Private Sub lblContent_Click() ' выровнять ширину столбцов по содержимому
   Dim colNum As Long
   For colNum = 0 To lstvParent.ColumnHeaders.Count - 1
      Call SendMessage(lstvParent.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvParent.ColumnHeaders.Count - 1
      Call SendMessage(lstvParent.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub

Private Sub UserForm_Initialize() ' инициализация формы
    Set colShapes = New Collection
    Set colPages = New Collection
    
    Me.Height = 213 ' высота формы
    
    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна перед созданием связи
    
    lblPLC.Caption = ""
    lblPLCMod.Caption = ""

    lstvParent.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке

End Sub

 Sub btnClose_Click() ' выгрузка формы

    With ActiveWindow
        .Page = shpChild.ContainingPage
        .Select shpChild, visDeselectAll + visSubSelect     ' выделение шейпа
        .SetViewRect pinLeft, pinTop, pinWidth, pinHeight  'Восстановление вида окна после закрытия формы
                    '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End With
    Unload Me
    
End Sub