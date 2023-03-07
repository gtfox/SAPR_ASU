'------------------------------------------------------------------------------------------------------------
' Module        : frmAddWireLink - Форма создания связи (перекрестной ссылки) для разрывов проводов
' Author        : gtfox на основе Shishok::Form_Find
' Date          : 2020.05.24
' Description   : Дочерний элемент выбирает себе родителя через форму
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе этого:
                '------------------------------------------------------------------------------------------------------------
                ' Module    : Form_Find поиск и выделение
                ' Author    : Shishok
                ' Date      : 11.06.2018
                ' Purpose   : Поиск и выделение шейпов по критерию(текст). Для Windows 7 x 32 или типа того
                ' Links     : https://github.com/Shishok/, https://yadi.sk/d/qbpj9WI9d2eqF
                '------------------------------------------------------------------------------------------------------------

'Option Explicit
'Option Base 1

#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#End If
Private Const LVM_FIRST As Long = &H1000   ' 4096
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)   ' 4126
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Dim shpChild As Visio.Shape 'шейп из модуля CrossReference
Dim shpParent As Visio.Shape 'шейп выбанный в форме lstvParent. нужен для создания связи
Dim colShapes As Collection
Dim colPages As Collection
Dim FindType As Integer 'Кто запустил создание связи (родитель/дочерний)
Public pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна перед созданием связи

Sub run(vsoShape As Visio.Shape) 'Приняли шейп из модуля CrossReference
    Set shpChild = vsoShape 'И определили его в форме frmAddReference
    
    FindType = ShapeSAType(shpChild)

    Fill_lstvPages
    Fill_ShapeCollection ActivePage
    Fill_lstvParent

    Call lblHeaders_Click

    lblResult.Caption = "Найдено фигур: " & colShapes.Count
    
    ReSize
    
    frmAddReferenceWireLink.Show
End Sub

Sub Fill_ShapeCollection(vsoPage As Visio.Page) 'Заполняем список с родительскими элементами
    Dim vsoShape As Visio.Shape
    
    If chkAllPages Then
        For Each vsoPage In ActiveDocument.Pages ' перебор страниц документа и шейпов
            For Each vsoShape In vsoPage.Shapes
                SelectType vsoShape, vsoPage
            Next
        Next
    Else
        For Each vsoShape In vsoPage.Shapes ' перебор шейпов на выбранной странице
            SelectType vsoShape, vsoPage
        Next
    End If
    
End Sub

Private Sub SelectType(vsoShape As Visio.Shape, vsoPage As Visio.Page) ' Выбор по типу
    
    If vsoShape.CellExistsU("User.SAType", 0) Then 'отсеиваем посторонние шейпы не имеющие поле ТИП
        Select Case FindType 'Определяемся в соответствии с типом вызвавшего макрос шейпа
            Case typeWireLinkR 'Если макрос активировался дочерним - значит искали родителей
                If ShapeSATypeIs(vsoShape, typeWireLinkS) Then

                    SelectText vsoShape, vsoPage
                End If
            Case typeWireLinkS 'Если макрос активировался родителем - значит искали дочерних
                If ShapeSATypeIs(vsoShape, typeWireLinkR) Then

                    SelectText vsoShape, vsoPage
                End If
        End Select
    End If
   
End Sub

Sub SelectText(vsoShape As Visio.Shape, vsoPage As Visio.Page) ' Выбор - свободные или свободные + занятые
    If optClear Then ' Только не соединенные шейпы
        If vsoShape.CellExistsU("User.LocLink", False) Then ' проверка данных шейпа
            If vsoShape.Cells("User.LocLink").ResultStr(0) = "" Then '
                Call AddToCol(vsoShape, vsoPage)
            End If
        End If
    ElseIf optAll Then 'Все разрывы проводов (соединенные и нет)
        Call AddToCol(vsoShape, vsoPage)
    End If
End Sub

Private Sub AddToCol(vsoShape As Visio.Shape, vsoPage As Visio.Page)  ' добавление элементов в коллекции
    On Error GoTo ExitLine
        colShapes.Add vsoShape.id ' коллекция ID шейпов
        colPages.Add vsoPage.id ' коллекция ID страниц
ExitLine:
End Sub

Sub FindShapes() ' процедура поиска
    Set colShapes = New Collection
    Set colPages = New Collection

    Fill_ShapeCollection ActiveDocument.Pages(lblCurPage.Caption)
    
    If chkAllPages.Value Then
        lblCurPageALL.Visible = True
        lblCurPage.Visible = False
    Else
        lblCurPageALL.Visible = False
        lblCurPage.Visible = True
    End If
    
    If colShapes.Count > 0 Then
        Fill_lstvParent
    Else
        lstvParent.ListItems.Clear
    End If

    lblResult.Caption = "Найдено фигур: " & colShapes.Count
    
    ReSize
    
    Call lblHeaders_Click
    
End Sub



Private Sub ReSize() ' изменение высоты формы. Зависит от количества элементов в listbox
    Dim H As Single
    
    H = lstvPages.ListItems.Count
    If H < lstvParent.ListItems.Count Then H = lstvParent.ListItems.Count

    
    H = H * 12 + 12
    If H < 48 Then H = 48
    If H > 328 Then H = 328
    
    Me.Height = lstvPages.Top + H + 26
    
    lstvPages.Height = H
    lstvParent.Height = H
    
'    H = Me.Height - 35
'
'    Label1.Top = H
'    lblHeaders.Top = H
'    lblContent.Top = H

    
End Sub

Private Sub chkAllPages_Click()

    FindShapes
    
End Sub



Private Sub lstvPages_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку

    With lstvPages
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
    
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
    Dim i As Integer
    Dim itmx As ListItem

    lstvParent.ListItems.Clear
    
    For i = 1 To colShapes.Count  ' добавить N ListItem в коллекцию ListItems
        With ActiveDocument.Pages.ItemFromID(colPages.Item(i)).Shapes.ItemFromID(colShapes.Item(i))
            Set itmx = lstvParent.ListItems.Add(, colPages.Item(i) & "/" & colShapes.Item(i), IIf(.Cells("Prop.Number").Result(0) = 0, "?", .Cells("Prop.Number").Result(0)) & ":" & .Cells("Prop.SymName").ResultStr(0)) 'IIf(.Cells("Prop.Number").Result(0) = 0, "?", .Cells("Prop.Number").Result(0))
                itmx.SubItems(1) = .Cells("User.LocLink").ResultStr(0)
                itmx.SubItems(2) = .Cells("User.Location").ResultStr(0)
                itmx.SubItems(3) = .ContainingPage.name
        End With
    Next i
    
End Sub

Private Sub Fill_lstvPages()   ' заполнение списка страниц
    Dim i As Integer
    Dim itmx As ListItem
    Dim vsoPage As Visio.Page
    
    lstvPages.ListItems.Clear
    
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.PageSheet.CellExistsU("Prop.SA_NazvanieShkafa", 0) Then
            Set itmx = lstvPages.ListItems.Add(, vsoPage.id & "/", vsoPage.name)
        End If
    Next
    
End Sub

Private Sub lstvPages_ItemClick(ByVal Item As MSComctlLib.ListItem)

    chkAllPages.Value = False
    lblCurPage.Caption = Item.text
    lblCurPage.Visible = True
    lblCurPageALL.Visible = False
    
    FindShapes
    
End Sub

Private Sub lstvParent_DblClick()

    Select Case FindType
        Case typeWireLinkR 'Если макрос активировался дочерним - значит искали родителей
            'Создаем связь как и было задумано
            AddReferenceWireLink shpChild, shpParent
        Case typeWireLinkS 'Если макрос активировался родителем - значит искали дочерних
            'Меняем местами родителя/дочернего, т.к. в переменной shpChild содержится родитель, а в shpParent дочерний
            AddReferenceWireLink shpParent, shpChild
    End Select
    
    'Активация событий. Они чета сомодезактивируются xD
'    Set vsoPagesEvent = ActiveDocument.Pages
    
    btnClose_Click
    
End Sub

Private Sub lstvParent_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim vsoShape As Visio.Shape
    Dim ShapeID As String
    Dim PageID As String
    Dim mstrShPgID() As String

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

    lblCurParent.Caption = Item.text + " " + vsoShape.Cells("User.Location").ResultStr(0)

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

Private Sub optAll_Click()

    FindShapes

End Sub

Private Sub optClear_Click()

    FindShapes
    
End Sub

Private Sub UserForm_Initialize() ' инициализация формы
    Set colShapes = New Collection
    Set colPages = New Collection
    
    Me.Height = 213 ' высота формы
    
    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна перед созданием связи
    
    optClear.Caption = "Не связанные"
    optAll.Caption = "Все"
    lblCurParent.Caption = ""
    lblCurPageALL.Caption = "Все страницы"
    lblCurPage.Caption = ActivePage.name
    chkAllPages.Value = False
    
    lstvPages.ColumnHeaders.Add , , "Страницы" ' добавить ColumnHeaders
    'Call SendMessage(lstvPages.hWnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE_USEHEADER) ' выровнять ширину столбцов по заголовкам
    'Call SendMessage(lstvPages.hWnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE) ' выровнять ширину столбцов по содержимому
    lstvPages.ColumnHeaders.Item(1).Width = lstvPages.Width - 18
    
    lstvParent.ColumnHeaders.Add , , "Провод" ' добавить ColumnHeaders
    lstvParent.ColumnHeaders.Add , , "Связь" ' добавить ColumnHeaders
    lstvParent.ColumnHeaders.Add , , "Адрес" ' добавить ColumnHeaders
    lstvParent.ColumnHeaders.Add , , "Страница" ' добавить ColumnHeaders
    
    lstvPages.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvParent.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    
End Sub

Private Sub btnClose_Click() ' выгрузка формы

    With ActiveWindow
        .Page = shpChild.ContainingPage
        .Select shpChild, visDeselectAll + visSelect     ' выделение шейпа
        .SetViewRect pinLeft, pinTop, pinWidth, pinHeight  'Восстановление вида окна после закрытия формы
                    '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End With
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
    
End Sub