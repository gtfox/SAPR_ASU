'Option Explicit
'------------------------------------------------------------------------------------------------------------
' Module        : frmDB - Форма поиска и задания данных для элемента схемы их БД
' Author        : gtfox
' Date          : 2021.02.22
' Description   : Выбор данных из БД прайс листа, фильтрация по категориям и полнотекстовый поиск
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
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

Dim glShape As Visio.Shape 'шейп из модуля DB

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
Dim bBlock As Boolean

Private Sub FiltersCmbxChange()
    Dim SQLQuery As String
    Dim fltrKategoriya As String
    Dim fltrGruppa As String
    Dim fltrPodgruppa As String
    Dim fltrMode As Integer
    Dim fltrWHERE As String
    Dim DBName As String


    If cmbxKategoriya.ListIndex = -1 Then
        fltrKategoriya = ""
    Else
        fltrKategoriya = "Прайс.КатегорииКод=" & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1)
    End If
    If cmbxGruppa.ListIndex = -1 Then
        fltrGruppa = ""
    Else
        fltrGruppa = "Прайс.ГруппыКод=" & cmbxGruppa.List(cmbxGruppa.ListIndex, 1)
    End If
    If cmbxPodgruppa.ListIndex = -1 Then
        fltrPodgruppa = ""
    Else
        fltrPodgruppa = "Прайс.ПодгруппыКод=" & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 1)
    End If
    
    fltrMode = IIf(fltrKategoriya = "", 0, 4) + IIf(fltrGruppa = "", 0, 2) + IIf(fltrPodgruppa = "", 0, 1)

    '    0   0   0
    '    0   0   1
    '    0   1   0
    '    0   1   1
    '    1   0   0
    '    1   0   1
    '    1   1   0
    '    1   1   1
    
    Select Case fltrMode
        Case 0
            fltrWHERE = ""
        Case 1
            fltrWHERE = " WHERE " & fltrPodgruppa
        Case 2
            fltrWHERE = " WHERE " & fltrGruppa
        Case 3
            fltrWHERE = " WHERE " & fltrGruppa & " AND " & fltrPodgruppa
        Case 4
            fltrWHERE = " WHERE " & fltrKategoriya
        Case 5
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrPodgruppa
        Case 6
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrGruppa
        Case 7
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrGruppa & " AND " & fltrPodgruppa
        Case Else
            fltrWHERE = ""
    End Select
    
    SQLQuery = "SELECT Прайс.КодПозиции, Прайс.Артикул, Прайс.Название, Прайс.Цена, Прайс.КатегорииКод, Прайс.ГруппыКод, Прайс.ПодгруппыКод, Прайс.ПроизводительКод " & _
                "FROM Прайс " & fltrWHERE & ";"


    Fill_lstvTablePrice "FilterSQLQuery", SQLQuery
    
    
    DBName = cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1)
    
    If fltrKategoriya = "" Then
        SQLQuery = "SELECT FilterSQLQuery.КатегорииКод, Категории.Категория " & _
                    "FROM Категории INNER JOIN FilterSQLQuery ON Категории.КодКатегории = FilterSQLQuery.КатегорииКод " & _
                    "GROUP BY FilterSQLQuery.КатегорииКод, Категории.Категория;"
        Fill_ComboBox DBName, SQLQuery, cmbxKategoriya
    End If
    
    If fltrGruppa = "" Then
        SQLQuery = "SELECT FilterSQLQuery.ГруппыКод, Группы.Группа " & _
                    "FROM Группы INNER JOIN FilterSQLQuery ON Группы.КодГруппы = FilterSQLQuery.ГруппыКод " & _
                    "GROUP BY FilterSQLQuery.ГруппыКод, Группы.Группа;"
        Fill_ComboBox DBName, SQLQuery, cmbxGruppa
    End If
    
    If fltrPodgruppa = "" Then
        SQLQuery = "SELECT FilterSQLQuery.ПодгруппыКод, Подгруппы.Подгруппа " & _
                    "FROM Подгруппы INNER JOIN FilterSQLQuery ON Подгруппы.КодПодгруппы = FilterSQLQuery.ПодгруппыКод " & _
                    "GROUP BY FilterSQLQuery.ПодгруппыКод, Подгруппы.Подгруппа;"
        Fill_ComboBox DBName, SQLQuery, cmbxPodgruppa
    End If

    ReSize
    
End Sub

Private Sub cmbxKategoriya_Change()
    If Not bBlock Then FiltersCmbxChange
End Sub

Private Sub cmbxGruppa_Change()
    If Not bBlock Then FiltersCmbxChange
End Sub

Private Sub cmbxPodgruppa_Change()
    If Not bBlock Then FiltersCmbxChange
End Sub

Private Sub UserForm_Initialize() ' инициализация формы

    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна перед созданием связи
    lstvTable.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    
    lstvTable.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
    lstvTable.ColumnHeaders.Add , , "Название" ' SubItems(1)
    lstvTable.ColumnHeaders.Add , , "Цена" ' SubItems(2)
    'lstvTable.ColumnHeaders.Add , , "Производитель" ' SubItems(3)

    frameTab.Top = frameFilters.Top + frameFilters.Height
    Me.Height = frameTab.Top + frameTab.Height + 36
    lblResult.Top = Me.Height - 35
    
    tbtnFiltr.Caption = ChrW(9650)
    tbtnBD = True

    Dim SQLQuery As String

    SQLQuery = "SELECT Производители.ИмяФайлаБазы, Производители.Производитель " & _
                "FROM Производители;"
                
    Fill_ComboBox "SAPR_ASU_Izbrannoe.accdb", SQLQuery, cmbxProizvoditel, True

End Sub

Sub Run(vsoShape As Visio.Shape) 'Приняли шейп из модуля DB
    Set glShape = vsoShape 'И определили его как глолбальный в форме frmDB

    'Fill_lstvTablePrice



    'ReSize
    
    frmDB.Show

End Sub



Private Sub ReSize() ' изменение формы. Зависит от длины в lstvTable
    Dim lstvTableWidth As Single
    
    lblHeaders_Click
    
    If lstvTable.ListItems(1).Width > 381 Then
        lstvTableWidth = lstvTable.ListItems(1).Width
    Else
        lstvTableWidth = 381
    End If
    
    lstvTable.Width = lstvTableWidth + 20
    frameTab.Width = lstvTable.Width + 10
    
    frameFilters.Width = frameTab.Width
    Me.Width = frameTab.Width + 14
    cmbxKategoriya.Width = frameFilters.Width - cmbxKategoriya.Left - 6
    cmbxGruppa.Width = frameFilters.Width - cmbxGruppa.Left - 6
    cmbxPodgruppa.Width = frameFilters.Width - cmbxPodgruppa.Left - 6
    btnClose.Left = Me.Width - btnClose.Width - 10
    tbtnFiltr.Left = Me.Width - tbtnFiltr.Width - 10
    btnFavAdd.Left = btnClose.Left - btnFavAdd.Width - 10
    btnETM.Left = btnFavAdd.Left - btnETM.Width - 2
    frameProizvoditel.Width = btnETM.Left - frameProizvoditel.Left - 6
    cmbxProizvoditel.Width = frameProizvoditel.Width - 12
    'lblResult.Top = Me.Height - 35
    lblResult.Left = frameTab.Width - lblResult.Width
    btnFind.Left = frameTab.Width - btnFind.Width - 6
    frameNazvanie.Width = btnFind.Left - frameNazvanie.Left - 6
    txtNazvanie1.Width = (frameNazvanie.Width - 16) / 2
    txtNazvanie2.Left = txtNazvanie1.Left + txtNazvanie1.Width
    txtNazvanie2.Width = frameNazvanie.Width / 4
    txtNazvanie3.Left = txtNazvanie2.Left + txtNazvanie2.Width
    txtNazvanie3.Width = frameNazvanie.Width / 4
'    Me.Hide
'    Me.Show
    
End Sub

Private Sub cmbxProizvoditel_Change()
    Reset_FiltersCmbx
    'Fill_lstvTablePrice
    'ReSize
End Sub

Private Sub tbtnBD_Click()
    tbtnFav = Not tbtnBD
    
End Sub

Private Sub tbtnFav_Click()
    tbtnBD = Not tbtnFav
End Sub

Private Sub lstvTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку

    With lstvTable
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
    
End Sub

Sub Fill_lstvTablePrice(QueryDefName As String, SQLQuery As String) ' заполнение списка найденных позиций в базе
    Dim i As Double
    Dim itmx As ListItem
    Dim dbsDatabase As DAO.Database
    Dim rstRecordset As DAO.Recordset
    Dim strPath As String

    'Создаем набор записей для получения списка
    strPath = ThisDocument.path & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1)
    Set dbsDatabase = GetDBEngine.OpenDatabase(strPath)
    On Error Resume Next
    dbsDatabase.QueryDefs.Delete QueryDefName
    Set rstRecordset = dbsDatabase.CreateQueryDef(QueryDefName, SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей

    lstvTable.ListItems.Clear
    i = 0
    rstRecordset.MoveFirst
    Do Until rstRecordset.EOF
        Set itmx = lstvTable.ListItems.Add(, """" & rstRecordset.Fields("КодПозиции").Value & """", rstRecordset.Fields("Артикул").Value)
        itmx.SubItems(1) = rstRecordset.Fields("Название").Value
        itmx.SubItems(2) = rstRecordset.Fields("Цена").Value
        
        rstRecordset.MoveNext
        i = i + 1
    Loop
    lblResult.Caption = "Найдено записей: " & i
    
    Set dbsDatabase = Nothing
    Set rstRecordset = Nothing
    
    Call lblHeaders_Click ' выровнять ширину столбцов по заголовкам
End Sub


Private Sub Reset_FiltersCmbx()
    Dim DBName As String
    Dim SQLQuery As String
    
    If cmbxProizvoditel.ListIndex = -1 Then Exit Sub
    
    bBlock = True
    
    DBName = cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1)
    
    SQLQuery = "SELECT Категории.КодКатегории, Категории.Категория " & _
                "FROM Категории;"
    Fill_ComboBox DBName, SQLQuery, cmbxKategoriya
    
    SQLQuery = "SELECT Группы.КодГруппы, Группы.Группа " & _
                "FROM Группы;"
    Fill_ComboBox DBName, SQLQuery, cmbxGruppa
        
    SQLQuery = "SELECT Подгруппы.КодПодгруппы, Подгруппы.Подгруппа " & _
                "FROM Подгруппы;"
    Fill_ComboBox DBName, SQLQuery, cmbxPodgruppa
    
    bBlock = False
    
    lstvTable.ListItems.Clear

End Sub

Private Sub lstvTable_DblClick()

    Select Case FindType
        Case typePLCModChild  'Если макрос активировался дочерним - значит искали родителей
            'Создаем связь как и было задумано
            AddReferencePLCMod glShape, shpParent
        Case typePLCIOChild 'Если макрос активировался родителем - значит искали дочерних
            'Меняем местами родителя/дочернего, т.к. в переменной glShape содержится родитель, а в shpParent дочерний
            AddReferencePLCIO glShape, shpParent
    End Select

    'Активация событий. Они чета сомодезактивируются xD
    'Set vsoPagesEvent = ActiveDocument.Pages
    
    btnClose_Click
    
End Sub

Private Sub lstvTable_ItemClick(ByVal Item As MSComctlLib.ListItem)
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
   For colNum = 0 To lstvTable.ColumnHeaders.Count - 1
      Call SendMessage(lstvTable.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvTable.ColumnHeaders.Count - 1
      Call SendMessage(lstvTable.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub

Private Sub tbtnFiltr_Click()
    If tbtnFiltr.Value Then
        frameFilters.Height = 84
        tbtnFiltr.Caption = ChrW(9650) 'вверх
    Else
        frameFilters.Height = 0
        tbtnFiltr.Caption = ChrW(9660) 'вниз
        Reset_FiltersCmbx
    End If
    frameTab.Top = frameFilters.Top + frameFilters.Height
    Me.Height = frameTab.Top + frameTab.Height + 36
    lblResult.Top = Me.Height - 35
End Sub



 Sub btnClose_Click() ' выгрузка формы

    With ActiveWindow
        .Page = glShape.ContainingPage
        .Select glShape, visDeselectAll + visSubSelect     ' выделение шейпа
        .SetViewRect pinLeft, pinTop, pinWidth, pinHeight  'Восстановление вида окна после закрытия формы
                    '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End With
    Unload Me
    
End Sub

