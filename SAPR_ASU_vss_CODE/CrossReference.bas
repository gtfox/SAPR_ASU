Attribute VB_Name = "CrossReference"
'------------------------------------------------------------------------------------------------------------
' Module        : CrossReference - Перекрестные ссылки элементов схемы
' Author        : gtfox
' Date          : 2020.05.17
' Description   : Перекрестные ссылки элементов схемы и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit


'Активация формы создания связи элементов схемы
Public Sub AddReferenceFrm(shpChild As Visio.Shape) 'Получили шейп с листа
    Load frmAddReference
    frmAddReference.Run shpChild 'Передали его в форму
End Sub

'Активация формы создания связи разрывов проводов
Public Sub AddReferenceWireLinkFrm(shpChild As Visio.Shape) 'Получили шейп с листа
    Load frmAddReferenceWireLink
    frmAddReferenceWireLink.Run shpChild 'Передали его в форму
End Sub


Sub AddLocThumb(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddLocThumb - Добавляет миниатюры контактов под реле
                'Вставляет под катушку реле миниатюры всех активных контактов
                'Вставляет под контакт реле миниатюру катушки реле
                
                'Вызов макроса из меню = CALLTHIS("CrossReference.AddLocThumb","SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim shpThumb As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim vsoMaster As Visio.Master
    Dim DeltaX As Single
    Dim DeltaY As Single
    Dim dN As Single 'смещение миниатюр по вертикали
    Dim i As Integer
    Dim n As Integer 'число контактов в катушке
    
    DeltaX = 0.295275590551181
    DeltaY = -0.246062992125984
    dN = -9.84251968503937E-02
    
    Set vsoPage = ActivePage
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_SHAPE.vss").Masters.Item("Thumb")
    
    
    If vsoShape.CellExistsU("User.Type", 0) Then
        'Выясняем кому надо вставить миниатюры
        Select Case vsoShape.Cells("User.Type").Result(0)
        
            Case typeNO, typeNC 'Контакты
            
                If vsoShape.Cells("Hyperlink.Coil.SubAddress").ResultStr(0) <> "" Then
                    'Вставляем миниатюру контакта Thumbnail
                    Set shpThumb = vsoPage.Drop(vsoMaster, vsoShape.Cells("PinX").Result(0), vsoShape.Cells("PinY").Result(0))
                    'Заполняем поля
                    shpThumb.Cells("User.LocType").FormulaU = typeCoil
                    shpThumb.Cells("User.Location").FormulaU = vsoShape.NameU & "!User.LocationParent"
                    shpThumb.Cells("User.AdrSource").FormulaU = Chr(34) & vsoShape.ContainingPageID & "/" & vsoShape.ID & Chr(34)
                    shpThumb.Cells("User.DeltaX").FormulaU = Chr(34) & DeltaX & Chr(34) 'shpThumb.Cells("PinX").ResultStrU("in")
                    shpThumb.Cells("User.DeltaY").FormulaU = Chr(34) & DeltaY & Chr(34) 'shpThumb.Cells("PinY").ResultStrU("in")
                    shpThumb.Cells("PinX").FormulaU = "=SETATREF(User.DeltaX,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.ID & "!PinX))+Sheet." & vsoShape.ID & "!PinX"
                    shpThumb.Cells("PinY").FormulaU = "=SETATREF(User.DeltaY,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.ID & "!PinY))+Sheet." & vsoShape.ID & "!PinY"
                End If
                
            Case typeCoil 'Катушка реле
            
                n = 0
                'Перебираем активные ссылки на контакты
                For i = 1 To vsoShape.Section(visSectionScratch).Count 'Ищем строку в Scratch
                    If vsoShape.CellsU("Scratch.A" & i).ResultStr(0) <> "" Then 'не пустая строка
                        'Вставляем миниатюру контакта Thumbnail
                        Set shpThumb = vsoPage.Drop(vsoMaster, vsoShape.Cells("PinX").Result(0), vsoShape.Cells("PinY").Result(0))
                        'Заполняем поля
                        shpThumb.Cells("User.LocType").FormulaU = vsoShape.NameU & "!Scratch.D" & i
                        shpThumb.Cells("User.Location").FormulaU = vsoShape.NameU & "!Scratch.C" & i
                        shpThumb.Cells("User.AdrSource").FormulaU = Chr(34) & vsoShape.ContainingPageID & "/" & vsoShape.ID & Chr(34)
                        shpThumb.Cells("User.DeltaX").FormulaU = Chr(34) & DeltaX & Chr(34) 'shpThumb.Cells("PinX").ResultStrU("in")
                        shpThumb.Cells("User.DeltaY").FormulaU = Chr(34) & (DeltaY + n * dN) & Chr(34) 'shpThumb.Cells("PinY").ResultStrU("in")
                        shpThumb.Cells("PinX").FormulaU = "=SETATREF(User.DeltaX,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.ID & "!PinX))+Sheet." & vsoShape.ID & "!PinX"
                        shpThumb.Cells("PinY").FormulaU = "=SETATREF(User.DeltaY,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.ID & "!PinY))+Sheet." & vsoShape.ID & "!PinY"
                        n = n + 1
                    End If
                Next
        End Select
    End If
    ActiveWindow.DeselectAll
End Sub





Sub AddReference(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReference - Создает связь между дочерним и родительским элементом

                'После выбора дочернего(контакт)/родительского(катушка) элемента заполняем небходимые поля для каждого из них
                'Имя(Sheet.4), Страница(Схема.3), Путь(Pages[Схема.3]!Sheet.4), Ссылка(HyperLink="Схема.3/Sheet.4"),
                'Тип контакта(NO/NC), Местоположение(/14.E7), Номер контакта(KL1.3)
                'Ссылки на контакты в катушке формируются формулами в ShapeSheet
                'Нумерация контактов автоматическая, формулами в Scratch.B1-B4 катушки
                'Контактов у катушки может быть 4
'------------------------------------------------------------------------------------------------------------
    'Dim shpParent As Visio.Shape
    Dim shpParentOld As Visio.Shape
    'Dim shpChild As Visio.Shape
    Dim PageParent, NameIdParent, AdrParent As String
    Dim PageChild, NameIdChild, AdrChild As String
    Dim i As Integer
    Dim HyperLinkToChild As String
    Dim HyperLinkToParentOld As String
    Dim mstrAdrParentOld() As String
    
    'Set shpChild = ActivePage.Shapes("Sheet.72") 'для отладки
    'Set shpParent = ActivePage.Shapes("Sheet.7") 'для отладки

    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild
    HyperLinkToChild = PageChild + "/" + NameIdChild

    'Проверяем текущую привязку контакта к старой катушке и чистим ее в старой катушке
    HyperLinkToParentOld = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
    If HyperLinkToParentOld <> "" Then 'Если ссылка есть - значит мы привязаны к родителю
        'Находим родителя разбивая HyperLink на имя страницы и имя шейпа
        mstrAdrParentOld = Split(HyperLinkToParentOld, "/")
        On Error GoTo netu_roditelya 'вдруг его уже удалили и ссылку забыли почистить
        Set shpParentOld = ActiveDocument.Pages(mstrAdrParentOld(0)).Shapes(mstrAdrParentOld(1))
        'Ищем строку в Scratch катушки(родителя) с адресом удаляемого контакта (дочернего)
        For i = 1 To shpParentOld.Section(visSectionScratch).Count
            If shpParentOld.CellsU("Scratch.A" & i).ResultStr(0) = HyperLinkToChild Then
                'Чистим родительский шейп
                shpParentOld.CellsU("Scratch.A" & i).FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
                shpParentOld.CellsU("Scratch.C" & i).FormulaForceU = ""
                shpParentOld.CellsU("Scratch.D" & i).FormulaForceU = ""
                Exit For
            End If
        Next
    End If
netu_roditelya:
    'Привязываем контакт к новой катушке
    For i = 1 To shpParent.Section(visSectionScratch).Count 'Ищем первую не заполненную строку в Scratch
        If shpParent.CellsU("Scratch.A" & i).ResultStr(0) <> "" Then
            If i = shpParent.Section(visSectionScratch).Count Then 'Последняя строка заполнена
                'нет свободных мест
            End If
        Else 'нашли первую не заполненную строку в Scratch
        
            'Заполняем родительский шейп
            shpParent.CellsU("Scratch.A" & i).FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
            shpParent.CellsU("Scratch.C" & i).FormulaU = AdrChild + "!User.Location"   'Pages[Схема.3]!Sheet.4!User.Location
            shpParent.CellsU("Scratch.D" & i).FormulaU = AdrChild + "!User.Type"  'Pages[Схема.3]!Sheet.4!User.Type
            
            'Заполняем дочерний шейп
            shpChild.Cells("Prop.AutoNum").FormulaU = True 'Переводим в автонумерацию
            shpChild.CellsU("User.NameParent").FormulaU = AdrParent + "!User.Name"  'Pages[Схема.3]!Sheet.4!User.Name
            shpChild.CellsU("User.Number").FormulaU = AdrParent + "!Scratch.B" + CStr(i) 'Pages[Схема.3]!Sheet.4!Scratch.B2
            shpChild.CellsU("User.LocationParent").FormulaU = AdrParent + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location
            
            If shpChild.CellExistsU("HyperLink.Coil", False) = False Then
               shpChild.AddNamedRow visSectionHyperlink, "Coil", 0
               shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkDescription).FormulaU = """Катушка ""&User.NameParent&"": ""&User.LocationParent"
            End If
            shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"
            
            Exit For
        End If
    Next

End Sub

Sub DeleteChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteChild - Удаляет дочерний элемент
                'Если контакт привязан, находим родителя, чистим его от удаляемого, и удаляем.
                'Удаляем миниатюру катушки, если она была
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape
    'Dim shpChild As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim shpThumb As Visio.Shape
    Dim colThumb As Collection
    Dim mstrAdrParent() As String
    Dim HyperLinkToParent As String
    Dim HyperLinkToChild As String
    Dim PageChild, NameIdChild As String
    Dim i As Integer
    
    Set colThumb = New Collection
    
    'Set shpChild = ActivePage.Shapes("Sheet.1") 'для отладки
    
    HyperLinkToParent = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
    
    'Проверяем что контакт привязан к катушке
    If HyperLinkToParent <> "" Then
    
        'Находим родителя разбивая HyperLink на имя страницы и имя шейпа
        mstrAdrParent = Split(HyperLinkToParent, "/")
        Set shpParent = ActiveDocument.Pages(mstrAdrParent(0)).Shapes(mstrAdrParent(1))
    
        PageChild = shpChild.ContainingPage.NameU
        NameIdChild = shpChild.NameID
        HyperLinkToChild = PageChild + "/" + NameIdChild
        
        'Ищем строку в Scratch катушки(родителя) с адресом удаляемого контакта (дочернего)
        For i = 1 To shpParent.Section(visSectionScratch).Count
            If shpParent.CellsU("Scratch.A" & i).ResultStr(0) <> HyperLinkToChild Then
                If i = shpParent.Section(visSectionScratch).Count Then 'Последняя строка не соответствует
                    'не нашли контакт в катушке
                End If
            Else 'нашли в Scratch адрес удаляемого контакта
            
                'Чистим родительский шейп
                shpParent.CellsU("Scratch.A" & i).FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
                shpParent.CellsU("Scratch.C" & i).FormulaForceU = ""
                shpParent.CellsU("Scratch.D" & i).FormulaForceU = ""
                
                'Удаляем дочерний шейп
                'shpChild.Delete
                
                Exit For
            End If
        Next
    Else
        'Удаляем контакт не связанный с катушкой  - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete
        'shpChild.Delete
        
        
    End If
    
    'Собираем миниатюры контактов, если они были, в коллекцию для удаления
    For Each vsoShape In ActivePage.Shapes
        If vsoShape.CellExistsU("User.Type", 0) Then
            If vsoShape.Cells("User.Type").Result(0) = typeThumb Then
                Set shpThumb = vsoShape
                If shpThumb.Cells("User.AdrSource").ResultStr(0) = shpChild.ContainingPage.ID & "/" & shpChild.ID Then
                    colThumb.Add shpThumb
                End If
            End If
        End If
    Next
    'Удаляем найденные контакты
    For Each shpThumb In colThumb
        shpThumb.Delete
    Next
    Set colThumb = Nothing
    
End Sub

Sub DeleteParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteParent - Удаляет родительский элемент
                'Смотрим ссылки в родительском, идем по ним и чистим дочерние, потом удаляем родителя.
                'Удаляем миниатюры контактов, если они были
'------------------------------------------------------------------------------------------------------------
    'Dim shpParent As Visio.Shape
    Dim shpChild As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim shpThumb As Visio.Shape
    Dim colThumb As Collection
    Dim mstrAdrChild() As String
    Dim HyperLinkToParent As String
    Dim HyperLinkToChild As String
    Dim LinkPlaceParent As String
    Dim PageParent, NameIdParent As String
    Dim i As Integer
    
    Set colThumb = New Collection
    
    'Set shpParent = ActivePage.Shapes("Sheet.6") 'для отладки
    
    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    LinkPlaceParent = PageParent + "/" + NameIdParent 'Для проверки ссылки в дочернем

    'Ищем строки в Scratch катушки(родителя) с адресами удаляемых контактов (дочерних)
    For i = 1 To shpParent.Section(visSectionScratch).Count
        HyperLinkToChild = shpParent.CellsU("Scratch.A" & i).ResultStr(0)
        If HyperLinkToChild <> "" Then 'нашли в Scratch адрес удаляемого контакта
            
            'Находим контакт разбивая HyperLink на имя страницы и имя шейпа
            mstrAdrChild = Split(HyperLinkToChild, "/")
            Set shpChild = ActiveDocument.Pages(mstrAdrChild(0)).Shapes(mstrAdrChild(1))
            'В контакте находим ссылку на катушку
            HyperLinkToParent = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
            
            'Проверяем что контакт привязан именно к нашей катушке
            If HyperLinkToParent = LinkPlaceParent Then
                'Чистим дочерний шейп
                shpChild.CellsU("User.NameParent").FormulaU = ""
                shpChild.CellsU("User.Number").FormulaU = ""
                shpChild.CellsU("User.LocationParent").FormulaU = ""
                shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """"""
            End If
        End If
    Next
    
    'Почистили все дочерние. Удаляем родителя. - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete
    'shpParent.Delete
    
    'Собираем миниатюры контактов, если они были, в коллекцию для удаления
    For Each vsoShape In ActivePage.Shapes
        If vsoShape.CellExistsU("User.Type", 0) Then
            If vsoShape.Cells("User.Type").Result(0) = typeThumb Then
                Set shpThumb = vsoShape
                If shpThumb.Cells("User.AdrSource").ResultStr(0) = shpParent.ContainingPage.ID & "/" & shpParent.ID Then
                   colThumb.Add shpThumb
                End If
            End If
        End If
    Next
    'Удаляем найденные контакты
    For Each shpThumb In colThumb
        shpThumb.Delete
    Next
    Set colThumb = Nothing
    
End Sub

'Sub ClearReferenceEvent(vsoShapeEvent As Visio.Shape)
''------------------------------------------------------------------------------------------------------------
'' Macros        : ClearReferenceEvent - Чистит дочерний при копировании
'                'Чистим ссылки в дочернем при его копировании.
'                'В EventDrop должна быть формула = CALLTHIS("ThisDocument.ClearReferenceEvent")
'                'Этот макрос расположен в ThisDocument
''------------------------------------------------------------------------------------------------------------
'    Set vsoWindowEvent = ActiveWindow
'    Set vsoShapePaste = vsoShapeEvent
'    Click = False
'    ClearReference vsoShapePaste
'End Sub

Sub ClearReference(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearReference - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReference.ClearReference", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим дочерний шейп
    shpChild.CellsU("User.NameParent").FormulaForceU = ""
    shpChild.CellsU("User.Number").FormulaForceU = ""
    shpChild.CellsU("User.LocationParent").FormulaForceU = ""
    shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaForceU = """"""

End Sub

Sub GoHyperLink(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : GoHyperLink - Переходит по ссылке в разрыве провода
                'Переходит по ссылке в разрыве провода по двойному клику
                
                'Вызов макроса в EventDblClick  =CALLTHIS("CrossReference.GoHyperLink","SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim shpTarget As Visio.Shape
    Dim HyperLinkToTarget As String
    Dim mstrAdrTarget() As String
'    Dim pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна
    
'    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна

    'Находим шейп-цель для последующего выделения
    HyperLinkToTarget = vsoShape.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
    If HyperLinkToTarget <> "" Then
        mstrAdrTarget = Split(HyperLinkToTarget, "/")
        On Error GoTo netu_celi
        Set shpTarget = ActiveDocument.Pages(mstrAdrTarget(0)).Shapes(mstrAdrTarget(1))
        'Переходим по ссылке
        vsoShape.Hyperlinks("1").Follow
        ActiveWindow.DeselectAll
'        ActiveWindow.SetViewRect shpTarget.Cells("PinX") - pinWidth / 2, shpTarget.Cells("PinY") + pinHeight / 2, pinWidth, pinHeight
        ActiveWindow.Select shpTarget, visSelect
    End If

netu_celi:
End Sub

Sub AddReferenceWireLink(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferenceWireLink - Создает связь между шейпами разрывов проводов

                'После выбора дочернего/родительского элемента заполняем небходимые поля для каждого из них
                'Номер провода Prop.Number(5), Название провода Prop.Name("24V"),Местоположение User.LocLink (/14.E7), Ссылка(HyperLink="Схема.3/Sheet.4"),
                'У одного родителя может быть и дочерний (связь 1:1)
'------------------------------------------------------------------------------------------------------------
    'Dim shpParent As Visio.Shape
    Dim shpParentOld As Visio.Shape
    'Dim shpChild As Visio.Shape
    Dim shpChildOld As Visio.Shape
    Dim PageParent, NameIdParent, AdrParent As String
    Dim PageChild, NameIdChild, AdrChild As String
    Dim i As Integer
    Dim HyperLinkToChild As String
    Dim HyperLinkToParentOld As String
    Dim mstrAdrParentOld() As String
    Dim HyperLinkToChildOld As String
    Dim mstrAdrChildOld() As String
    
    'Set shpChild = ActivePage.Shapes("Sheet.72") 'для отладки
    'Set shpParent = ActivePage.Shapes("Sheet.7") 'для отладки

    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild
    HyperLinkToChild = PageChild + "/" + NameIdChild

    'Проверяем текущую привязку разрыва провода(дочернего) к старому разрыву(родильскому) и чистим его в старом разрыве.
    
    'А еще в старом рарыве была вторая половинка - старый дочерний. Его тоже чистим.
    
    HyperLinkToParentOld = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
    If HyperLinkToParentOld <> "" Then 'Если ссылка есть - значит мы привязаны к родителю
        'Находим родителя разбивая HyperLink на имя страницы и имя шейпа
        mstrAdrParentOld = Split(HyperLinkToParentOld, "/")
        On Error GoTo netu_roditelya 'вдруг его уже удалили и ссылку забыли почистить
        Set shpParentOld = ActiveDocument.Pages(mstrAdrParentOld(0)).Shapes(mstrAdrParentOld(1))
        'Чистим родительский шейп
        shpParentOld.CellsU("User.LocLink").FormulaU = """"""
        shpParentOld.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
   
        
        'Находим подключенный к новому родителю дочерний шейп (если он есть)
        HyperLinkToChildOld = shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
        If HyperLinkToChildOld <> "" Then
            mstrAdrChildOld = Split(HyperLinkToChildOld, "/")
            On Error GoTo netu_dochernego 'вдруг его уже удалили и ссылку забыли почистить
            Set shpChildOld = ActiveDocument.Pages(mstrAdrChildOld(0)).Shapes(mstrAdrChildOld(1))
                
            'Чистим дочерний шейп
            shpChildOld.CellsU("Prop.Number").FormulaU = ""
            shpChildOld.CellsU("Prop.Name").FormulaU = """"""
            shpChildOld.CellsU("User.LocLink").FormulaU = """"""
            shpChildOld.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """"""
        End If
netu_dochernego:

 End If
netu_roditelya:

    'Привязываемся к новому разрыву провода
    
    'Заполняем родительский шейп
    shpParent.CellsU("User.LocLink").FormulaU = AdrChild + "!User.Location"  'Pages[Схема.3]!Sheet.4!User.Location
    shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
    
    'Заполняем дочерний шейп
    shpChild.CellsU("Prop.Number").FormulaU = AdrParent + "!Prop.Number"
    shpChild.CellsU("Prop.Name").FormulaU = AdrParent + "!Prop.Name"
    shpChild.CellsU("User.LocLink").FormulaU = AdrParent + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location
    shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"


End Sub

Sub DeleteChildWireLink(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteChildWireLink - Удаляет дочерний элемент
                'Если разрыв провода привязан, находим родителя, чистим его от удаляемого, и удаляем.
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape
    'Dim shpChild As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim mstrAdrParent() As String
    Dim HyperLinkToParent As String
    Dim i As Integer
    
    'Set shpChild = ActivePage.Shapes("Sheet.1") 'для отладки
    
    HyperLinkToParent = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
    
    'Проверяем что разрыв провода привязан родителю
    If HyperLinkToParent <> "" Then
    
        'Находим родителя разбивая HyperLink на имя страницы и имя шейпа
        mstrAdrParent = Split(HyperLinkToParent, "/")
        On Error GoTo netu_roditelya 'вдруг его уже удалили и ссылку забыли почистить
        Set shpParent = ActiveDocument.Pages(mstrAdrParent(0)).Shapes(mstrAdrParent(1))
            
        'Чистим родительский шейп
        shpParent.CellsU("User.LocLink").FormulaU = """"""
        shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
    
    End If
    
netu_roditelya:
    'Удаляем дочерний шейп - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete
End Sub

Sub DeleteParentWireLink(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteParentWireLink - Удаляет родительский элемент
                'Смотрим ссылки в родительском, идем по ним и чистим дочерние, потом удаляем родителя.
'------------------------------------------------------------------------------------------------------------
    'Dim shpParent As Visio.Shape
    Dim shpChild As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim mstrAdrChild() As String
    Dim HyperLinkToParent As String
    Dim HyperLinkToChild As String
    Dim LinkPlaceParent As String
    Dim PageParent, NameIdParent As String
    Dim i As Integer
    
    'Set shpParent = ActivePage.Shapes("Sheet.6") 'для отладки
    
    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    LinkPlaceParent = PageParent + "/" + NameIdParent 'Для проверки ссылки в дочернем
    
        'Находим подключенный дочерний (через HyperLink)
        HyperLinkToChild = shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
        If HyperLinkToChild <> "" Then 'нашли адрес очищаемого
            
            'Находим дочерний шейп разбивая HyperLink на имя страницы и имя шейпа
            mstrAdrChild = Split(HyperLinkToChild, "/")
            On Error GoTo netu_dochernego 'вдруг его уже удалили и ссылку забыли почистить
            Set shpChild = ActiveDocument.Pages(mstrAdrChild(0)).Shapes(mstrAdrChild(1))
            'В контакте находим ссылку на катушку
            HyperLinkToParent = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
            
            'Проверяем что контакт привязан именно к нашей катушке
            If HyperLinkToParent = LinkPlaceParent Then
                'Чистим дочерний шейп
                shpChild.CellsU("Prop.Number").FormulaU = ""
                shpChild.CellsU("Prop.Name").FormulaU = """"""
                shpChild.CellsU("User.LocLink").FormulaU = """"""
                shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """"""
            End If
        End If
    
netu_dochernego:
'Почистили все дочерние. Удаляем родителя. - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete

    
End Sub

'Sub ClearReferenceWireLinkEvent(vsoShapeEvent As Visio.Shape)
''------------------------------------------------------------------------------------------------------------
'' Macros        : ClearReferenceWireLinkEvent - Чистит при копировании
'                'Чистим ссылки в при копировании разрыва провода.
'                'В EventDrop должна быть формула = CALLTHIS("ThisDocument.ClearReferenceWireLinkEvent")
'                'Этот макрос расположен в ThisDocument
''------------------------------------------------------------------------------------------------------------
'    Set vsoWindowEvent = ActiveWindow
'    Set vsoShapePaste = vsoShapeEvent
'    Click = False
'    ClearReferenceWireLink vsoShapePaste
'End Sub

Sub ClearReferenceWireLink(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearReferenceWireLink - Чистит при копировании
                'Чистим ссылки в при копировании разрыва провода.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReference.ClearReferenceWireLink", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим шейп
    vsoShape.CellsU("Prop.Number").FormulaU = ""
    vsoShape.CellsU("Prop.Name").FormulaU = """"""
    vsoShape.CellsU("User.LocLink").FormulaU = """"""
    vsoShape.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """"""

End Sub
