'------------------------------------------------------------------------------------------------------------
' Module        : CrossReferenceRelay - Перекрестные ссылки элементов схемы
' Author        : gtfox
' Date          : 2020.05.17
' Description   : Перекрестные ссылки элементов схемы и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit
Public colThumbToDelete As Collection

'Активация формы создания связи элементов схемы
Public Sub AddReferenceRelayFrm(shpChild As Visio.Shape) 'Получили шейп с листа
    Load frmAddReferenceRelay
    frmAddReferenceRelay.run shpChild 'Передали его в форму
End Sub

Sub AddReferenceRelay(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferenceRelay - Создает связь между дочерним и родительским элементом

                'После выбора дочернего(контакт)/родительского(катушка) элемента заполняем небходимые поля для каждого из них
                'Имя(Sheet.4), Страница(Схема.3), Путь(Pages[Схема.3]!Sheet.4), Ссылка(HyperLink="Схема.3/Sheet.4"),
                'Тип контакта(NO/NC), Местоположение(/14.E7), Номер контакта(KL1.3)
                'Ссылки на контакты в катушке формируются формулами в ShapeSheet
                'Нумерация контактов автоматическая, формулами в Scratch.B1-B4 катушки
                'Контактов у катушки может быть 4
'------------------------------------------------------------------------------------------------------------

    Dim shpParentOld As Visio.Shape
    Dim PageParent As String, NameIdParent As String, AdrParent As String, GUIDParent As String
    Dim PageChild  As String, NameIdChild As String, AdrChild As String, GUIDChild As String
    Dim i As Integer
    Dim HyperLinkToChild As String
    Dim HyperLinkToParentOld As String
    Dim mstrAdrParentOld() As String
    Dim Kontaktov As Integer 'число контактов в катушке/родительском элементе

    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
    GUIDParent = shpParent.UniqueID(visGetOrMakeGUID)
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild
    HyperLinkToChild = PageChild + "/" + NameIdChild
    GUIDChild = shpChild.UniqueID(visGetOrMakeGUID)
    
    Kontaktov = shpParent.Cells("Prop.Kontaktov").Result(0)

    'Проверяем текущую привязку контакта к старой катушке и чистим ее в старой катушке
    Set shpParentOld = ShapeByGUID(shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkExtraInfo).ResultStr(0))
    
    If Not shpParentOld Is Nothing Then
        'Ищем строку в Scratch катушки(родителя) с адресом удаляемого контакта (дочернего)
        For i = 1 To shpParentOld.Section(visSectionScratch).Count
            If shpParentOld.CellsU("Scratch.A" & i).ResultStr(0) = HyperLinkToChild And shpParentOld.CellsSRC(visSectionHyperlink, i - 1, visHLinkExtraInfo).ResultStr(0) = GUIDChild Then
                'Чистим родительский шейп
                shpParentOld.CellsU("Scratch.A" & i).FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
                shpParentOld.CellsU("Scratch.C" & i).FormulaForceU = ""
                shpParentOld.CellsU("Scratch.D" & i).FormulaForceU = ""
                shpParentOld.CellsSRC(visSectionHyperlink, i - 1, visHLinkExtraInfo).Formula = ""
                Exit For
            End If
        Next
    End If

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
            shpParent.CellsU("Scratch.D" & i).FormulaU = AdrChild + "!User.SAType"  'Pages[Схема.3]!Sheet.4!User.SAType
            shpParent.CellsSRC(visSectionHyperlink, i - 1, visHLinkExtraInfo).Formula = GUIDChild
            
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
            shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkExtraInfo).Formula = GUIDParent
            'Ссылка на описание из катушки
            If shpChild.Shapes("Desc").CellExists("Fields.Value", 0) Then
                shpChild.Shapes("Desc").Cells("Fields.Value").FormulaU = "SHAPETEXT(" + "Pages[" + PageParent + "]!" + shpParent.Shapes("Desc").NameID + "!TheText)"
            Else
                shpChild.Shapes("Desc").Characters.AddCustomFieldU "SHAPETEXT(" + "Pages[" + PageParent + "]!" + shpParent.Shapes("Desc").NameID + "!TheText)", visFmtNumGenNoUnits
            End If
            
            Exit For
        End If
        'Ограничение числа контактов катушки/родительского элемента
        If i = Kontaktov Then
            MsgBox "В элементе " & Kontaktov & " контакта, и все они задействованы" & vbCrLf, vbOKOnly + vbInformation, "САПР-АСУ: Нет свободных контактов"
            Exit For
        End If
    Next

End Sub

Sub DeleteRelayChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteRelayChild - Удаляет дочерний элемент
                'Если контакт привязан, находим родителя, чистим его от удаляемого, и удаляем.
                'Удаляем миниатюру катушки, если она была
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape
    Dim mstrAdrParent() As String
    Dim HyperLinkToParent As String
    Dim HyperLinkToChild As String
    Dim GUIDChild As String
    Dim PageChild, NameIdChild As String
    Dim i As Integer
    
    'Проверяем текущую привязку
    Set shpParent = ShapeByGUID(shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkExtraInfo).ResultStr(0))
    
    If Not shpParent Is Nothing Then
    
        PageChild = shpChild.ContainingPage.NameU
        NameIdChild = shpChild.NameID
        HyperLinkToChild = PageChild + "/" + NameIdChild
        GUIDChild = shpChild.UniqueID(visGetOrMakeGUID)
        
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
                shpParent.CellsSRC(visSectionHyperlink, i - 1, visHLinkExtraInfo).Formula = ""
                'Удаляем миниатюры у родителя, если они были
                ThumbDelete shpParent
                'Удаляем дочерний шейп
                'shpChild.Delete
                
                Exit For
            End If
        Next
    Else
        'Удаляем контакт не связанный с катушкой  - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete
        'shpChild.Delete

    End If
    
    'Удаляем миниатюры контактов, если они были
    ThumbDelete shpChild
    
    'Отключаем провода от элемента
    UnplugWire 1, shpChild
    
End Sub

Sub DeleteRelayParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteRelayParent - Удаляет родительский элемент
                'Смотрим ссылки в родительском, идем по ним и чистим дочерние, потом удаляем родителя.
                'Удаляем миниатюры контактов, если они были
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    'Dim shpParent As Visio.Shape
    Dim shpChild As Visio.Shape
    Dim mstrAdrChild() As String
    Dim HyperLinkToParent As String
    Dim HyperLinkToChild As String
    Dim LinkPlaceParent As String
    Dim PageParent As String
    Dim NameIdParent As String
    Dim GUIDParent As String
    Dim GUIDChild As String
    Dim i As Integer

    GUIDParent = shpParent.UniqueID(visGetOrMakeGUID)

    'Ищем строки в Scratch катушки(родителя) с адресами удаляемых контактов (дочерних)
    If shpParent.SectionExists(visSectionScratch, 0) Then
        For i = 1 To shpParent.Section(visSectionScratch).Count
            GUIDChild = shpParent.CellsSRC(visSectionHyperlink, i - 1, visHLinkExtraInfo).ResultStr(0)
            Set shpChild = ShapeByGUID(GUIDChild)
            'Проверяем что контакт привязан именно к нашей катушке
            If GUIDParent = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkExtraInfo).ResultStr(0) Then
                'Чистим дочерний шейп
                ClearRelayChild shpChild
            End If
        Next
    End If
    'Почистили все дочерние. Удаляем родителя. - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete
    'shpParent.Delete
    
    'Удаляем миниатюры контактов, если они были
    ThumbDelete shpParent
    
    'Отключаем провода от элемента
    UnplugWire 1, shpParent
    
End Sub


Sub ClearRelayChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearRelayChild - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferenceRelay.ClearRelayChild", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим дочерний шейп
    shpChild.CellsU("User.NameParent").FormulaForceU = ""
    shpChild.CellsU("User.Number").FormulaForceU = ""
    shpChild.CellsU("User.LocationParent").FormulaForceU = ""
    shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaForceU = """"""
    shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkExtraInfo).FormulaForceU = ""
    shpChild.Shapes("Desc").Cells("Fields.Value").FormulaU = ""
    'Удаляем миниатюры контактов, если они были
    ThumbDelete shpChild

End Sub

Sub ClearRelayParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearRelayParent - Чистит родительский при копировании
                'Чистим ссылки в родительском при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'ClearParentScratch вызывается в ThisDocument.EventDropAutoNum
                'В EventMultiDrop должна быть формула = CALLTHIS("AutoNumber.AutoNum", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    If shpParent.SectionExists(visSectionScratch, 0) Then
        For i = 1 To shpParent.Section(visSectionScratch).Count
            'Чистим шейп
            shpParent.CellsU("Scratch.A" & i).FormulaU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
            shpParent.CellsU("Scratch.C" & i).FormulaU = ""
            shpParent.CellsU("Scratch.D" & i).FormulaU = ""
            shpParent.CellsSRC(visSectionHyperlink, i - 1, visHLinkExtraInfo).FormulaU = ""
        Next
    End If
End Sub

Sub AddLocThumbAllInDoc()
'------------------------------------------------------------------------------------------------------------
' Macros        : AddLocThumbAllInDoc - Добавляет миниатюры контактов под реле во всём документе
                'Вставляет под катушку реле миниатюры всех активных контактов
                'Вставляет под контакт реле миниатюру катушки реле
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim vsoShapeOnPage As Visio.Shape
    Dim PageName As String
    PageName = cListNameCxema  'Имена листов
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If vsoPage.name Like PageName & "*" Then    'Берем те, что содержат "Схема" в имени
            For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
                Select Case ShapeSAType(vsoShapeOnPage)
                    Case typeNO, typeNC, typeCoil, typeParent
                        AddLocThumb vsoShapeOnPage
                End Select
            Next
        End If
    Next
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
    Set vsoMaster = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("Thumb")
    
    'Удаляем миниатюры контактов, если они были
    ThumbDelete vsoShape
    
    If vsoShape.CellExistsU("User.SAType", 0) Then
        'Выясняем кому надо вставить миниатюры
        Select Case ShapeSAType(vsoShape)
        
            Case typeNO, typeNC 'Контакты
            
                If vsoShape.Cells("Hyperlink.Coil.SubAddress").ResultStr(0) <> "" Then
                    'Вставляем миниатюру контакта Thumbnail
                    Set shpThumb = vsoPage.Drop(vsoMaster, vsoShape.Cells("PinX").Result(0), vsoShape.Cells("PinY").Result(0))
                    'Заполняем поля
                    shpThumb.Cells("User.LocType").FormulaU = typeCoil
                    shpThumb.Cells("User.Location").FormulaU = vsoShape.NameU & "!User.LocationParent"
                    shpThumb.Cells("User.AdrSource").FormulaU = Chr(34) & vsoShape.ContainingPageID & "/" & vsoShape.id & Chr(34)
                    shpThumb.Cells("User.DeltaX").FormulaU = Chr(34) & DeltaX & Chr(34) 'shpThumb.Cells("PinX").ResultStrU("in")
                    shpThumb.Cells("User.DeltaY").FormulaU = Chr(34) & DeltaY & Chr(34) 'shpThumb.Cells("PinY").ResultStrU("in")
                    shpThumb.Cells("PinX").FormulaU = "=SETATREF(User.DeltaX,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.id & "!PinX))+Sheet." & vsoShape.id & "!PinX"
                    shpThumb.Cells("PinY").FormulaU = "=SETATREF(User.DeltaY,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.id & "!PinY))+Sheet." & vsoShape.id & "!PinY"
                    shpThumb.Cells("User.AdrSource.Prompt").FormulaU = vsoShape.UniqueID(visGetOrMakeGUID)
                End If
                
            Case typeCoil, typeParent 'Катушка реле
            
                n = 0
                'Перебираем активные ссылки на контакты
                For i = 1 To vsoShape.Section(visSectionScratch).Count 'Ищем строку в Scratch
                    If vsoShape.CellsU("Scratch.A" & i).ResultStr(0) <> "" Then 'не пустая строка
                        'Вставляем миниатюру контакта Thumbnail
                        Set shpThumb = vsoPage.Drop(vsoMaster, vsoShape.Cells("PinX").Result(0), vsoShape.Cells("PinY").Result(0))
                        'Заполняем поля
                        shpThumb.Cells("User.LocType").FormulaU = vsoShape.NameU & "!Scratch.D" & i
                        shpThumb.Cells("User.Location").FormulaU = vsoShape.NameU & "!Scratch.C" & i
                        shpThumb.Cells("User.AdrSource").FormulaU = Chr(34) & vsoShape.ContainingPageID & "/" & vsoShape.id & Chr(34)
                        shpThumb.Cells("User.DeltaX").FormulaU = Chr(34) & DeltaX & Chr(34) 'shpThumb.Cells("PinX").ResultStrU("in")
                        shpThumb.Cells("User.DeltaY").FormulaU = Chr(34) & (DeltaY + n * dN) & Chr(34) 'shpThumb.Cells("PinY").ResultStrU("in")
                        shpThumb.Cells("PinX").FormulaU = "=SETATREF(User.DeltaX,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.id & "!PinX))+Sheet." & vsoShape.id & "!PinX"
                        shpThumb.Cells("PinY").FormulaU = "=SETATREF(User.DeltaY,SETATREFEVAL(SETATREFEXPR(0)-Sheet." & vsoShape.id & "!PinY))+Sheet." & vsoShape.id & "!PinY"
                        shpThumb.Cells("User.AdrSource.Prompt").FormulaU = vsoShape.UniqueID(visGetOrMakeGUID)
                        n = n + 1
                    End If
                Next
        End Select
    End If
    ActiveWindow.DeselectAll
End Sub

Sub ThumbDelete(shpDelete As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ThumbDelete - Удаляет миниатюры контактов
                'Собираем миниатюры контактов, если они были и ссылались на нас, и удаляем
'------------------------------------------------------------------------------------------------------------
    Dim vsoShape As Visio.Shape
    Dim colThumb As Collection
    
    Set colThumb = New Collection

    'Собираем миниатюры контактов, если они были, в коллекцию для удаления
    For Each vsoShape In ActivePage.Shapes
        If ShapeSATypeIs(vsoShape, typeThumb) Then
            If vsoShape.Cells("User.AdrSource.Prompt").ResultStr(0) = shpDelete.UniqueID(visGetGUID) Then
'            If vsoShape.Cells("User.AdrSource").ResultStr(0) = shpDelete.ContainingPage.id & "/" & shpDelete.id Then
                colThumb.Add vsoShape
            End If
        End If
    Next
    'Удаляем найденные контакты
    For Each vsoShape In colThumb
        Application.EventsEnabled = False
        vsoShape.Delete
        Application.EventsEnabled = True
    Next
    Set colThumb = Nothing
End Sub

Sub UnplugWire(CleareWire As Boolean, vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : UnplugWire - Отключает провода от элемента
                'Ищем провода подключенные к нам и отцепляем
'------------------------------------------------------------------------------------------------------------
    Dim DeletedConnect As Visio.connect
    Dim ConnectedShape As Visio.Shape
    Dim i As Integer, ii As Integer
    Dim ShapeType As Integer
    
    'Ищем провода подключенные к нам и отцепляем. Перебор FromConnects.
    For i = 1 To vsoShape.FromConnects.Count
        Set DeletedConnect = vsoShape.FromConnects(i)
        Set ConnectedShape = DeletedConnect.FromSheet
        
        ShapeType = ShapeSAType(ConnectedShape)
        
        If ShapeType = typeWire Then
            If CleareWire Then
                If Not (ConnectedShape.Cells("Prop.Number").FormulaU Like "*!*") Or (ConnectedShape.Cells("User.AdrSource.Prompt").ResultStr(0) = vsoShape.UniqueID(visGetGUID)) Then 'Не Дочерний? или дочерний, но ссылается на нас (другой провод или разрыв провода)
                    'Чистим Провод
                    ConnectedShape.Cells("Prop.Number").FormulaU = ""
                    ConnectedShape.Cells("Prop.SymName").FormulaU = ""
                    ConnectedShape.Cells("User.AdrSource").FormulaU = ""
                    ConnectedShape.Cells("Prop.AutoNum").FormulaU = True
                    ConnectedShape.Cells("Prop.HideNumber").FormulaU = False
                    ConnectedShape.Cells("Prop.HideName").FormulaU = True
                    ConnectedShape.Cells("User.AdrSource.Prompt").FormulaU = ""
                    'Присваиваем номер проводу
                    'AutoNum ConnectedShape
                End If
            End If
            'Ищем каким концом повод приклеен к нам
            For ii = 1 To ConnectedShape.Connects.Count
                If ConnectedShape.Connects(ii).ToSheet = vsoShape Then
                    SetArrow 254, ConnectedShape.Connects(ii) 'Возвращаем красную стрелку
                    'UnGlue ConnectedShape.Connects(ii) 'Отклеиваем
                    Exit For
                End If
            Next
        End If
    Next
End Sub


