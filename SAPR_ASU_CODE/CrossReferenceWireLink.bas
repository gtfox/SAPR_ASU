'------------------------------------------------------------------------------------------------------------
' Module        : CrossReferenceWireLink - Перекрестные ссылки разрывов проводов
' Author        : gtfox
' Date          : 2020.06.02
' Description   : Перекрестные ссылки разрывов проводов и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit

'Активация формы создания связи разрывов проводов
Public Sub AddReferenceWireLinkFrm(shpChild As Visio.Shape) 'Получили шейп с листа
    Load frmAddReferenceWireLink
    frmAddReferenceWireLink.run shpChild 'Передали его в форму
End Sub

Sub AddReferenceWireLink(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferenceWireLink - Создает связь между шейпами разрывов проводов

                'После выбора дочернего/родительского элемента заполняем небходимые поля для каждого из них
                'Номер провода Prop.Number(5), Название провода Prop.Name("24V"),Местоположение User.LocLink (/14.E7), Ссылка(HyperLink="Схема.3/Sheet.4"),
                'У одного родителя может быть и дочерний (связь 1:1)
'------------------------------------------------------------------------------------------------------------

    Dim shpParentOld As Visio.Shape
    Dim shpChildOld As Visio.Shape
    Dim PageParent, NameIdParent, AdrParent As String
    Dim PageChild, NameIdChild, AdrChild As String
    Dim i As Integer
    Dim HyperLinkToChild As String
    Dim HyperLinkToParentOld As String
    Dim mstrAdrParentOld() As String
    Dim HyperLinkToChildOld As String
    Dim mstrAdrChildOld() As String

    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild
    HyperLinkToChild = PageChild + "/" + NameIdChild

    'Проверяем текущую привязку разрыва провода(дочернего) к старому разрыву(родильскому) и чистим его в старом разрыве.
    'А еще в старом разрыве была вторая половинка - старый дочерний. Его тоже чистим.
    Set shpParentOld = ShapeByHyperLink(shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0))
    If Not shpParentOld Is Nothing Then
        'Чистим родительский шейп
        shpParentOld.CellsU("User.LocLink").FormulaU = """"""
        shpParentOld.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
    End If
    
    'Находим подключенный к новому родителю дочерний шейп (если он есть)
    Set shpChildOld = ShapeByHyperLink(shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0))
    If Not shpChildOld Is Nothing Then
        ClearReferenceWireLink shpChildOld
    End If

    'Привязываемся к новому разрыву провода
    
    'Заполняем родительский шейп
    shpParent.CellsU("User.LocLink").FormulaU = AdrChild + "!User.Location"  'Pages[Схема.3]!Sheet.4!User.Location
    shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
    
    'Заполняем дочерний шейп
    shpChild.CellsU("Prop.Number").FormulaU = AdrParent + "!Prop.Number"
    shpChild.CellsU("Prop.SymName").FormulaU = AdrParent + "!Prop.SymName"
    shpChild.CellsU("User.LocLink").FormulaU = AdrParent + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location
    shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"

End Sub

Sub DeleteWireLinkChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteWireLinkChild - Удаляет дочерний элемент
                'Если разрыв провода привязан, находим родителя, чистим его от удаляемого, и удаляем.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim mstrAdrParent() As String
    Dim HyperLinkToParent As String
    Dim i As Integer
    
    'Проверяем что разрыв провода привязан родителю
    Set shpParent = ShapeByHyperLink(shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0))
    If Not shpParent Is Nothing Then
            
        'Чистим родительский шейп
        shpParent.CellsU("User.LocLink").FormulaU = """"""
        shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
    
    End If

    'Отключаем провод, чистим в нем ссылки, автонумерация, стрелка
    UnplugWire 1, shpChild

    'Удаляем дочерний шейп - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete
End Sub

Sub DeleteWireLinkParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteWireLinkParent - Удаляет родительский элемент
                'Смотрим ссылки в родительском, идем по ним и чистим дочерние, потом удаляем родителя.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpChild As Visio.Shape
    Dim vsoShape As Visio.Shape
    Dim mstrAdrChild() As String
    Dim HyperLinkToParent As String
    Dim HyperLinkToChild As String
    Dim LinkPlaceParent As String
    Dim PageParent, NameIdParent As String
    Dim i As Integer
    
    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    LinkPlaceParent = PageParent + "/" + NameIdParent 'Для проверки ссылки в дочернем
    
    'Находим подключенный дочерний (через HyperLink)
    Set shpChild = ShapeByHyperLink(shpParent.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0))
    If Not shpChild Is Nothing Then
    
        'В контакте находим ссылку на катушку
        HyperLinkToParent = shpChild.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0)
        
        'Проверяем что контакт привязан именно к нашей катушке
        If HyperLinkToParent = LinkPlaceParent Then
            ClearReferenceWireLink shpChild
        End If
    End If

    'Проверяем подключенный провод, стрелка
    UnplugWire 0, shpParent

    'Почистили все дочерние. Удаляем родителя. - автоматически т.к. макрос вызывается в событии vsoPagesEvent_BeforeShapeDelete
End Sub


Sub ClearReferenceWireLink(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearReferenceWireLink - Чистит при копировании
                'Чистим ссылки при копировании разрыва провода.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReference.ClearReferenceWireLink", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим шейп
    vsoShape.CellsU("Prop.Number").FormulaU = ""
    vsoShape.CellsU("Prop.SymName").FormulaU = """"""
    vsoShape.CellsU("User.LocLink").FormulaU = """"""
    vsoShape.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).FormulaU = """"""

End Sub


Sub GoHyperLink(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : GoHyperLink - Переходит по ссылке в разрыве провода
                'Переходит по ссылке в разрыве провода по двойному клику
                'Вызов макроса в EventDblClick  =CALLTHIS("CrossReference.GoHyperLink","SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim shpTarget As Visio.Shape
    Dim shpParent As Visio.Shape
    Dim HyperLinkToTarget As String
    Dim mstrAdrTarget() As String
    
'    Dim pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна
'    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна

    'Находим шейп-цель для последующего выделения
    Set shpTarget = ShapeByHyperLink(vsoShape.CellsSRC(visSectionHyperlink, 0, visHLinkSubAddress).ResultStr(0))
    If Not shpParent Is Nothing Then
        'Переходим по ссылке
        vsoShape.Hyperlinks("1").Follow
        ActiveWindow.DeselectAll
'        ActiveWindow.SetViewRect shpTarget.Cells("PinX") - pinWidth / 2, shpTarget.Cells("PinY") + pinHeight / 2, pinWidth, pinHeight
        ActiveWindow.Select shpTarget, visSelect
    End If

End Sub