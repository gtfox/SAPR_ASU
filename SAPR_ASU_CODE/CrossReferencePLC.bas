'------------------------------------------------------------------------------------------------------------
' Module        : CrossReferencePLC - Перекрестные ссылки и связи PLC
' Author        : gtfox
' Date          : 2020.09.12
' Description   : Перекрестные ссылки и связи PLC и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit

'Активация формы создания связи PLC
Public Sub AddReferencePLCFrm(shpChild As Visio.Shape) 'Получили шейп с листа
    Load frmAddReferencePLC
    frmAddReferencePLC.run shpChild 'Передали его в форму
End Sub

Public Sub AddReferencePLC(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferencePLC - Создает связь между дочерним и родительским элементом

                'После выбора дочернего(PLCChild)/родительского(PLCParent) элемента заполняем небходимые поля для каждого из них
                'Имя(Sheet.4), Страница(Схема.3), Путь(Pages[Схема.3]!Sheet.4), Ссылка(HyperLink="Схема.3/Sheet.4"),Местоположение(/14.E7)
'------------------------------------------------------------------------------------------------------------
    Dim shpParentOld As Visio.Shape
    Dim shpChildOld As Visio.Shape
    Dim PageParent As String
    Dim NameIdParent As String
    Dim AdrParent As String
    Dim PageChild As String
    Dim NameIdChild As String
    Dim AdrChild As String
    Dim HyperLinkToChild As String
    Dim HyperLinkToParentOld As String
    Dim mstrAdrParentOld() As String
    Dim HyperLinkToChildOld As String
    Dim mstrAdrChildOld() As String
    Dim i As Integer

    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild
    HyperLinkToChild = PageChild + "/" + NameIdChild
    
    '---Отвязываемся от сущ PLCParent (чистим родителя от себя)
    
    'Это эквивалентно действиям перед удалением себя DeletePLCChild только в конце удаления непроисходит
    'Проверяем текущую привязку PLCChild к старому PLCParent и чистим ее в старом PLCParent
    DeletePLCChild shpChild 'чистим связи у начинки сидящей в PLCParent

    '---Привязываем PLCChild к новому PLCParent---
    
    'Пишем себя PLCChild в свободную строку Hyperlink в родителе PLCParent
    For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
        If shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0) Like "" Then 'нашли первую пустую строку в родительском
            'Заполняем родительский шейп
            shpParent.CellsU("Hyperlink." & i & ".SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
            shpParent.CellsU("Hyperlink." & i & ".ExtraInfo").FormulaU = AdrChild + "!User.Location"   'Pages[Схема.3]!Sheet.4!User.Location
            
            shpParent.CellsU("Hyperlink." & i & ".Description").FormulaU = """" & PageChild & ": """ & Chr(38) & "Hyperlink." & i & ".ExtraInfo"
            shpParent.CellsU("Hyperlink." & i & ".Invisible").FormulaU = "STRSAME(Hyperlink." & i & ".SubAddress,"""")"
            Exit For
        Else 'Нет свободной строки в родителе, создаем и прописываемся
            If i = shpParent.Section(visSectionHyperlink).Count Then
                'Добавляем новую строку, т.к. нет свободных существующих строк
                shpParent.AddRow visSectionHyperlink, visRowLast, visTagDefault
                shpParent.CellsSRC(visSectionHyperlink, visRowLast, visHLinkDescription).RowNameU = i + 1
                'Заполняем родительский шейп
                shpParent.CellsU("Hyperlink." & (i + 1) & ".SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
                shpParent.CellsU("Hyperlink." & (i + 1) & ".ExtraInfo").FormulaU = AdrChild + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location
                
                 shpParent.CellsU("Hyperlink." & (i + 1) & ".Description").FormulaU = """" & PageChild & ": """ & Chr(38) & "Hyperlink." & (i + 1) & ".ExtraInfo"
                shpParent.CellsU("Hyperlink." & (i + 1) & ".Invisible").FormulaU = "STRSAME(Hyperlink." & (i + 1) & ".SubAddress,"""")"
                
            End If
        End If
    Next

    'Заполняем дочерний шейп
    shpChild.CellsU("Hyperlink.PLC.SubAddress").FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"
    shpChild.CellsU("Hyperlink.PLC.ExtraInfo").FormulaU = AdrParent + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location
    
    shpChild.CellsU("Hyperlink.PLC.Description").FormulaU = """ПЛК  " & PageParent & ": """ & Chr(38) & "Hyperlink.PLC.ExtraInfo"
    shpChild.CellsU("Hyperlink.PLC.Invisible").FormulaU = "STRSAME(Hyperlink.PLC.SubAddress,"""")"
    
    shpChild.CellsU("User.NameParent").FormulaU = AdrParent + "!User.Name"  'Pages[Схема.3]!Sheet.4!User.Name
    shpChild.CellsU("User.LocationParent").FormulaU = AdrParent + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location

End Sub

Sub DeletePLCChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeletePLCChild - Чистит ссылки от себя перед удалением
                'Если PLCChild привязан, находим родителя (PLCParent), чистим его от удаляемого, и удаляем.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape
    Dim shpPLCModChild As Visio.Shape
    Dim mstrAdrParent() As String
    Dim HyperLinkToParent As String
    Dim HyperLinkToChild As String
    Dim PageChild As String
    Dim NameIdChild As String
    Dim AdrChild As String
    Dim i As Integer
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild
    HyperLinkToChild = PageChild + "/" + NameIdChild
    
    'Перебираем все PLCModChild внутри PLCChild и чистим ссылки во всех связанных PLCModParent
    For Each shpPLCModChild In shpChild.Shapes
        If shpPLCModChild.Name Like "PLCModChild*" Then
            DeletePLCModChild shpPLCModChild 'чистим ссылки в связанных PLCIOParent
        End If
    Next

    'Проверяем текущую привязку PLCChild к старому PLCParent и чистим ее в старом PLCParent
    Set shpParent = ShapeByHyperLink(shpChild.CellsU("Hyperlink.PLC.SubAddress").ResultStr(0))
    If Not shpParent Is Nothing Then
        
        'Перебираем ссылки на подключенные PLCChild в родительском шейпе
        For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
            If shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0) Like HyperLinkToChild Then 'нашли дочернего в родительском
                'Чистим родительский шейп
                shpParent.CellsU("Hyperlink." & i & ".SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
                shpParent.CellsU("Hyperlink." & i & ".ExtraInfo").FormulaForceU = ""
            End If
        Next
    End If
    
    ClearPLCChild shpChild 'чистим данные в себе перед удалением (т.к. этот макрос используется в перепривязке)т.к. все что выше чистило не наши кишки

End Sub

Sub DeletePLCParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeletePLCParent - Чистит ссылки от себя перед удалением
                'Смотрим ссылки в родительском PLCParent, идем по ним и чистим дочерние PLCChild, потом удаляем родителя.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------

    Dim shpChild As Visio.Shape
    Dim shpPLCModParent As Visio.Shape
    Dim mstrAdrChild() As String
    Dim HyperLinkToChild As String
    Dim HyperLinkToParent As String
    Dim LinkPlaceParent As String
    Dim PageParent As String
    Dim NameIdParent As String
    Dim i As Integer
    
    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    LinkPlaceParent = PageParent + "/" + NameIdParent 'Для проверки ссылки в дочернем
    
    'Перебираем все PLCModParent внутри PLCParent и чистим ссылки во всех связанных PLCModChild
    For Each shpPLCModParent In shpParent.Shapes
        If shpPLCModParent.Name Like "PLCModParent*" Then
            DeletePLCModParent shpPLCModParent 'чистим ссылки в связанных PLCIOParent
        End If
    Next

    'Перебираем ссылки на подключенные PLCChild в родительском шейпе
    For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
        Set shpChild = ShapeByHyperLink(shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0))
        If Not shpChild Is Nothing Then
        
            'В PLCChild находим ссылку на PLCParent
            HyperLinkToParent = shpChild.CellsU("Hyperlink.PLC.SubAddress").ResultStr(0)
            'Проверяем что контакт привязан именно к нашей катушке
            If HyperLinkToParent = LinkPlaceParent Then
                'Чистим дочерний шейп
                shpChild.CellsU("Hyperlink.PLC.SubAddress").FormulaForceU = """"""
                shpChild.CellsU("Hyperlink.PLC.ExtraInfo").FormulaForceU = ""
                shpChild.CellsU("User.NameParent").FormulaForceU = ""
                shpChild.CellsU("User.LocationParent").FormulaForceU = ""
            End If
        End If
    Next
    
    ClearPLCParent shpParent 'чистим данные в себе перед удалением (т.к. этот макрос используется в перепривязке)т.к. все что выше чистило не наши кишки

End Sub

Sub ClearPLCChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearPLCChild - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferencePLC.ClearPLCChild", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim shpPLCModChild As Visio.Shape
    
    'Перебираем все PLCModChild внутри PLCChild и чистим в них все ссылки
    For Each shpPLCModChild In shpChild.Shapes
        If shpPLCModChild.Name Like "PLCModChild*" Then
            ClearPLCModChild shpPLCModChild 'чистим ссылки
        End If
    Next
    
    'Чистим дочерний шейп
    shpChild.CellsU("Hyperlink.PLC.SubAddress").FormulaForceU = """"""
    shpChild.CellsU("Hyperlink.PLC.ExtraInfo").FormulaForceU = ""
    shpChild.CellsU("User.NameParent").FormulaForceU = ""
    shpChild.CellsU("User.LocationParent").FormulaForceU = ""

End Sub

Sub ClearPLCParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearPLCParent - Чистит родительский при копировании
                'Чистим ссылки в родительском при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'ClearPLCParent вызывается в ThisDocument.EventDropAutoNum
                'В EventMultiDrop должна быть формула = CALLTHIS("AutoNumber.AutoNum", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim shpPLCModParent As Visio.Shape
    
    'Перебираем все PLCModParent внутри PLCParent и чистим в них все ссылки
    For Each shpPLCModParent In shpParent.Shapes
        If shpPLCModParent.Name Like "PLCModParent*" Then
            ClearPLCModParent shpPLCModParent 'чистим ссылки
        End If
    Next
    
    'Перебираем ссылки на подключенные PLCChild в родительском шейпе
    For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
        If Not shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0) Like "" Then 'нашли дочернего в родительском
            'Чистим родительский шейп
            shpParent.CellsU("Hyperlink." & i & ".SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
            shpParent.CellsU("Hyperlink." & i & ".ExtraInfo").FormulaForceU = ""
        End If
    Next
End Sub