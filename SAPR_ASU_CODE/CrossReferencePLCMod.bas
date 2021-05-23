'------------------------------------------------------------------------------------------------------------
' Module        : CrossReferencePLCMod - Перекрестные ссылки и связи модулей внутри PLC
' Author        : gtfox
' Date          : 2020.09.14
' Description   : Перекрестные ссылки и связи модулей внутри PLC и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit

'Активация формы создания связи модулей внутри PLC
Public Sub AddReferencePLCModFrm(shpChild As Visio.Shape) 'Получили шейп с листа
    Load frmAddReferencePLCMod
    frmAddReferencePLCMod.run shpChild 'Передали его в форму
End Sub


'------------------------------------------------------------------------------------------------------------
'----------------------------------------------PLCMod--------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'---------AddReferencePLCMod
'---------DeletePLCModChild
'---------DeletePLCModParent
'---------ClearPLCModChild
'---------ClearPLCModParent

Public Sub AddReferencePLCMod(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferencePLCMod - Создает связь между дочерним и родительским элементом

                'После выбора дочернего(PLCModChild)/родительского(PLCModParent) элемента заполняем небходимые поля для каждого из них
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

    '---Отвязываемся от сущ PLCModParent (чистим родителя от себя)
    
    'Это эквивалентно действиям перед удалением себя DeletePLCModChild только в конце удаления непроисходит
    'Проверяем текущую привязку PLCModChild к старому PLCModParent и чистим ее в старом PLCModParent
    DeletePLCModChild shpChild 'чистим связи у начинки сидящей в shpChild (PLCModChild)

    '---Привязываем PLCModChild к новому PLCModParent---
    
    'Перебираем ссылки на подключенные PLCModChild в родительском шейпе
    For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
        If shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0) Like "" Then 'нашли первую пустую строку в родительском
        
            'Заполняем родительский шейп
            shpParent.CellsU("Hyperlink." & i & ".SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
            shpParent.CellsU("Hyperlink." & i & ".ExtraInfo").FormulaU = "Pages[" + shpChild.Parent.ContainingPage.NameU + "]!" + shpChild.Parent.NameID + "!User.Location"
            
            shpParent.CellsU("Hyperlink." & i & ".Description").FormulaU = """Модуль ПЛК  " & PageChild & ": """ & Chr(38) & "Hyperlink." & i & ".ExtraInfo"
            shpParent.CellsU("Hyperlink." & i & ".Invisible").FormulaU = "STRSAME(Hyperlink." & i & ".SubAddress,"""")"
            Exit For
        Else
            If i = shpParent.Section(visSectionHyperlink).Count Then 'все строки заняты
                'Добавляем новую строку, т.к. нет свободных существующих строк
                shpParent.AddRow visSectionHyperlink, visRowLast, visTagDefault
                shpParent.CellsSRC(visSectionHyperlink, visRowLast, visHLinkDescription).RowNameU = i + 1
                
                'Заполняем родительский шейп
                shpParent.CellsU("Hyperlink." & (i + 1) & ".SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
                shpParent.CellsU("Hyperlink." & (i + 1) & ".ExtraInfo").FormulaU = "Pages[" + shpChild.Parent.ContainingPage.NameU + "]!" + shpChild.Parent.NameID + "!User.Location"
                
                shpParent.CellsU("Hyperlink." & (i + 1) & ".Description").FormulaU = """Модуль ПЛК  " & PageChild & ": """ & Chr(38) & "Hyperlink." & (i + 1) & ".ExtraInfo"
                shpParent.CellsU("Hyperlink." & (i + 1) & ".Invisible").FormulaU = "STRSAME(Hyperlink." & (i + 1) & ".SubAddress,"""")"
                
            End If
        End If
    Next

    'Заполняем дочерний шейп
    shpChild.CellsU("Hyperlink.PLCMod.SubAddress").FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"
    shpChild.CellsU("Hyperlink.PLCMod.ExtraInfo").FormulaU = "Pages[" + shpParent.Parent.ContainingPage.NameU + "]!" + shpParent.Parent.NameID + "!User.Location" 'Ссылка на ПЛК а не на модуль
    
    shpChild.CellsU("Hyperlink.PLCMod.Description").FormulaU = """Модуль ПЛК  " & PageParent & ": """ & Chr(38) & "Hyperlink.PLCMod.ExtraInfo"
    shpChild.CellsU("Hyperlink.PLCMod.Invisible").FormulaU = "STRSAME(Hyperlink.PLCMod.SubAddress,"""")"
    
    shpChild.CellsU("User.NameParent").FormulaU = AdrParent + "!User.Name"  'Pages[Схема.3]!Sheet.4!User.Name
    shpChild.CellsU("User.LocationParent").FormulaU = "Pages[" + shpParent.Parent.ContainingPage.NameU + "]!" + shpParent.Parent.NameID + "!User.Location" 'Ссылка на ПЛК а не на модуль
    
    shpChild.CellsU("Prop.Model").FormulaU = AdrParent + "!Prop.Model"
End Sub

Sub DeletePLCModChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeletePLCModChild - Удаляет дочерний элемент
                'Если PLCModChild привязан, находим родителя (PLCModParent), чистим его от удаляемого, и удаляем.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape
    Dim shpPLCIOChild As Visio.Shape
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
    
    'Перебираем все PLCIOChild внутри PLCModChild и чистим ссылки во всех связанных PLCIOParent
    For Each shpPLCIOChild In shpChild.Shapes
        If shpPLCIOChild.Name Like "PLCIO*" Then
            DeletePLCIOChild shpPLCIOChild 'чистим ссылки в связанных PLCIOParent
        End If
    Next

    'Проверяем текущую привязку PLCModChild к старому PLCModParent и чистим ее в старом PLCModParent
    Set shpParent = ShapeByHyperLink(shpChild.CellsU("Hyperlink.PLCMod.SubAddress").ResultStr(0))
    If Not shpParent Is Nothing Then
        
        'Перебираем ссылки на подключенные PLCModChild в родительском шейпе
        For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
            If shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0) Like HyperLinkToChild Then 'нашли дочернего в родительском
                'Чистим родительский шейп
                shpParent.CellsU("Hyperlink." & i & ".SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
                shpParent.CellsU("Hyperlink." & i & ".ExtraInfo").FormulaForceU = ""
            End If
        Next
    End If
    
    ClearPLCModChild shpChild 'чистим данные в себе перед удалением (т.к. этот макрос используется в перепривязке)т.к. все что выше чистило не наши кишки
    
End Sub

Sub DeletePLCModParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeletePLCModParent - Удаляет родительский элемент
                'Смотрим ссылки в родительском PLCModParent, идем по ним и чистим дочерние PLCModChild, потом удаляем родителя.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpChild As Visio.Shape
    Dim shpPLCIOParent As Visio.Shape
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

    'Перебираем все PLCIOParent внутри PLCModParent и чистим ссылки во всех связанных PLCIOChild
    For Each shpPLCIOParent In shpParent.Shapes
        If (shpPLCIOParent.Name Like "PLCIOL*") Or (shpPLCIOParent.Name Like "PLCIOR*") Then
            DeletePLCIOParent shpPLCIOParent 'чистим ссылки в связанных PLCIOChild
        End If
    Next

    'Перебираем ссылки на подключенные PLCModChild в родительском шейпе
    For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
        Set shpChild = ShapeByHyperLink(shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0))
        If Not shpChild Is Nothing Then
        
            'В PLCModChild находим ссылку на PLCModParent
            HyperLinkToParent = shpChild.CellsU("Hyperlink.PLCMod.SubAddress").ResultStr(0)
            'Проверяем что он привязан именно к нашей нам
            If HyperLinkToParent = LinkPlaceParent Then
                'Чистим дочерний шейп
                shpChild.CellsU("Hyperlink.PLCMod.SubAddress").FormulaForceU = """"""
                shpChild.CellsU("Hyperlink.PLCMod.ExtraInfo").FormulaForceU = ""
                shpChild.CellsU("User.NameParent").FormulaU = ""
                shpChild.CellsU("User.LocationParent").FormulaU = ""
                shpChild.CellsU("Prop.Model").FormulaU = ""
            End If
        End If
    Next
    
    ClearPLCModParent shpParent 'чистим данные в себе перед удалением (т.к. этот макрос используется в перепривязке)т.к. все что выше чистило не наши кишки

End Sub

Sub ClearPLCModChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearPLCModChild - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferencePLCMod.ClearPLCModChild", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim shpPLCIOChild As Visio.Shape
    
    'Перебираем все PLCIOChild внутри PLCModChild и чистим в них ссылки
    For Each shpPLCIOChild In shpChild.Shapes
        If shpPLCIOChild.Name Like "PLCIO*" Then
            ClearPLCIOChild shpPLCIOChild 'чистим ссылки
        End If
    Next

    'Чистим дочерний шейп
    shpChild.CellsU("Hyperlink.PLCMod.SubAddress").FormulaForceU = """"""
    shpChild.CellsU("Hyperlink.PLCMod.ExtraInfo").FormulaForceU = ""
    shpChild.CellsU("User.NameParent").FormulaU = ""
    shpChild.CellsU("User.LocationParent").FormulaU = ""
    shpChild.CellsU("Prop.Model").FormulaU = ""
End Sub

Sub ClearPLCModParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearPLCModParent - Чистит родительский при копировании
                'Чистим ссылки в родительском при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferencePLCMod.ClearPLCModParent", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim shpPLCIOParent As Visio.Shape
    
    'Перебираем все PLCIOParent внутри PLCModParent и чистим в них ссылки
    For Each shpPLCIOParent In shpParent.Shapes
        If (shpPLCIOParent.Name Like "PLCIOL*") Or (shpPLCIOParent.Name Like "PLCIOR*") Then
            ClearPLCIOParent shpPLCIOParent 'чистим ссылки
        End If
    Next
    
    'Перебираем ссылки на подключенные PLCModChild в родительском шейпе
    For i = 1 To shpParent.Section(visSectionHyperlink).Count 'Ищем строку в Hyperlink
        If Not shpParent.CellsU("Hyperlink." & i & ".SubAddress").ResultStr(0) Like "" Then 'нашли дочернего в родительском
            'Чистим родительский шейп
            shpParent.CellsU("Hyperlink." & i & ".SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
            shpParent.CellsU("Hyperlink." & i & ".ExtraInfo").FormulaForceU = ""
        End If
    Next
End Sub

'------------------------------------------------------------------------------------------------------------
'----------------------------------------------PLCIO---------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'---------AddReferencePLCIO
'---------DeletePLCIOChild
'---------DeletePLCIOParent
'---------ClearPLCIOChild
'---------ClearPLCIOParent

Sub AddReferencePLCIO(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferencePLCIO - Создает связь между дочерним и родительским элементом

                'После выбора PLCIOChild в дочернем модуле и typePLCIOParent в родительском модуле заполняем небходимые поля для каждого из них
'------------------------------------------------------------------------------------------------------------
    Dim shpParentOld As Visio.Shape
    Dim shpChildOld As Visio.Shape
    Dim PageParent As String, NameIdParent As String, AdrParent As String
    Dim PageChild  As String, NameIdChild As String, AdrChild As String
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
    
    '

    '---Отвязываемся от сущ PLCIOParent (чистим родителя от себя)---
    'Это эквивалентно действиям перед удалением себя DeletePLCIOChild только в конце удаления непроисходит
    'Проверяем привязку текущего typePLCIOChild к старому typePLCIOParent (в родительском модуле) и чистим ее там
    DeletePLCIOChild shpChild 'чистим связи у начинки сидящей в shpChild (PLCIOChild)
    
    '---Отвязываем новый PLCIOParent от старых связей---
    'Если новый typePLCIOParent связан с другим typePLCIOChild, то сначала чистим другой typePLCIOChild, а потом привязываемся
    DeletePLCIOParent shpParent 'чистим связи у начинки сидящей в shpParent (PLCIOParent)

    '---Привязываем typePLCIOChild к новому typePLCIOParent---

    'Заполняем родительский шейп
    shpParent.CellsU("Hyperlink.IO.SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
    shpParent.CellsU("Hyperlink.IO.ExtraInfo").FormulaU = "Pages[" + shpChild.Parent.Parent.ContainingPage.NameU + "]!" + shpChild.Parent.Parent.NameID + "!User.Location" 'Ссылка на ПЛК а не на модуль
    shpParent.CellsU("Hyperlink.IO.Description").FormulaU = """Вх./Вых. ПЛК  " & PageChild & ": """ & Chr(38) & "Hyperlink.IO.ExtraInfo"
    shpParent.CellsU("Hyperlink.IO.Invisible").FormulaU = "STRSAME(Hyperlink.IO.SubAddress,"""")"

    'Заполняем дочерний шейп
    shpChild.CellsU("Hyperlink.IO.SubAddress").FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"
    shpChild.CellsU("Hyperlink.IO.ExtraInfo").FormulaU = "Pages[" + shpParent.Parent.Parent.ContainingPage.NameU + "]!" + shpParent.Parent.Parent.NameID + "!User.Location" 'Ссылка на ПЛК а не на модуль
    
    shpChild.CellsU("Hyperlink.IO.Description").FormulaU = """Вх./Вых. ПЛК  " & PageParent & ": """ & Chr(38) & "Hyperlink.IO.ExtraInfo"
    shpChild.CellsU("Hyperlink.IO.Invisible").FormulaU = "STRSAME(Hyperlink.IO.SubAddress,"""")"
    
    shpChild.CellsU("User.AdrParent").FormulaU = """" + AdrParent + """"
    shpChild.CellsU("Prop.AutoNum").FormulaU = True

End Sub

Sub DeletePLCIOChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeletePLCIOChild - Удаляет дочерний элемент
                'Если typePLCIOChild привязан, находим родителя typePLCIOParent, чистим его от удаляемого, и удаляем.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape
    Dim mstrAdrParent() As String
    Dim HyperLinkToParent As String

    'Проверяем текущую привязку typePLCIOChild к typePLCIOParent и чистим ее в typePLCIOParent
    Set shpParent = ShapeByHyperLink(shpChild.CellsU("Hyperlink.IO.SubAddress").ResultStr(0))
    If Not shpParent Is Nothing Then
        ClearPLCIOParent shpParent
    End If
    
    ClearPLCIOChild shpChild 'чистим данные в себе перед удалением (т.к. этот макрос используется в перепривязке)т.к. все что выше чистило не наши кишки
    
End Sub

Sub DeletePLCIOParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeletePLCIOParent - Удаляет родительский элемент
                'Смотрим ссылки в родительском typePLCIOParent, идем по ним и чистим дочерние typePLCIOChild, потом удаляем родителя.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------

    Dim shpChild As Visio.Shape
    Dim mstrAdrChild() As String
    Dim HyperLinkToChild As String
    Dim LinkPlaceParent As String
    
    'Если typePLCIOParent связан с typePLCIOChild, то сначала чистим typePLCIOChild, а потом удаляем typePLCIOParent
    Set shpChild = ShapeByHyperLink(shpParent.CellsU("Hyperlink.IO.SubAddress").ResultStr(0))
    If Not shpChild Is Nothing Then
        ClearPLCIOChild shpChild
    End If
    
    ClearPLCIOParent shpParent 'чистим данные в себе перед удалением (т.к. этот макрос используется в перепривязке)т.к. все что выше чистило не наши кишки

End Sub

Sub ClearPLCIOChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearPLCIOChild - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferencePLCMod.ClearPLCIOChild", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим дочерний шейп
    shpChild.CellsU("Hyperlink.IO.SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
    shpChild.CellsU("Hyperlink.IO.ExtraInfo").FormulaForceU = ""
    shpChild.CellsU("User.AdrParent").FormulaForceU = """"""
    shpChild.CellsU("Prop.AutoNum").FormulaU = False

End Sub

Sub ClearPLCIOParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearPLCIOParent - Чистит родительский при копировании
                'Чистим ссылки в родительском при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferencePLCMod.ClearPLCIOParent", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим родительский шейп
    shpParent.CellsU("Hyperlink.IO.SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
    shpParent.CellsU("Hyperlink.IO.ExtraInfo").FormulaForceU = ""
End Sub


