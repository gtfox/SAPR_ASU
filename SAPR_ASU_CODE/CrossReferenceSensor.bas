'------------------------------------------------------------------------------------------------------------
' Module        : CrossReferenceSensor - Перекрестные ссылки элементов ВНЕ ШКАФА
' Author        : gtfox
' Date          : 2020.09.09
' Description   : Перекрестные ссылки элементов ВНЕ ШКАФА и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit

'Активация формы создания связи элементов ВНЕ ШКАФА
Public Sub AddReferenceSensorFrm(shpChild As Visio.Shape) 'Получили шейп с листа
    Load frmAddReferenceSensor
    frmAddReferenceSensor.run shpChild 'Передали его в форму
End Sub

'------------------------------------------------------------------------------------------------------------
'----------------------------------------------Sensor---------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'---------AddReferenceSensor
'---------DeleteSensorChild
'---------DeleteSensorParent
'---------ClearSensorChild
'---------ClearSensorParent

Sub AddReferenceSensor(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferenceSensor - Создает связь между датчиком на ФСА и датчиком наэл. схеме

                'После выбора дочернего(Датчик на ФСА = ДФ)/родительского(Датчик на эл.схеме = ДЭ) элемента заполняем небходимые поля для каждого из них
                'Имя(Sheet.4), Страница(Схема.3), Путь(Pages[Схема.3]!Sheet.4), Ссылка(HyperLink="Схема.3/Sheet.4"),Местоположение(/14.E7)

'------------------------------------------------------------------------------------------------------------
    Dim shpParentOld As Visio.Shape
    Dim shpChildOld As Visio.Shape
    Dim PageParent As String, NameIdParent As String, AdrParent As String
    Dim PageChild  As String, NameIdChild As String, AdrChild As String

    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild


    '---Отвязываем сущ ДЭ---
    'Проверяем текущую привязку ДФ к старому ДЭ и чистим ее в старом ДЭ
    'DeleteSensorChild нельзя использовать т.к. там чистится подвал, а тут этого не надо
    Set shpParentOld = ShapeByHyperLink(shpChild.CellsU("Hyperlink.Shema.SubAddress").ResultStr(0))
    If Not shpParentOld Is Nothing Then
        ClearSensorParent shpParentOld
    End If

    '---Отвязываем сущ ДФ---
    'Если новый ДЭ связан с другим ДФ, то сначала чистим другой ДФ, а потом привязываемся
    DeleteSensorParent shpParent

    '---Привязываем ДФ к новому ДЭ---

    'Заполняем родительский шейп
    shpParent.CellsU("Hyperlink.FSA.SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
    shpParent.CellsU("Hyperlink.FSA.ExtraInfo").FormulaU = AdrChild + "!User.Location"   'Pages[Схема.3]!Sheet.4!User.Location
    shpParent.CellsU("User.NameChild").FormulaU = AdrChild + "!User.Name"  'Pages[Схема.3]!Sheet.4!User.Name
    
    'Заполняем дочерний шейп
    shpChild.CellsU("Hyperlink.Shema.SubAddress").FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"
    shpChild.CellsU("Hyperlink.Shema.ExtraInfo").FormulaU = AdrParent + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location
    shpChild.CellsU("User.NameParent").FormulaU = AdrParent + "!User.Name"  'Pages[Схема.3]!Sheet.4!User.Name

End Sub

Sub DeleteSensorChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteSensorChild - Удаляет дочерний элемент
                'Если ДФ привязан, находим родителя (ДЭ), чистим его от удаляемого, и удаляем.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape

    'Проверяем текущую привязку ДФ к ДЭ и чистим ее в ДЭ
    Set shpParent = ShapeByHyperLink(shpChild.CellsU("Hyperlink.Shema.SubAddress").ResultStr(0))
    If Not shpParent Is Nothing Then
        ClearSensorParent shpParent
    End If
    
    'Чистим ссылки относящиеся к подвалу ФСА
    DeleteFSAPodvalParent shpChild
    
End Sub

Sub DeleteSensorParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteSensorParent - Удаляет родительский элемент
                'Смотрим ссылки в родительском ДЭ, идем по ним и чистим дочерние ДФ, потом удаляем родителя.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpChild As Visio.Shape
    
    'Если ДЭ связан с ДФ, то сначала чистим ДФ, а потом удаляем ДЭ
    Set shpChild = ShapeByHyperLink(shpParent.CellsU("Hyperlink.FSA.SubAddress").ResultStr(0))
    If Not shpChild Is Nothing Then
        ClearSensorChild shpChild
    End If

End Sub

Sub ClearSensorChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearSensorChild - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferenceSensor.ClearSensorChild", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим дочерний шейп
    shpChild.CellsU("Hyperlink.Shema.SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
    shpChild.CellsU("Hyperlink.Shema.ExtraInfo").FormulaForceU = ""
    shpChild.CellsU("User.NameParent").FormulaForceU = ""
    
    ClearFSAPodvalParent shpChild 'чистим подвальные ссылки в датчике ФСА
    
End Sub

Sub ClearSensorParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearSensorParent - Чистит родительский при копировании
                'Чистим ссылки в родительском при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'ClearSensorParent вызывается в ThisDocument.EventDropAutoNum
                'В EventMultiDrop должна быть формула = CALLTHIS("AutoNumber.AutoNum", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'Чистим родительский шейп
    shpParent.CellsU("Hyperlink.FSA.SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
    shpParent.CellsU("Hyperlink.FSA.ExtraInfo").FormulaForceU = ""
    shpParent.CellsU("User.NameChild").FormulaForceU = ""
End Sub

'------------------------------------------------------------------------------------------------------------
'----------------------------------------------FSAPodval-----------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'---------AddReferenceFSAPodval
'---------DeleteFSAPodvalChild
'---------DeleteFSAPodvalParent
'---------ClearFSAPodvalChild
'---------ClearFSAPodvalParent

Sub AddReferenceFSAPodval(shpChild As Visio.Shape, shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddReferenceFSAPodval - Создает связь между датчиком на ФСА (FSASensor) и каналом в подвале ФСА (FSAPodval)

                'После выбора дочернего канала в подвале ФСА/родительского(Датчик на ФСА) элемента заполняем небходимые поля для каждого из них
'------------------------------------------------------------------------------------------------------------
    Dim shpParentOld As Visio.Shape
    Dim shpChildOld As Visio.Shape
    Dim PageParent As String, NameIdParent As String, AdrParent As String
    Dim PageChild  As String, NameIdChild As String, AdrChild As String

    PageParent = shpParent.ContainingPage.NameU
    NameIdParent = shpParent.NameID
    AdrParent = "Pages[" + PageParent + "]!" + NameIdParent
    
    PageChild = shpChild.ContainingPage.NameU
    NameIdChild = shpChild.NameID
    AdrChild = "Pages[" + PageChild + "]!" + NameIdChild

    '---Отвязываем сущ FSASensor---
    'Проверяем текущую привязку FSAPodval к старому FSASensor и чистим ее в старом FSASensor
    DeleteFSAPodvalChild shpChild

    '---Отвязываем сущ FSAPodval---
    'Если новый FSASensor связан с другим FSAPodval, то сначала чистим другой FSAPodval, а потом привязываемся
    DeleteFSAPodvalParent shpParent

    '---Привязываем FSAPodval к новому FSASensor---

    'Заполняем родительский шейп FSASensor
    shpParent.CellsU("Hyperlink.FSA.SubAddress").FormulaU = """" + PageChild + "/" + NameIdChild + """" ' "Схема.3/Sheet.4"
    shpParent.CellsU("Hyperlink.FSA.ExtraInfo").FormulaU = AdrChild + "!User.Location"   'Pages[Схема.3]!Sheet.4!User.Location
    shpParent.CellsU("Prop.KanalNumber").FormulaU = AdrChild + "!Prop.Number"  'Pages[Схема.3]!Sheet.4!User.Name
    
    'Заполняем дочерний шейп FSAPodval
    shpChild.CellsU("Hyperlink.FSA.SubAddress").FormulaU = """" + PageParent + "/" + NameIdParent + """" ' "Схема.3/Sheet.4"
    shpChild.CellsU("Hyperlink.FSA.ExtraInfo").FormulaU = AdrParent + "!User.Location" 'Pages[Схема.3]!Sheet.4!User.Location
    shpChild.CellsU("User.NameParent").FormulaU = AdrParent + "!User.Name"  'Pages[Схема.3]!Sheet.4!User.Name
    shpChild.Shapes("Pomestu").CellsU("Prop.Place").FormulaU = AdrParent + "!Prop.Place"
    shpChild.Shapes("Pomestu").CellsU("Prop.Forma").FormulaU = AdrParent + "!Prop.Forma"
    shpChild.Shapes("Pomestu").CellsU("Prop.SymName").FormulaU = AdrParent + "!Prop.SymName"
    shpChild.Shapes("Pomestu").CellsU("Prop.Number").FormulaU = AdrParent + "!Prop.Number"
    shpChild.Shapes("Pomestu").CellsU("Prop.NameKontur").FormulaU = AdrParent + "!Prop.NameKontur"

End Sub

Sub DeleteFSAPodvalChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteFSAPodvalChild - Удаляет дочерний элемент
                'Если FSAPodval привязан, находим родителя FSASensor, чистим его от удаляемого, и удаляем.
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim shpParent As Visio.Shape

    'Проверяем текущую привязку FSAPodval к FSASensor и чистим ее в FSASensor
    Set shpParent = ShapeByHyperLink(shpChild.CellsU("Hyperlink.FSA.SubAddress").ResultStr(0))
    If Not shpParent Is Nothing Then
        ClearFSAPodvalParent shpParent
    End If
End Sub

Sub DeleteFSAPodvalParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteFSAPodvalParent - Удаляет родительский элемент
                'Смотрим ссылки в родительском FSASensor, идем по ним и чистим дочерние FSAPodval, потом удаляем родителя.
                'Этот макрос вызывается из DeleteSensorChild
'------------------------------------------------------------------------------------------------------------
    Dim shpChild As Visio.Shape

    'Если FSASensor связан с FSAPodval, то сначала чистим FSAPodval, а потом удаляем FSASensor
    Set shpChild = ShapeByHyperLink(shpParent.CellsU("Hyperlink.FSA.SubAddress").ResultStr(0))
    If Not shpChild Is Nothing Then
        ClearFSAPodvalChild shpChild
    End If

End Sub

Sub ClearFSAPodvalChild(shpChild As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearFSAPodvalChild - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("CrossReferenceSensor.ClearFSAPodvalChild", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
        'Чистим дочерний шейп
        shpChild.CellsU("Hyperlink.FSA.SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
        shpChild.CellsU("Hyperlink.FSA.ExtraInfo").FormulaForceU = ""
        shpChild.Shapes("Pomestu").CellsU("Prop.Place").FormulaForceU = "INDEX(0,Prop.Place.Format)"
        shpChild.Shapes("Pomestu").CellsU("Prop.Forma").FormulaForceU = "INDEX(0,Prop.Forma.Format)"
        shpChild.Shapes("Pomestu").CellsU("Prop.SymName").FormulaForceU = ""
        shpChild.Shapes("Pomestu").CellsU("Prop.Number").FormulaForceU = ""
        shpChild.Shapes("Pomestu").CellsU("Prop.NameKontur").FormulaForceU = ""

End Sub

Sub ClearFSAPodvalParent(shpParent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearFSAPodvalParent - Чистит родительский при копировании
                'Чистим ссылки в родительском при его копировании.
'------------------------------------------------------------------------------------------------------------
        'Чистим родительский шейп
        shpParent.CellsU("Hyperlink.FSA.SubAddress").FormulaForceU = """""" 'Пишем в ShapeSheet пустые кавычки. Если записать пустую строку, то будет NoFormula и нумерация контактов сломается
        shpParent.CellsU("Hyperlink.FSA.ExtraInfo").FormulaForceU = ""
        shpParent.CellsU("Prop.KanalNumber").FormulaForceU = 0
End Sub