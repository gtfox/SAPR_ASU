'------------------------------------------------------------------------------------------------------------
' Module        : AutoNumber - Автонумерация
' Author        : gtfox
' Date          : 2020.05.11
' Description   : Автонумерация/Перенумерация элементов схемы
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------





Public MaxNumber As Integer   'Максимальное значение нумерации существующих элементов. Это не общее число элементов, а макс цифра в обозначении.
Public MaxNumberFSA As Integer   'Максимальное значение нумерации существующих элементов. Это не общее число элементов, а макс цифра в обозначении.


'Sub EventDropAutoNum(vsoShapeEvent As Shape)
''------------------------------------------------------------------------------------------------------------
'' Macros        : EventDropAutoNum - Автонумерация для одиночной вставки
'                'Когда происходит вставка применяется привязка к курсору
'                'Если вставка была из набора элементов - привязка к курсору не происходит
'                '(после вставки на лист в щейпе ставится бит User.Dropped, и он начинает привязываться)
'                'В EventDrop должна быть формула =CALLTHIS("ThisDocument.EventDropAutoNum")
''------------------------------------------------------------------------------------------------------------
'    ..............
'End Sub

Public Sub AutoNum(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AutoNum - Автонумерация элементов при вбросе/копировании
                'Нумерация всегда продолжается с максимального значения нумерации существующих элементов
                'Если, в начале схемы был удален элемент, его номер больше не появится
                'Для лотания дыр в нумерации используйте перенумерацию элементов ReNumber()
                
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("AutoNumber.AutoNum", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'If ThisDocument.BlockMacros Then Exit Sub
    
    Dim SymName As String       'Буквенная часть нумерации
    Dim NomerShemy As Integer   'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
'    Dim MaxNumber As Integer   'Максимальное значение нумерации существующих элементов. Это не общее число элементов, а макс цифра в обозначении.

'    Dim TheDoc As Visio.Shape
'    Set TheDoc = Application.ActiveDocument.DocumentSheet
    
    Dim ThePage As Visio.Shape
    Set ThePage = ActivePage.PageSheet
    
    Dim vsoShapeOnPage As Visio.Shape

    Dim vsoPage As Visio.Page
    Dim PageName As String
    PageName = "Схема"  'Имена листов где возможна нумерация
    If ThePage.CellExists("User.NomerShemy", 0) Then NomerShemy = ThePage.Cells("User.NomerShemy").Result(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер.
    
    'Узнаем тип и буквенное обозначение элемента, который вставили на схему
    If vsoShape.CellExists("User.SAType", 0) Then UserType = vsoShape.Cells("User.SAType").Result(0)
    If vsoShape.CellExists("Prop.SymName", 0) Then SymName = vsoShape.Cells("Prop.SymName").ResultStr(0)
    
    'Чистим номер, чтобы он не участвовал в поиске
    vsoShape.Cells("Prop.Number").FormulaU = 0
    
    'Чистим максимум
    MaxNumber = 0

    'Цикл поиска максимального номера существующих элементов схемы
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If InStr(1, vsoPage.Name, PageName) > 0 Then    'Берем те, что содержат "Схема" в имени
            If vsoPage.PageSheet.Cells("User.NomerShemy").Result(0) = NomerShemy Then    'Берем все схемы с номером той, на которую вставляем элемент
                For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                    If vsoShapeOnPage.CellExists("User.SAType", 0) Then   'Если в шейпе есть тип, то -
                        If vsoShapeOnPage.Cells("User.SAType").Result(0) = UserType Then    '- проверяем чтобы совпадал с нашим (который вставили)
                            If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                                Select Case UserType
                                    Case typeWire 'Провода
                                        FindMAX vsoShapeOnPage
                                    Case typeCableSH 'Кабели на схеме электрической
                                        FindMAX vsoShapeOnPage
                                End Select
                                If (vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) = SymName) Then 'Буквы совпадают                     'And (vsoShapeOnPage.NameID <> vsoShape.NameID) и это не тот же шейп который вставили
                                    Select Case UserType
                                        Case typeTerminal 'Клеммы
                                            If vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) = vsoShape.Cells("Prop.NumberKlemmnik").Result(0) Then 'Выбираем клеммы из одного клеммника
                                                FindMAX vsoShapeOnPage
                                            End If
                                        Case typeCoil, typeParent, typeElement, typeTerminal, typeSensor, typeActuator, typeFSASensor, typeFSAPodval, typePLCParent, typeElectroPlan, typeElectroOneWire, typeOPSPlan 'Остальные элементы
                                            FindMAX vsoShapeOnPage
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next

    'Во вставленный элемент заносим максимальный найденный номер + 1
    vsoShape.Cells("Prop.Number").FormulaU = MaxNumber + 1
    
    'Активация событий. Они чета сомодезактивируются xD
    'Set vsoPagesEvent = ActiveDocument.Pages
    
End Sub

'Ищем максимальное значение номера элемента
Sub FindMAX(vsoShapeOnPage As Visio.Shape)
    If vsoShapeOnPage.Cells("Prop.Number").Result(0) > MaxNumber Then    'Ищем максимальное значение номера элемента
        MaxNumber = vsoShapeOnPage.Cells("Prop.Number").Result(0)    'Запоменаем. А те что меньше сюда не попадут
        'Debug.Print vsoShapeOnPage.Name & " " & MaxNumber
    End If
End Sub

Public Sub ReNumber()
'------------------------------------------------------------------------------------------------------------
' Macros        : ReNumber - Перенумерация элементов

                'Перенумерация происходит слева направо, сверху вниз
                'независимо от порядка появления элементов на схеме
                'и независимо от их номеров до перенумерации.
                'Если в элементе Prop.AutoNum=0 то он не участвует в перенумерации
                'Перенумерация элементов идет в пределах одной схемы (одного номера схемы)
'------------------------------------------------------------------------------------------------------------
    Dim shpElement As Shape
    Dim Prev As Shape
    Dim colShp As Collection
    Dim shpMas() As Shape
    Dim shpTemp As Shape
    Dim ss As String
    Dim i As Integer, ii As Integer, j As Integer, N As Integer
    
    Set colShp = New Collection
    
    'Собираем в коллекцию нужные для сортировки шейпы
    For Each shpElement In ActivePage.Shapes
        If shpElement.CellExists("User.SAType", False) Then
            If shpElement.Cells("User.SAType").Result(0) = typeCoil Then 'Будет задано из формы
                colShp.Add shpElement
                'Debug.Print shpElement.Cells("PinX").Result("mm") & " - " & shpElement.Cells("PinY").Result("mm")
            End If
        End If
    Next
    
    'из коллекции передаем их в массив для сортировки
    ReDim shpMas(colShp.Count - 1)
    i = 0
    For Each shpElement In colShp
        Set shpMas(i) = shpElement
        i = i + 1
    Next

    ' "Сортировка вставками" массива шейпов по возрастанию коордонаты Х
    '--V--Сортируем по возрастанию коордонаты Х
    UbMas = UBound(shpMas)
    For j = 1 To UbMas
        Set shpTemp = shpMas(j)
        i = j
        While shpMas(i - 1).Cells("PinX").Result("mm") > shpTemp.Cells("PinX").Result("mm") '>:возрастание, <:убывание
            Set shpMas(i) = shpMas(i - 1)
            i = i - 1
            If i <= 0 Then GoTo ExitWhileX
        Wend
ExitWhileX:  Set shpMas(i) = shpTemp
    Next
    '--Х--Сортировка по возрастанию коордонаты Х
    
    
'    Debug.Print "---"
'    For i = 0 To UbMas
'        Debug.Print shpMas(i).Cells("PinX").Result("mm") & " - " & shpMas(i).Cells("PinY").Result("mm")
'    Next
    
    'Находим шейпы с одинаковой координатой Х и сортируем Y-ки
    'Debug.Print "---"
    Group = False
    Set colShp = New Collection
    For ii = 1 To UbMas
        If (Abs(shpMas(ii - 1).Cells("PinX").Result("mm") - shpMas(ii).Cells("PinX").Result("mm")) < 0.5) And (ii < UbMas) Then
            'colShp.Add shpMas(i)
            If Group = False Then
                StartIndex = ii - 1 'На первом элементе запоменаем его номер
                Group = True    'Начали собирать одинакое координаты
            End If
            'Debug.Print shpMas(i).Cells("PinX").Result("mm") & " - " & shpMas(i).Cells("PinY").Result("mm")
        ElseIf Group Then
            'colShp.Add shpMas(i)
            Group = False   'Попался первый не одинаковый. Закончили.
            EndIndex = ii - 1
            If (ii = UbMas) And (Abs(shpMas(ii - 1).Cells("PinX").Result("mm") - shpMas(ii).Cells("PinX").Result("mm")) < 0.5) Then EndIndex = ii 'Если последний элемент, то включаем его в сортировку
           'Debug.Print shpMas(i).Cells("PinX").Result("mm") & " - " & shpMas(i).Cells("PinY").Result("mm")

            '--V--Сортируем по убыванию коордонаты Y
            For j = StartIndex + 1 To EndIndex
                Set shpTemp = shpMas(j)
                i = j
                While shpMas(i - 1).Cells("PinY").Result("mm") < shpTemp.Cells("PinY").Result("mm") '>:возрастание, <:убывание
                    Set shpMas(i) = shpMas(i - 1)
                    i = i - 1
                    If i <= StartIndex Then GoTo ExitWhileY
                Wend
ExitWhileY:     Set shpMas(i) = shpTemp
            Next
            '--Х--Сортировка по убыванию коордонаты Y
        End If
    Next
    Set colShp = Nothing
    
    'Перенумеровываем отсортированный массив
    For i = 0 To UbMas
        shpMas(i).Text = "KL" & (i + 1)
    Next

'    Debug.Print "---"
'    For i = 0 To UbMas
'        Debug.Print shpMas(i).Cells("PinX").Result("mm") & " - " & shpMas(i).Cells("PinY").Result("mm")
'    Next

    'Активация событий. Они чета сомодезактивируются xD
    Set vsoPagesEvent = ActiveDocument.Pages
    
End Sub

Sub AutoNumFSA(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AutoNumFSA - Автонумерация элементов на ФСА при вбросе/копировании
                'Нумеруются датчики с одинаковыми буквенными обозначениями (PT,TE,...) и в педелах одного контура (РТ/1-П,РТ/2-П,РТ/3-П)
                'Если контур не указан, то только с одинаковыми буквенными обозначениями (РТ/1,РТ/2,РТ/3)
                'Нумерация всегда продолжается с максимального значения нумерации существующих элементов
                'Если, в начале схемы был удален элемент, его номер больше не появится
                'Для лотания дыр в нумерации используйте перенумерацию элементов ReNumberFSA()
                
                'Когда происходит массовая вставка не применяется привязка к курсору
                'В EventMultiDrop должна быть формула = CALLTHIS("AutoNumber.AutoNumFSA", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    If BlockMacros Then Exit Sub
    
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim SymName As String       'Буквенная часть нумерации
    Dim NameKontur              'Имя контура
    Dim NomerFSA As Integer     'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    
'    Dim MaxNumber As Integer   'Максимальное значение нумерации существующих элементов. Это не общее число элементов, а макс цифра в обозначении.

'    Dim TheDoc As Visio.Shape
'    Set TheDoc = Application.ActiveDocument.DocumentSheet
    
    Dim ThePage As Visio.Shape
    Set ThePage = ActivePage.PageSheet
    
    Dim vsoShapeOnPage As Visio.Shape

    Dim vsoPage As Visio.Page
    Dim PageName As String
    PageName = "ФСА"  'Имена листов где возможна нумерация
    If ThePage.CellExists("User.NomerFSA", 0) Then NomerFSA = ThePage.Cells("User.NomerFSA").Result(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер. По умолчанию = 1
    
    'Узнаем тип и буквенное обозначение элемента, который вставили на схему
    If vsoShape.CellExists("User.SAType", 0) Then UserType = vsoShape.Cells("User.SAType").Result(0)
    If vsoShape.CellExists("Prop.SymName", 0) Then SymName = vsoShape.Cells("Prop.SymName").ResultStr(0)
    If vsoShape.CellExists("Prop.NameKontur", 0) Then NameKontur = vsoShape.Cells("Prop.NameKontur").ResultStr(0)
    
    
    
    'Чистим номер, чтобы он не участвовал в поиске
    vsoShape.Cells("Prop.Number").FormulaU = 0
    
    'Чистим максимум
    MaxNumberFSA = 0

    'Цикл поиска максимального номера существующих элементов схемы
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If InStr(1, vsoPage.Name, PageName) > 0 Then    'Берем те, что содержат "Схема" в имени
            If vsoPage.PageSheet.Cells("User.NomerFSA").Result(0) = NomerFSA Then    'Берем все схемы с номером той, на которую вставляем элемент
                For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                    If vsoShapeOnPage.CellExists("User.SAType", 0) Then   'Если в шейпе есть тип, то -
                        If vsoShapeOnPage.Cells("User.SAType").Result(0) = UserType Then    '- проверяем чтобы совпадал с нашим (который вставили)
                            If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                                    Select Case UserType
                                        Case typeFSAPodval
                                            FindMAXFSA vsoShapeOnPage
                                        End Select
                                If (vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) = SymName) Then 'Буквы совпадают
                                    Select Case UserType
                                        Case typeFSASensor 'датчики ФСА
                                            If vsoShapeOnPage.Cells("Prop.NameKontur").ResultStr(0) = vsoShape.Cells("Prop.NameKontur").ResultStr(0) Then 'Выбираем датчики из одного контура
                                                FindMAXFSA vsoShapeOnPage
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next

    'Во вставленный элемент заносим максимальный найденный номер + 1
    vsoShape.Cells("Prop.Number").FormulaU = MaxNumberFSA + 1
    
    'Активация событий. Они чета сомодезактивируются xD
    'Set vsoPagesEvent = ActiveDocument.Pages
    
End Sub

'Ищем максимальное значение номера элемента
Sub FindMAXFSA(vsoShapeOnPage As Visio.Shape)
    If vsoShapeOnPage.Cells("Prop.Number").Result(0) > MaxNumberFSA Then    'Ищем максимальное значение номера элемента
        MaxNumberFSA = vsoShapeOnPage.Cells("Prop.Number").Result(0)    'Запоменаем. А те что меньше сюда не попадут
        'Debug.Print vsoShapeOnPage.Name & " " & MaxNumberFSA
    End If
End Sub



Sub ReNumberFSA()

End Sub

Sub HideWireNumChildOnPage()
    HideWireNumChild ActivePage
End Sub

Sub HideWireNumChildInDoc()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    PageName = "Схема"  'Имена листов
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If InStr(1, vsoPage.Name, PageName) > 0 Then    'Берем те, что содержат "Схема" в имени
            HideWireNumChild vsoPage
        End If
    Next
End Sub


Public Sub HideWireNumChild(vsoPage As Visio.Page)
'------------------------------------------------------------------------------------------------------------
' Macros        : HideWireNumChild - Скрывает номера в дочерних проводах (номера полученные по ссылке)
                'На листе остаются только провода с уникальными именами
                'Номера ВСЕХ проводов нужны только при рисовании схемы - для контроля правильности соединения
'------------------------------------------------------------------------------------------------------------
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String
    Dim vsoShapeOnPage As Visio.Shape
    Dim ThePage As Visio.Shape
    Set ThePage = vsoPage.PageSheet
    
    PageName = "Схема"  'Имена листов где возможна нумерация
    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер. По умолчанию = 1
    If ThePage.CellExists("User.NomerShemy", 0) Then NomerShemy = ThePage.Cells("User.NomerShemy").Result(0)
    
    'Цикл поиска проводов и скрытия номера
    For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
        If vsoShapeOnPage.CellExists("User.SAType", 0) Then   'Если в шейпе есть тип, то -
            If vsoShapeOnPage.Cells("User.SAType").Result(0) = typeWire Then    '- проверяем чтобы был провод
                If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 0 Then    'Отсеиваем шейпы нумеруемые в автомате
                    If vsoShapeOnPage.Cells("Prop.Number").FormulaU Like "*!*" Then 'Находим дочерние
                        'Прячем номер/название
                        vsoShapeOnPage.Cells("Prop.HideNumber").FormulaU = True
                        vsoShapeOnPage.Cells("Prop.HideName").FormulaU = True
                    End If
                End If
            End If
        End If
    Next

End Sub

'Добавление на листы номера/названия схемы
Sub NomerShemyAdd()
    Dim ThePage As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim PageName As String
    PageName = "Схема"
    
'    For Each vsoPage In ActiveDocument.Pages
'        If InStr(1, vsoPage.Name, PageName) > 0 Then
'            Set ThePage = vsoPage.PageSheet
Set ThePage = ActivePage.PageSheet
            If Not ThePage.CellExists("Prop.NomerShemy", 0) Then
                'Prop
                ThePage.AddRow visSectionProp, visRowLast, visTagDefault
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "NomerShemy"
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Название схемы"""
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Нумерация элементов идет в пределах одной схемы"""
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "4"
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsFormat).FormulaForceU = "SETATREF(TheDoc!Prop.NazvaniayShemDocumenta.Format)"
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsValue).FormulaForceU = "INDEX(1,Prop.NomerShemy.Format)"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsSortKey).FormulaForceU = """"""
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsInvis).FormulaForceU = "FALSE"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsAsk).FormulaForceU = "FALSE"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsLangID).FormulaForceU = "1033"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsCalendar).FormulaForceU = "0"
            End If
'-------------------------------------------------------------------------------------------------------------
            If Not ThePage.CellExists("User.NomerShemy", 0) Then
                'User
                ThePage.AddRow visSectionUser, visRowLast, visTagDefault
                ThePage.CellsSRC(visSectionUser, visRowLast, visUserValue).RowNameU = "NomerShemy"
                ThePage.CellsSRC(visSectionUser, visRowLast, visUserValue).FormulaU = "LOOKUP(Prop.NomerShemy,TheDoc!Prop.NazvaniayShemDocumenta.Format)"
                ThePage.CellsSRC(visSectionUser, visRowLast, visUserPrompt).FormulaU = ""
            End If
'        End If
'    Next

End Sub

'Добавление на листы номера/названия схемы
Sub NomerFSA_Add()
    Dim ThePage As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim PageName As String
    PageName = "ФСА"
    
'    For Each vsoPage In ActiveDocument.Pages
'        If InStr(1, vsoPage.Name, PageName) > 0 Then
'            Set ThePage = vsoPage.PageSheet
Set ThePage = ActivePage.PageSheet
            If Not ThePage.CellExists("Prop.NomerFSA", 0) Then
                'Prop
                ThePage.AddRow visSectionProp, visRowLast, visTagDefault
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsValue).RowNameU = "NomerFSA"
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsLabel).FormulaForceU = """Название ФСА"""
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsPrompt).FormulaForceU = """Нумерация элементов идет в пределах одной схемы"""
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsType).FormulaForceU = "4"
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsFormat).FormulaForceU = "SETATREF(TheDoc!Prop.NazvaniayFSADocumenta.Format)"
                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsValue).FormulaForceU = "INDEX(1,Prop.NomerFSA.Format)"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsSortKey).FormulaForceU = """"""
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsInvis).FormulaForceU = "FALSE"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsAsk).FormulaForceU = "FALSE"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsLangID).FormulaForceU = "1033"
'                ThePage.CellsSRC(visSectionProp, visRowLast, visCustPropsCalendar).FormulaForceU = "0"
            End If
'-------------------------------------------------------------------------------------------------------------
            If Not ThePage.CellExists("User.NomerFSA", 0) Then
                'User
                ThePage.AddRow visSectionUser, visRowLast, visTagDefault
                ThePage.CellsSRC(visSectionUser, visRowLast, visUserValue).RowNameU = "NomerFSA"
                ThePage.CellsSRC(visSectionUser, visRowLast, visUserValue).FormulaU = "LOOKUP(Prop.NomerFSA,TheDoc!Prop.NazvaniayFSADocumenta.Format)"
                ThePage.CellsSRC(visSectionUser, visRowLast, visUserPrompt).FormulaU = ""
            End If
'        End If
'    Next

End Sub


''' запись значения списка в ячейки user-defined документа '''
Public Sub EditListValue(nameCell As String, numValue As Integer, newValue)
' 1 аргумент - имя ячейки для изменения
' 2 аргумент - номер значения для поиска в списке значений (начиная с 1)
' 3 аргумент - новое значение для замены
Dim arrList
Dim docSheet As Visio.Shape
Dim visDoc As Visio.Document
Set visDoc = Application.ActiveDocument
Set docSheet = visDoc.DocumentSheet

With docSheet.Cells(nameCell)
     arrList = Split(.FormulaU, ";")  ' создаем массив значений из формулы
     arrList(numValue - 1) = newValue ' меняем одно из значений
     .FormulaU = Join(arrList, ";")   ' создаем строку из значений массива и записываем назад в ячейку
End With

End Sub

''' чтение значения списка из ячейки user-defined документа '''
Public Function ReadListValue(nameCell As String, numValue As Integer)
' 1 аргумент - имя ячейки для чтения
' 2 аргумент - номер значения для поиска в списке значений (начиная с 1)
Dim arrList
Dim docSheet As Visio.Shape
Dim visDoc As Visio.Document
Set visDoc = Application.ActiveDocument
Set docSheet = visDoc.DocumentSheet

With docSheet.Cells(nameCell)
    arrList = Split(.FormulaU, ";")       ' создаем массив значений из формулы
    ReadListValue = arrList(numValue - 1) ' извлекаем значение
End With

End Function

Function ExtractOboz(Oboz) ' Функция определения неизменяемой части обозначения

Dim ObozF As String, i As Integer, Flag As Boolean
Flag = Oboz Like "*[-.,/\]*"

For i = 1 To Len(Oboz)
    If Not Flag And Mid(Oboz, i, 1) Like "[a-zA-Zа-яА-Я ]" Then GoSub AddChar
    If Flag And Mid(Oboz, i, 1) Like "[a-zA-Zа-яА-Я0-9 ]" Then GoSub AddChar
    If Flag And Mid(Oboz, i, 1) Like "[-.,/\]" Then GoSub AddChar
Next
    
ExtractOboz = ObozF
Exit Function

AddChar:
    ObozF = ObozF + Mid(Oboz, i, 1)
Return
End Function ' ***************************** AutoNum *************************

