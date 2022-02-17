'------------------------------------------------------------------------------------------------------------
' Module        : AutoNumber - Автонумерация
' Author        : gtfox
' Date          : 2020.05.11
' Description   : Автонумерация/Перенумерация элементов схемы
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
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
'    Макрос в ThisDocument ..............
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
    
    Dim SymName As String       'Буквенная часть нумерации
    Dim NazvanieShemy As String   'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim ThePage As Visio.Shape
    Dim vsoShapeOnPage As Visio.Shape
    Dim vsoPage As Visio.Page
    Dim PageName As String
    
    Set ThePage = ActivePage.PageSheet
    If ThePage.CellExists("Prop.SA_NazvanieShemy", 0) Then NazvanieShemy = ThePage.Cells("Prop.SA_NazvanieShemy").ResultStr(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер.
    PageName = cListNameCxema  'Имена листов где возможна нумерация
    'Узнаем тип и буквенное обозначение элемента, который вставили на схему
    UserType = ShapeSAType(vsoShape)
    If vsoShape.CellExists("Prop.SymName", 0) Then SymName = vsoShape.Cells("Prop.SymName").ResultStr(0)
    
    'Чистим номер, чтобы он не участвовал в поиске
    vsoShape.Cells("Prop.Number").FormulaU = 0
    
    'Чистим максимум
    MaxNumber = 0

    'Цикл поиска максимального номера существующих элементов схемы
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If Left(vsoPage.name, Len(PageName)) = PageName Then    'Берем те, что содержат "Схема" в имени
            If vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) = NazvanieShemy Then    'Берем все схемы с именем той, на которую вставляем элемент
                For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                    If ShapeSATypeIs(vsoShapeOnPage, UserType) Then     'Если в шейпе есть тип, то проверяем чтобы совпадал с нашим (который вставили)
                        If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 1 Then    'Отсеиваем шейпы нумеруемые вручную
                            Select Case UserType
                                Case typeWire 'Провода
                                    FindMAX vsoShapeOnPage
                                Case typeCableSH 'Кабели на схеме электрической
                                    FindMAX vsoShapeOnPage
                            End Select
                            If (vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) = SymName) Then 'Буквы совпадают                     'And (vsoShapeOnPage.NameID <> vsoShape.NameID) и это не тот же шейп который вставили
                                Select Case UserType
                                    Case typeTerm 'Клеммы
                                        If vsoShapeOnPage.Cells("Prop.NumberKlemmnik").Result(0) = vsoShape.Cells("Prop.NumberKlemmnik").Result(0) Then 'Выбираем клеммы из одного клеммника
                                            FindMAX vsoShapeOnPage
                                        End If
                                    Case typeCoil, typeParent, typeElement, typePLCParent, typeSensor, typeActuator, typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                                        FindMAX vsoShapeOnPage
                                End Select
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

Sub ShowReNumber()
    frmReNumber.Show
End Sub

Public Function ReNumber(colShp As Collection, StartNumber As Integer) As Integer
'------------------------------------------------------------------------------------------------------------
' Macros        : ReNumber - Перенумерация элементов
                'Сортировка массивов координат и перенумерация
'------------------------------------------------------------------------------------------------------------
    Dim shpElement As Shape
    Dim masShape() As Shape
    Dim shpTemp As Shape
    Dim XPred As Double
    Dim XTekush As Double
    Dim i As Integer, ii As Integer, j As Integer, n As Integer

    'из коллекции передаем их в массив для сортировки
    ReDim masShape(colShp.Count - 1)
    i = 0
    For Each shpElement In colShp
        Set masShape(i) = shpElement
        i = i + 1
    Next

    ' "Сортировка вставками" массива шейпов по возрастанию коордонаты Х
    '--V--Сортируем по возрастанию коордонаты Х
    UbMas = UBound(masShape)
    For j = 1 To UbMas
        Set shpTemp = masShape(j)
        i = j
        If SAType = typeWire Then
            While WireX(masShape(i - 1)) > WireX(shpTemp) '>:возрастание, <:убывание
                Set masShape(i) = masShape(i - 1)
                i = i - 1
                If i <= 0 Then GoTo ExitWhileX
            Wend
        Else
            While masShape(i - 1).Cells("PinX").Result("mm") > shpTemp.Cells("PinX").Result("mm") '>:возрастание, <:убывание
                Set masShape(i) = masShape(i - 1)
                i = i - 1
                If i <= 0 Then GoTo ExitWhileX
            Wend
        End If

ExitWhileX:  Set masShape(i) = shpTemp
    Next
    '--Х--Сортировка по возрастанию коордонаты Х

    'Находим шейпы с одинаковой координатой Х и сортируем Y-ки
    Group = False
    Set colShp = New Collection
    For ii = 1 To UbMas
        If SAType = typeWire Then
            XPred = WireX(masShape(ii - 1))
            XTekush = WireX(masShape(ii))
        Else
            XPred = masShape(ii - 1).Cells("PinX").Result("mm")
            XTekush = masShape(ii).Cells("PinX").Result("mm")
        End If
        
        If (Abs(XPred - XTekush) < 0.5) And (ii < UbMas) Then
            If Group = False Then
                StartIndex = ii - 1 'На первом элементе запоменаем его номер
                Group = True    'Начали собирать одинаковые координаты
            End If
        ElseIf Group Then
            Group = False   'Попался первый не одинаковый. Закончили.
            EndIndex = ii - 1
            If (ii = UbMas) And (Abs(XPred - XTekush) < 0.5) Then EndIndex = ii 'Если последний элемент, то включаем его в сортировку

            '--V--Сортируем по убыванию коордонаты Y
            For j = StartIndex + 1 To EndIndex
                Set shpTemp = masShape(j)
                i = j
                If SAType = typeWire Then
                    While WireY(masShape(i - 1)) < WireY(shpTemp) '>:возрастание, <:убывание
                        Set masShape(i) = masShape(i - 1)
                        i = i - 1
                        If i <= StartIndex Then GoTo ExitWhileY
                    Wend
                Else
                    While masShape(i - 1).Cells("PinY").Result("mm") < shpTemp.Cells("PinY").Result("mm") '>:возрастание, <:убывание
                        Set masShape(i) = masShape(i - 1)
                        i = i - 1
                        If i <= StartIndex Then GoTo ExitWhileY
                    Wend
                End If
ExitWhileY:     Set masShape(i) = shpTemp
            Next
            '--Х--Сортировка по убыванию коордонаты Y
        End If
    Next
    Set colShp = Nothing
    
    'Перенумеровываем отсортированный массив
    For i = 0 To UbMas
        masShape(i).Cells("Prop.Number").FormulaU = StartNumber + i + 1
    Next
    
    ReNumber = masShape(UbMas).Cells("Prop.Number").Result(0)
    
End Function

Function WireX(vsoShape As Visio.Shape) As Double
    Dim BeginX As Double
    Dim EndX As Double
    BeginX = vsoShape.Cells("BeginX").Result("mm")
    EndX = vsoShape.Cells("EndX").Result("mm")
    WireX = IIf(BeginX < EndX, BeginX, EndX) ' Начало провода по X = Слева
End Function

Function WireY(vsoShape As Visio.Shape) As Double
    Dim BeginY As Double
    Dim EndY As Double
    BeginY = vsoShape.Cells("BeginY").Result("mm")
    EndY = vsoShape.Cells("EndY").Result("mm")
    WireY = IIf(BeginY > EndY, BeginY, EndY) ' Начало провода по Y = Сверху
End Function

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
    If vsoShape Is Nothing Or vsoShape.id = 0 Then Exit Sub
    
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim SymName As String       'Буквенная часть нумерации
    Dim NameKontur              'Имя контура
    Dim NazvanieFSA As String     'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    
'    Dim MaxNumber As Integer   'Максимальное значение нумерации существующих элементов. Это не общее число элементов, а макс цифра в обозначении.

'    Dim TheDoc As Visio.Shape
'    Set TheDoc = Application.ActiveDocument.DocumentSheet
    
    Dim ThePage As Visio.Shape
    Set ThePage = ActivePage.PageSheet
    
    Dim vsoShapeOnPage As Visio.Shape

    Dim vsoPage As Visio.Page
    Dim PageName As String
    PageName = cListNameFSA  'Имена листов где возможна нумерация
    If ThePage.CellExists("Prop.SA_NazvanieFSA", 0) Then NazvanieFSA = ThePage.Cells("Prop.SA_NazvanieFSA").ResultStr(0)    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер. По умолчанию = 1
    
    'Узнаем тип и буквенное обозначение элемента, который вставили на схему
    UserType = ShapeSAType(vsoShape)
    If vsoShape.CellExists("Prop.SymName", 0) Then SymName = vsoShape.Cells("Prop.SymName").ResultStr(0)
    If vsoShape.CellExists("Prop.NameKontur", 0) Then NameKontur = vsoShape.Cells("Prop.NameKontur").ResultStr(0)
    
    
    
    'Чистим номер, чтобы он не участвовал в поиске
    vsoShape.Cells("Prop.Number").FormulaU = 0
    
    'Чистим максимум
    MaxNumberFSA = 0

    'Цикл поиска максимального номера существующих элементов схемы
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If InStr(1, vsoPage.name, PageName) > 0 Then    'Берем те, что содержат "Схема" в имени
            If vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0) = NazvanieFSA Then    'Берем все схемы с номером той, на которую вставляем элемент
                For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы в найденных листах
                    If ShapeSATypeIs(vsoShapeOnPage, UserType) Then      'Если в шейпе есть тип, то проверяем чтобы совпадал с нашим (который вставили)
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
    PageName = cListNameCxema  'Имена листов
    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        If InStr(1, vsoPage.name, PageName) > 0 Then    'Берем те, что содержат "Схема" в имени
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
    
    PageName = cListNameCxema  'Имена листов где возможна нумерация
    'Номер схемы. Если одна схема на весь проект, то на всех листах должен быть один номер. По умолчанию = 1
    If ThePage.CellExists("Prop.SA_NazvanieShemy", 0) Then NazvanieShemy = ThePage.Cells("Prop.SA_NazvanieShemy").ResultStr(0)
    
    'Цикл поиска проводов и скрытия номера
    For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
        If ShapeSATypeIs(vsoShapeOnPage, typeWire) Then     'Если в шейпе есть тип, то проверяем чтобы был провод
            If vsoShapeOnPage.Cells("Prop.AutoNum").Result(0) = 0 Then    'Отсеиваем шейпы нумеруемые в автомате
                If vsoShapeOnPage.Cells("Prop.Number").FormulaU Like "*!*" Then 'Находим дочерние
                    'Прячем номер/название
                    vsoShapeOnPage.Cells("Prop.HideNumber").FormulaU = True
                    vsoShapeOnPage.Cells("Prop.HideName").FormulaU = True
                End If
            End If
        End If
    Next

End Sub

'------------------------------------------------------------------------------------------------------------
' Macros        : ExtractOboz - Функция определения неизменяемой части обозначения
' Author        : Shishok
' Date          : 2014.12.01
' Description   : Определения неизменяемой части обозначения Например: 1, ГР1, р, Гр1.1, ППР1-1, Выкл, П122.1 или типа того
' Link          : https://visio.getbb.ru/viewtopic.php?p=5904#p5904, https://github.com/shishok, https://disk.yandex.ru/d/qbpj9WI9d2eqF
'------------------------------------------------------------------------------------------------------------
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
End Function

