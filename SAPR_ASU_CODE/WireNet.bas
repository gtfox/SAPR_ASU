'------------------------------------------------------------------------------------------------------------
' Module        : WireNet - Провода на схеме электрической принципиальной
' Author        : gtfox
' Date          : 2020.06.03
' Description   : Соединение/отсоединение проводов, нумерация, удаление, стрелки/точки на концах, взаимодействие с элементами
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Public bUnGlue As Boolean 'Запрет обработки события ConnectionsDeleted


Sub ConnectWire(Connects As IVConnects)
'------------------------------------------------------------------------------------------------------------
' Macros        : ConnectWire - Цепляет провод к проводу или элементу
                'В зависимости от того к чему цепляем: провод, разрыв провода или элемент, производим заполнение
                'полей имени, номера, источника номера. Для проводов ставим точку на конце, для элементов - убираем красную стрелку
                'Макрос вызывается событием ConnectionsAdded
'------------------------------------------------------------------------------------------------------------
    Dim shpProvod As Visio.Shape
    Dim RefNashegoProvoda As String, AdrNashegoProvoda As String, RefSource As String, AdrSource As String, kogo As String, kem As String
    Dim i As Integer, ii As Integer
    Dim ShapeType As Integer
    Dim ShapeTypeNaDrugomKonce As Integer
    
    Set shpProvod = Connects.FromSheet
    
    RefSource = "Pages[" & Connects.ToSheet.ContainingPage.NameU & "]!" & Connects.ToSheet.NameID 'Адрес источника нумерации (к которому подключаемся)
    RefNashegoProvoda = "Pages[" & shpProvod.ContainingPage.NameU & "]!" & shpProvod.NameID 'Адрес нашего (подключаемого) провода
    AdrSource = Connects.ToSheet.ContainingPage.NameU & "/" & Connects.ToSheet.NameID
    AdrNashegoProvoda = shpProvod.ContainingPage.NameU & "/" & shpProvod.NameID
    
    ShapeType = Connects.ToSheet.Cells("User.SAType").Result(0) 'Тип шейпа, к которому подсоединили провод
    
    Select Case shpProvod.Connects.Count 'кол-во соединенных концов у провода
    
        Case 1 'С одной стороны
        
            'Если шейп, к которому подсоединили провод - оказался тоже провод или конечный разрыв провода (дочерний)
            If (ShapeType = typeWire) Or (ShapeType = typeWireLinkR) Then

                shpProvod.Cells("Prop.Number").FormulaU = RefSource & "!Prop.Number" 'Получаем номер от существующего провода (к которому подсоединились)
                shpProvod.Cells("Prop.SymName").FormulaU = RefSource & "!Prop.SymName" 'Получаем имя от существующего провода (к которому подсоединились)
                shpProvod.Cells("Prop.AutoNum").FormulaU = False 'Убираем автонумерацию (т.к. номер получаем по ссылке от другого провода)
                If ShapeType = typeWire Then
                    SetArrow 10, Connects(1) 'Ставим точку если это провод а не разрыв
                ElseIf ShapeType = typeWireLinkR Then
                    SetArrow 0, Connects(1) 'Убираем стрелку
                End If
                shpProvod.Cells("User.AdrSource").FormulaU = Chr(34) & AdrSource & Chr(34) 'Сохраняем адрес источника номера
                'shpProvod.Cells("Prop.HideNumber").FormulaU = True 'Скрываем номер (возможно)
                'shpProvod.Cells("Prop.HideName").FormulaU = True 'Скрываем название (возможно)
            Else
            'Если шейп, к которому подсоединили провод - оказался НЕ провод... (элемент)
                SetArrow 0, Connects(1) 'Убираем стрелку
                
                'Пишем номер провода в родительский ПЛК
                If ShapeType = typePLCTerm Then
                    WireToPLCTerm shpProvod, Connects.ToSheet, True
                End If
                
                'Если это начальный разрыв провода (родительский) - присваиваем ему имя и номер провода
                If ShapeType = typeWireLinkS Then
                    Connects.ToSheet.Cells("Prop.Number").FormulaU = RefNashegoProvoda & "!Prop.Number" 'Записываем номер нашего провода
                    Connects.ToSheet.Cells("Prop.SymName").FormulaU = RefNashegoProvoda & "!Prop.SymName" 'Записываем имя нашего провода
                    Connects.ToSheet.Cells("User.AdrSource").FormulaU = Chr(34) & AdrNashegoProvoda & Chr(34) 'Сохраняем адрес источника номера
                End If
            End If
            
        Case 2 'С двух сторон
        
            'Находим тип шейпа, на друм конце нашего провода
            For i = 1 To shpProvod.Connects.Count 'смотрим все соединения (их 2 :) )
                If shpProvod.Connects(i).FromPart <> Connects(1).FromPart Then 'Отбрасывам то, которое только что произошло (берем другой конец)
                    If shpProvod.Connects(i).ToSheet.CellExistsU("User.SAType", 0) Then
                        ShapeTypeNaDrugomKonce = shpProvod.Connects(i).ToSheet.Cells("User.SAType").Result(0) 'Тип шейпа, на друм конце нашего провода
                    End If
                End If
            Next
            
            'Если шейп, к которому подсоединили провод - оказался тоже провод или конечный разрыв провода (дочерний)
            If (ShapeType = typeWire) Or (ShapeType = typeWireLinkR) Then
            
                If ShapeType = typeWire Then
                    SetArrow 10, Connects(1) 'Ставим точку если это провод а не разрыв
                ElseIf ShapeType = typeWireLinkR Then
                    SetArrow 0, Connects(1) 'Убираем стрелку
                End If
                
                'если другой конец подсоединен НЕ к проводу - получаем номер от провода к которому подсоединились
                If (ShapeTypeNaDrugomKonce <> typeWire) And (ShapeTypeNaDrugomKonce <> typeWireLinkR) Then 'Смотрим что на другом конце НЕ провод и НЕ конечный разрыв провода (дочерний)
               
                    shpProvod.Cells("Prop.Number").FormulaU = RefSource & "!Prop.Number" 'Получаем номер от существующего провода (к которому подсоединились)
                    shpProvod.Cells("Prop.SymName").FormulaU = RefSource & "!Prop.SymName" 'Получаем имя от существующего провода (к которому подсоединились)
                    shpProvod.Cells("Prop.AutoNum").FormulaU = False 'Убираем автонумерацию (т.к. номер получаем по ссылке от другого провода)
                    shpProvod.Cells("User.AdrSource").FormulaU = Chr(34) & AdrSource & Chr(34) 'Сохраняем адрес источника номера
'                    shpProvod.Cells("Prop.HideNumber").FormulaU = True 'Скрываем номер (возможно)
'                    shpProvod.Cells("Prop.HideName").FormulaU = True 'Скрываем название (возможно)
                Else
                'если другой конец подсоединен к проводу - то проводу, к которому подсоединились, присваиваем номер от нашего присоединенного провода
                    kogo = Connects.ToSheet.Cells("Prop.Number").Result(0) & ": " & Connects.ToSheet.Cells("Prop.SymName").ResultStr(0)
                    kem = shpProvod.Cells("Prop.Number").Result(0) & ": " & shpProvod.Cells("Prop.SymName").ResultStr(0)

                    If MsgBox("Перезаписать провод" & vbCrLf & vbCrLf & kem & " -> " & kogo, vbOKCancel + vbExclamation, "Перезапись провода") = vbOK Then
                    
                        If ShapeType = typeWireLinkR Then 'Нельзя перезаписать "приемник разрыва провода" (дочерний), т.к. номер ему присвоен от "источника разрыва провода" (родителя)
                        
                            MsgBox "Нельзя перезаписать ""Приемник разрыва провода"" (дочерний), т.к. номер ему присвоен от ""Источника разрыва провода"" (родителя)" & vbCrLf & vbCrLf & kem & " -X- " & kogo, vbOKOnly + vbCritical, "Перезапись провода"
                            SetArrow 254, Connects(1) 'Возвращаем красную стрелку
                            UnGlue Connects(1) 'Отклеиваем конец

                        ElseIf Connects.ToSheet.Cells("Prop.Number").Result(0) = shpProvod.Cells("Prop.Number").Result(0) Then 'Номера проводов совпадают
                        
                            MsgBox "Номера проводов совпадают" & vbCrLf & vbCrLf & kem & " -X- " & kogo, vbOKOnly + vbCritical, "Перезапись провода"
                            SetArrow 254, Connects(1) 'Возвращаем красную стрелку
                            UnGlue Connects(1) 'Отклеиваем конец

                        ElseIf Connects.ToSheet.Cells("Prop.Number").FormulaU Like "*!*" Then 'Нельзя перезаписать номер провода полученный по ссылке от друго провода
                        
                            MsgBox "Нельзя перезаписать номер провода полученный по ссылке от друго провода" & vbCrLf & vbCrLf & kem & " -X- " & kogo, vbOKOnly + vbCritical, "Перезапись провода"
                            SetArrow 254, Connects(1) 'Возвращаем красную стрелку
                            UnGlue Connects(1) 'Отклеиваем конец
                       
                        Else
                        
                            'Ничего не мешает перезаписать провод
                            Connects.ToSheet.Cells("Prop.Number").FormulaU = RefNashegoProvoda & "!Prop.Number" 'Записывам номер подключаемого провода в существующий (к которому подсоединились)
                            Connects.ToSheet.Cells("Prop.SymName").FormulaU = RefNashegoProvoda & "!Prop.SymName" 'Записывам имя подключаемого провода в существующий (к которому подсоединились)
                            Connects.ToSheet.Cells("Prop.AutoNum").FormulaU = False 'Убираем автонумерацию (т.к. номер получаем по ссылке от другого провода)
                            Connects.ToSheet.Cells("User.AdrSource").FormulaU = Chr(34) & AdrNashegoProvoda & Chr(34) 'Сохраняем адрес источника номера
'                            Connects.ToSheet.Cells("Prop.HideNumber").FormulaU = True 'Скрываем номер (возможно)
'                            Connects.ToSheet.Cells("Prop.HideName").FormulaU = True 'Скрываем название (возможно)
                        End If
                    Else    'Если отказались перезаписывать провод
                        SetArrow 254, Connects(1) 'Возвращаем красную стрелку
                        UnGlue Connects(1) 'Отклеиваем конец
                    End If
                End If
            Else
            'Если шейп, к которому подсоединили провод - оказался НЕ провод... (элемент)
            
                'если другой конец подсоединен к проводу - только убираем стрелку
                SetArrow 0, Connects(1) 'Убираем стрелку
                
                'Если это начальный разрыв провода (родительский) - присваиваем ему имя и номер провода
                If ShapeType = typeWireLinkS Then
                    Connects.ToSheet.Cells("Prop.Number").FormulaU = RefNashegoProvoda & "!Prop.Number" 'Записываем номер нашего провода
                    Connects.ToSheet.Cells("Prop.SymName").FormulaU = RefNashegoProvoda & "!Prop.SymName" 'Записываем имя нашего провода
                    Connects.ToSheet.Cells("User.AdrSource").FormulaU = Chr(34) & AdrNashegoProvoda & Chr(34) 'Сохраняем адрес источника номера
                End If
                
                'если другой конец подсоединен НЕ к проводу и НЕ к конечному разрыву провода (дочернему) - присваиваем номер проводу
                If (ShapeTypeNaDrugomKonce <> typeWire) And (ShapeTypeNaDrugomKonce <> typeWireLinkR) Then 'Смотрим что на другом конце НЕ провод и НЕ конечный разрыв провода (дочерний)
                    'Присваиваем номер проводу
                    shpProvod.Cells("Prop.SymName").FormulaU = "" 'Чистим название провода
                    shpProvod.Cells("Prop.AutoNum").FormulaU = True 'Включаем автонумерацию (т.к. это независимый провод)
                    shpProvod.Cells("Prop.HideNumber").FormulaU = False 'Показываем номер
                    shpProvod.Cells("Prop.HideName").FormulaU = True 'Скрываем название
                    'Присваиваем номер проводу
                    AutoNum shpProvod

                End If
                
                'Пишем номер провода в родительский ПЛК
                If ShapeType = typePLCTerm Then
                    WireToPLCTerm shpProvod, Connects.ToSheet, True
                End If
                
            End If
        Case Else
    End Select
    
    'Ищем Дочерних которые ссылаются не нас - отцепляем
    FindZombie shpProvod
    
End Sub


Sub DisconnectWire(Connects As IVConnects)
'------------------------------------------------------------------------------------------------------------
' Macros        : DisconnectWire - Отцепляет провод от провода или элемента
                'В зависимости от того от чего отцепляем: провода, разрыва провода или элемента, производим чистку
                'полей имени, номера, источника номера. Убираем точку на конце и возвращаем красную стрелку
                'Макрос вызывается событием ConnectionsDeleted
'------------------------------------------------------------------------------------------------------------
    Dim shpProvod As Visio.Shape
    Dim AdrNashegoProvoda As String, AdrSource As String, AdrNaDrugomKonce As String
    Dim i As Integer, ii As Integer
    Dim ShapeType As Integer
    Dim ShapeTypeNaDrugomKonce As Integer
    
    Set shpProvod = Connects.FromSheet
    
    If bUnGlue Then bUnGlue = False: Exit Sub
    
    AdrSource = Connects.ToSheet.ContainingPage.NameU & "/" & Connects.ToSheet.NameID
    AdrNashegoProvoda = shpProvod.ContainingPage.NameU & "/" & shpProvod.NameID
    
    ShapeType = Connects.ToSheet.Cells("User.SAType").Result(0) 'Тип шейпа, от которого отсоединили провод
    
    Select Case shpProvod.Connects.Count 'кол-во соединенных концов у провода
    
        Case 0 'С одной стороны
        
            'Оторвали от Любого (Источник номера (Провод или >- или Элемент) или Элемент)
            
            'Чистим наш
            shpProvod.Cells("Prop.Number").FormulaU = ""
            shpProvod.Cells("Prop.SymName").FormulaU = ""
            shpProvod.Cells("Prop.AutoNum").FormulaU = False
            shpProvod.Cells("User.AdrSource").FormulaU = ""
            SetArrow 254, Connects(1) 'Возвращаем красную стрелку
            shpProvod.Cells("Prop.HideNumber").FormulaU = False
            shpProvod.Cells("Prop.HideName").FormulaU = True
            
            'Пишем 0 в номер провода в родительский ПЛК
            If ShapeType = typePLCTerm Then
                WireToPLCTerm shpProvod, Connects.ToSheet, False
            End If
                
            'Но если он еще и Дочерний (Оторвали от Дочернего (Провод или ->))
            If (ShapeType = typeWire) Or (ShapeType = typeWireLinkS) Then
                If Connects.ToSheet.Cells("User.AdrSource").ResultStr(0) = AdrNashegoProvoda Then 'Дочерний?
                'Чистим Дочерний
    
                    Connects.ToSheet.Cells("Prop.Number").FormulaU = ""
                    Connects.ToSheet.Cells("Prop.SymName").FormulaU = ""
                    Connects.ToSheet.Cells("User.AdrSource").FormulaU = ""
                    
                    'Если это был провод - то + автонумерация дочернего провода
                    If ShapeType = typeWire Then
                        Connects.ToSheet.Cells("Prop.AutoNum").FormulaU = False
                        Connects.ToSheet.Cells("Prop.HideNumber").FormulaU = False
                        Connects.ToSheet.Cells("Prop.HideName").FormulaU = True
                        If Connects.ToSheet.Connects.Count = 2 Then
                            Connects.ToSheet.Cells("Prop.AutoNum").FormulaU = False
                            'Присваиваем номер проводу
                            AutoNum Connects.ToSheet
                        End If
                    End If
                 End If
            End If
            
        Case 1, 2 '1 - С двух сторон, 2 - С двух сторон, но в момент быстрого переприклеивания провода
            
            'Оторвали от Провода или ->
            If (ShapeType = typeWire) Or (ShapeType = typeWireLinkS) Then
                'От Дочернего
                If Connects.ToSheet.Cells("User.AdrSource").ResultStr(0) = AdrNashegoProvoda Then 'Дочерний?
                
                    'Чистим Дочерний
                    Connects.ToSheet.Cells("Prop.Number").FormulaU = ""
                    Connects.ToSheet.Cells("Prop.SymName").FormulaU = ""
                    Connects.ToSheet.Cells("User.AdrSource").FormulaU = ""
                    SetArrow 254, Connects(1) 'Возвращаем красную стрелку
                    
                    'Если это был провод - то + автонумерация дочернего провода
                    If ShapeType = typeWire Then
                        Connects.ToSheet.Cells("Prop.AutoNum").FormulaU = False
                        Connects.ToSheet.Cells("Prop.HideNumber").FormulaU = False
                        Connects.ToSheet.Cells("Prop.HideName").FormulaU = True
                        If Connects.ToSheet.Connects.Count = 2 Then
                            Connects.ToSheet.Cells("Prop.AutoNum").FormulaU = True
                            'Присваиваем номер проводу
                            AutoNum Connects.ToSheet
                        End If
                    End If
                Else
                'От НЕ Дочернего
                    'Чистим наш
                    shpProvod.Cells("Prop.Number").FormulaU = ""
                    shpProvod.Cells("Prop.SymName").FormulaU = ""
                    shpProvod.Cells("User.AdrSource").FormulaU = ""
                    shpProvod.Cells("Prop.AutoNum").FormulaU = False
                    SetArrow 254, Connects(1) 'Возвращаем красную стрелку
                    shpProvod.Cells("Prop.HideNumber").FormulaU = False
                    shpProvod.Cells("Prop.HideName").FormulaU = True
                End If
            Else
            'Оторвали от Любого (Источник номера (Провод или >- или Элемент) или Элемент)
                'Находим шейп, на друм конце нашего провода
                For i = 1 To shpProvod.Connects.Count 'смотрим все соединения (их 2 :) )
                   If shpProvod.Connects(i).FromPart <> Connects(1).FromPart Then 'Отбрасывам то, которое только что произошло (берем другой конец)
                       AdrNaDrugomKonce = shpProvod.Connects(i).ToSheet.ContainingPage.NameU & "/" & shpProvod.Connects(i).ToSheet.NameID 'Адрес шейпа, на друм конце нашего провода
                       If shpProvod.Cells("User.AdrSource").ResultStr(0) <> AdrNaDrugomKonce Then 'Проверка на то что мы сами не являемся дочерним и на другом конце не провод или >-
                            'Чистим наш
                            shpProvod.Cells("Prop.Number").FormulaU = ""
                            shpProvod.Cells("Prop.SymName").FormulaU = ""
                            shpProvod.Cells("User.AdrSource").FormulaU = ""
                       End If
                   End If
                Next
                'являемся дочерним
                shpProvod.Cells("Prop.AutoNum").FormulaU = False
                SetArrow 254, Connects(1) 'Возвращаем красную стрелку
                'shpProvod.Cells("Prop.HideNumber").FormulaU = False
                'shpProvod.Cells("Prop.HideName").FormulaU = True
                
                'Пишем 0 в номер провода в родительский ПЛК
                If ShapeType = typePLCTerm Then
                    WireToPLCTerm shpProvod, Connects.ToSheet, False
                End If
                
            End If

    End Select
    
    'Ищем Дочерних которые ссылаются не нас - отцепляем
    FindZombie shpProvod
    
End Sub


Sub DeleteWire(DeletedShape As IVShape)
'------------------------------------------------------------------------------------------------------------
' Macros        : DeleteWire - Удаляет провод
                'Перебераем элементы секций Connects и FromConnects, производим чистку
                'полей имени, номера, источника номера.
                'У подключенных к нам проводов убираем точку на конце и возвращаем красную стрелку
                'Макрос вызывается событием BeforeShapeDelete
'------------------------------------------------------------------------------------------------------------
    Dim DeletedConnect As Visio.connect
    Dim ConnectedShape As Visio.Shape
    Dim i As Integer, ii As Integer
    Dim AdrNashegoProvoda As String
    Dim ShapeType As Integer

    AdrNashegoProvoda = DeletedShape.ContainingPage.NameU & "/" & DeletedShape.NameID
    
    'Перебор Connects
    For i = 1 To DeletedShape.Connects.Count
        Set DeletedConnect = DeletedShape.Connects(i)
        Set ConnectedShape = DeletedConnect.ToSheet
        
        ShapeType = ConnectedShape.Cells("User.SAType").Result(0)
        
        If (ShapeType = typeWire) Or (ShapeType = typeWireLinkS) Then
            If ConnectedShape.Cells("User.AdrSource").ResultStr(0) = AdrNashegoProvoda Then
                'Чистим Дочерний
                ConnectedShape.Cells("Prop.Number").FormulaU = ""
                ConnectedShape.Cells("Prop.SymName").FormulaU = ""
                ConnectedShape.Cells("User.AdrSource").FormulaU = ""

                'Если это был провод - то + автонумерация дочернего провода
                If ShapeType = typeWire Then
                    ConnectedShape.Cells("Prop.AutoNum").FormulaU = False
                    ConnectedShape.Cells("Prop.HideNumber").FormulaU = False
                    ConnectedShape.Cells("Prop.HideName").FormulaU = True
                    If ConnectedShape.Connects.Count = 2 Then
                        ConnectedShape.Cells("Prop.AutoNum").FormulaU = True
                        'Присваиваем номер проводу
                        AutoNum ConnectedShape
                    Else
                        If ConnectedShape.Connects.Count = 1 Then
                            SetArrow 254, ConnectedShape.Connects(1) 'Возвращаем красную стрелку
                            UnGlue ConnectedShape.Connects(1) 'Отклеиваем конец
                        End If
                    End If
                End If
            End If
        End If
        
        'Пишем 0 в номер провода в родительский ПЛК
        If ShapeType = typePLCTerm Then
            WireToPLCTerm DeletedShape, DeletedConnect.ToSheet, False
        End If
        
    Next
    'Перебор FromConnects
    For i = 1 To DeletedShape.FromConnects.Count
        Set DeletedConnect = DeletedShape.FromConnects(i)
        Set ConnectedShape = DeletedConnect.FromSheet
        
        ShapeType = ConnectedShape.Cells("User.SAType").Result(0)
        
        If (ShapeType = typeWire) Or (ShapeType = typeWireLinkS) Then
            If ConnectedShape.Cells("User.AdrSource").ResultStr(0) = AdrNashegoProvoda Then
                'Чистим Дочерний
                ConnectedShape.Cells("Prop.Number").FormulaU = ""
                ConnectedShape.Cells("Prop.SymName").FormulaU = ""
                ConnectedShape.Cells("User.AdrSource").FormulaU = ""
                'Ищем каким концом дочерний приклеен к нам
                For ii = 1 To ConnectedShape.Connects.Count '(возможно это надо убрать под следующий if)
                    If ConnectedShape.Connects(ii).ToSheet = DeletedShape Then
                        SetArrow 254, ConnectedShape.Connects(ii) 'Возвращаем красную стрелку
                    End If
                Next
                'Если это был провод - то + автонумерация дочернего провода
                If ShapeType = typeWire Then
                    ConnectedShape.Cells("Prop.AutoNum").FormulaU = False
                    ConnectedShape.Cells("Prop.HideNumber").FormulaU = False
                    ConnectedShape.Cells("Prop.HideName").FormulaU = True
                    If ConnectedShape.Connects.Count = 2 Then
                        ConnectedShape.Cells("Prop.AutoNum").FormulaU = True
                        'Присваиваем номер проводу
                        AutoNum ConnectedShape
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub ClearWire(vsoShape As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearWire - Чистит при копировании
                'Чистим номер и ссылку при копировании провода.
                'В EventMultiDrop должна быть формула = CALLTHIS("WireNet.ClearWire", "SAPR_ASU")
'------------------------------------------------------------------------------------------------------------
    'If ThisDocument.BlockMacros Then Exit Sub
    'Чистим шейп
    vsoShape.CellsU("Prop.Number").FormulaU = ""
    vsoShape.CellsU("Prop.SymName").FormulaU = ""
    vsoShape.Cells("User.AdrSource").FormulaU = ""
    vsoShape.Cells("Prop.AutoNum").FormulaU = False
    vsoShape.Cells("Prop.HideNumber").FormulaU = False
    vsoShape.Cells("Prop.HideName").FormulaU = True
    'Если подключен 2-мя концами - нумеруем
    If vsoShape.Connects.Count = 2 Then
        vsoShape.Cells("Prop.AutoNum").FormulaU = True
        'Присваиваем номер проводу
        AutoNum vsoShape
    End If

End Sub


Sub SetArrow(Arrow As String, connect As IVConnect)
'------------------------------------------------------------------------------------------------------------
' Macros        : SetArrow - Задает вид окончания провода
'------------------------------------------------------------------------------------------------------------
    If Arrow = "254" Then Arrow = "USE(""RedArrow"")"
    Select Case connect.FromPart
        Case visBegin
            connect.FromSheet.Cells("BeginArrow").Formula = Arrow
        Case visEnd
            connect.FromSheet.Cells("EndArrow").Formula = Arrow
    End Select
End Sub


Sub UnGlue(connect As IVConnect)
'------------------------------------------------------------------------------------------------------------
' Macros        : UnGlue - Отклеивает окончание провода
'------------------------------------------------------------------------------------------------------------
    Select Case connect.FromPart
        Case visBegin
            connect.FromSheet.Cells("BeginX").FormulaU = Chr(34) & connect.FromSheet.Cells("BeginX").Result(0) & Chr(34)
            connect.FromSheet.Cells("BeginY").FormulaU = Chr(34) & connect.FromSheet.Cells("BeginY").Result(0) & Chr(34)
        Case visEnd
            connect.FromSheet.Cells("EndX").FormulaU = Chr(34) & connect.FromSheet.Cells("EndX").Result(0) & Chr(34)
            connect.FromSheet.Cells("EndY").FormulaU = Chr(34) & connect.FromSheet.Cells("EndY").Result(0) & Chr(34)
    End Select
    bUnGlue = True
End Sub


Sub FindZombie(shpProvod As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : FindZombie - Ищем Дочерних которые ссылаются не нас - отцепляем
'------------------------------------------------------------------------------------------------------------
    Dim AdrNashegoProvoda As String
    Dim i As Integer, ii As Integer
    Dim ShapeType As Integer
    
    AdrNashegoProvoda = shpProvod.ContainingPage.NameU & "/" & shpProvod.NameID
    
    'Ищем Дочерних которые ссылаются не нас - отцепляем. Перебор FromConnects.
    For i = 1 To shpProvod.FromConnects.Count
        If i > shpProvod.FromConnects.Count Then Exit For
        Set DeletedConnect = shpProvod.FromConnects(i)
        Set ConnectedShape = DeletedConnect.FromSheet
        
        ShapeType = ConnectedShape.Cells("User.SAType").Result(0)
        
        If (ShapeType = typeWire) Or (ShapeType = typeWireLinkS) Then
            If ConnectedShape.Cells("User.AdrSource").ResultStr(0) <> AdrNashegoProvoda Then 'Дочерний - но ссылается не нас - отцепляем
                'Ищем каким концом дочерний приклеен к нам
                For ii = 1 To ConnectedShape.Connects.Count
                    If ii > ConnectedShape.Connects.Count Then Exit For
                    If ConnectedShape.Connects(ii).ToSheet = shpProvod Then
                        SetArrow 254, ConnectedShape.Connects(ii) 'Возвращаем красную стрелку
                        UnGlue ConnectedShape.Connects(ii) 'Отклеиваем
                    End If
                Next
            End If
        End If
    Next
End Sub

Sub WireToPLCTerm(shpProvod As Visio.Shape, shpPLCTerm As Visio.Shape, bConnect As Boolean)
'------------------------------------------------------------------------------------------------------------
' Macros        : WireToPLCTerm - При подключении провода к клемме входа ПЛК (дочернего)
                'записывает номер провода в родителя PLCIOParent
                'а там, если не 0 то появляется провод с номером подключенного провода,
                'при отключении - возвращаем 0
'------------------------------------------------------------------------------------------------------------
    Dim shpPLCIOParent As Visio.Shape
    Dim LinkWireNumber As String
    Dim PinNumber As Integer
    
    'Ссылка на номер провода
    LinkWireNumber = "Pages[" & shpProvod.ContainingPage.NameU & "]!" & shpProvod.NameID & "!Prop.Number"
    On Error GoTo ExitSub
    'Номер контакта во входе ПЛК
    PinNumber = CInt(Right(shpPLCTerm.Name, 1))
    'Находим родительский вход ПЛК
    Set shpPLCIOParent = HyperLinkToShape(shpPLCTerm.Parent.CellsU("Hyperlink.IO.SubAddress").ResultStr(0))
    'Пишем в него ссылку на номер провода или 0 (когда происходит отсоединение или удаление провода)
    shpPLCIOParent.CellsU("User.w" & PinNumber).FormulaU = IIf(bConnect, LinkWireNumber, 0)
ExitSub:
End Sub