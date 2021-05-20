Option Explicit


Dim WithEvents vsoPagesEvent As Visio.Pages 'события
Dim WithEvents vsoWindowEvent As Visio.Window 'мышь

Dim vsoShapePaste As Visio.Shape
Public MouseClick As Boolean
Public SelectionMoreOne As Boolean
'Public BlockMacros As Boolean



''Перед удалением кучи шейпов сначала выкидываем миниатюры из выделения, иначе крашится, т.к.
''в удалении шейпа сидит удаление миниатюры, что вызывает повторное срабатывание BeforeShapeDelete,
''но уже с другим объектом, а предыдущее не завершилось...
''Или повторное вызывается для уже удаленного объекта...

''НЕ ПОМОГЛО. Какаято хуйня творится во время удаления.

'Private Sub vsoPagesEvent_BeforeSelectionDelete(ByVal Selection As IVSelection)
'    Dim vsoShape As Visio.Shape
'
'    For Each vsoShape In Selection
'        If ShapeSATypeis(vsoShape, typeThumb) Then
'            Selection.Select vsoShape, visDeselect
'        End If
'    Next
'End Sub



'Перед удалением шейпа чистим что-либо
Private Sub vsoPagesEvent_BeforeShapeDelete(ByVal vsoShape As IVShape)
    If vsoShape.CellExists("User.SAType", 0) Then   'Если в шейпе есть тип, то он элемент SAPR ASU
        Select Case ShapeSAType(vsoShape) 'В зависимости от типа выбираем способ удаления
            Case typeNO, typeNC 'Контакт реле NO,NC (дочерний)
                DeleteRelayChild vsoShape
            Case typeCoil, typeParent, typeElement, typeTerm 'Катушка реле KL (родительский)
            'Добавить все остальные, которые соединяются проводами
                DeleteRelayParent vsoShape
            Case typeWireLinkR  'Разрыв провода (дочерний)
                DeleteWireLinkChild vsoShape
            Case typeWireLinkS 'Разрыв провода (родительский)
                DeleteWireLinkParent vsoShape
            Case typeWire   'Провод
                DeleteWire vsoShape
            Case typeCableSH   'Кабель на эл. схеме
                DeleteCableSH vsoShape
            Case typeFSASensor   'Датчик ФСА
                If Not (vsoShape.ContainingPage.NameU Like cListNamePlan & "*") Then
                    DeleteSensorChild vsoShape
                End If
            Case typeFSAPodval 'Подвал на ФСА
                DeleteFSAPodvalChild vsoShape
            Case typeSensor, typeActuator   'Датчик/Привод на эл. схеме
                DeleteSensorParent vsoShape
            Case typePLCParent   'ПЛК (родительский)
                DeletePLCParent vsoShape
            Case typePLCChild   'ПЛК (дочерний)
                DeletePLCChild vsoShape
            Case typePLCModParent   'Модуль ПЛК (родительский)
                DeletePLCModParent vsoShape
            Case typePLCModChild   'Модуль ПЛК (дочерний)
                DeletePLCModChild vsoShape
            Case typePLCIOParent   'Вход ПЛК (родительский)
                DeletePLCIOParent vsoShape
            Case typePLCIOChild   'Вход ПЛК (дочерний)
                DeletePLCIOChild vsoShape
        End Select
    End If

End Sub

'Активация событий при старте
Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    Set vsoPagesEvent = ActiveDocument.Pages
    AddToolBar
End Sub

'Соединение шейпов
Private Sub vsoPagesEvent_ConnectionsAdded(ByVal Connects As IVConnects)
    If Connects.FromSheet.CellExistsU("User.SAType", 0) Then 'То что цепляем - объект SAPR_ASU
       If Connects.ToSheet.CellExistsU("User.SAType", 0) Then 'То к чему цепляем - объект SAPR_ASU
            Select Case ShapeSAType(Connects.FromSheet) 'То что цепляем - это...
                Case typeWire   'Цепляем провод
                    ConnectWire Connects
                Case typeVynoskaPL 'Цепляем выноску
                    CableInfoPlan Connects
            End Select
        End If
    End If
End Sub

'Отсоединение шейпов
Private Sub vsoPagesEvent_ConnectionsDeleted(ByVal Connects As IVConnects)
    If Connects.FromSheet.CellExistsU("User.SAType", 0) Then 'То что отцепляем - объект SAPR_ASU
       If Connects.ToSheet.CellExistsU("User.SAType", 0) Then 'То от чего отцепляем - объект SAPR_ASU
            Select Case ShapeSAType(Connects.FromSheet) 'То что отцепляем - это...
                Case typeWire   'Отцепляем провод
                    DisconnectWire Connects
                Case typeVynoskaPL 'Отцепляем выноску
                    CableInfoPlan Connects
            End Select
        End If
    End If
End Sub

'Если в выделении больше 1 элемента - привязку к курсору не делаем
Private Sub vsoPagesEvent_SelectionAdded(ByVal Selection As IVSelection)
    If Selection.Count > 1 Then
        SelectionMoreOne = True
    End If
End Sub

'Таскаем фируру за мышкой
Private Sub vsoWindowEvent_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal Y As Double, CancelDefault As Boolean)
    If Not MouseClick Then
        On Error Resume Next
        vsoShapePaste.Cells("PinX") = x
        vsoShapePaste.Cells("PinY") = Y
    End If
End Sub

'Закончили таскать фигуру кликом мыши
Private Sub vsoWindowEvent_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal Y As Double, CancelDefault As Boolean)
    MouseClick = True
    Set vsoWindowEvent = Nothing
End Sub

Sub EventDropAutoNum(vsoShapeEvent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : EventDropAutoNum - Автонумерация для одиночной вставки
                'Когда происходит вставка применяется привязка к курсору
                'Если вставка была из набора элементов - привязка к курсору не происходит
                '(Для контроля вброса из набора применяется переменная Dropped в каждой фигуре)
                'В EventDrop должна быть формула =CALLTHIS("ThisDocument.EventDropAutoNum")
'------------------------------------------------------------------------------------------------------------
    
'    InitEvent 'Активация событий

    Set vsoShapePaste = vsoShapeEvent
    
    If vsoShapeEvent.Cells("User.Dropped").Result(0) = 0 Then 'Вбросили из набора элементов
        vsoShapeEvent.Cells("User.Dropped").FormulaU = 1
    ElseIf Not SelectionMoreOne Then 'Если в выделении больше 1 элемента - привязку к курсору не делаем
        Set vsoWindowEvent = ActiveWindow 'Вбросили при копировании - привязываем к курсору
    Else    'Запрет привязки делаем только 1 раз после SelectionAdded
        SelectionMoreOne = False 'Разрешаем привязку
    End If
    
    ClearAndAutoNum vsoShapeEvent
    
    MouseClick = False 'Начинаем ждать клика
End Sub

Sub ClearAndAutoNum(vsoShapeEvent As Visio.Shape)

    'AutoNum без мыши но с очисткой - для применения в MultiDrop

    Select Case ShapeSAType(vsoShapeEvent)
    
        Case typeNO, typeNC 'Контакты
        
            ClearRelayChild vsoShapePaste 'Чистим ссылки в дочернем при его копировании.
            
        Case typeWireLinkS, typeWireLinkR 'Разрывы проводов
        
            ClearReferenceWireLink vsoShapePaste 'Чистим ссылки в при копировании разрыва провода.
            
        Case typeWire 'Провода
        
            'Не нумеруем, т.к. нумеруется в процессе соединения
            ClearWire vsoShapePaste
            
        Case typeCableVP, typeCablePL, typeDuctPlan, typeVidShkafaDIN, typeVidShkafaDver, typeVidShkafaShkaf, typeBox
        
            'Не нумеруем при вбросе
        
        Case typeFSASensor 'Датчик на ФСА
        
            'Отвязываем и нумеруем
            ClearSensorChild vsoShapePaste 'Чистим ссылки
            AutoNumFSA vsoShapePaste 'Автонумерация
            
        Case typeFSAPodval 'Канал в подвале ФСА
            
            'Отвязываем и нумеруем
            ClearFSAPodvalChild vsoShapePaste 'Чистим ссылки
            AutoNumFSA vsoShapePaste 'Автонумерация
        
        Case typeSensor, typeActuator 'датчики, двигатели, приводы вне шкафа
            
            'Отвязываем и нумеруем
            ClearSensorParent vsoShapePaste 'Чистим ссылки
            AutoNum vsoShapePaste 'Автонумерация
            
        Case typePLCChild 'ПЛК дочерний
        
            'Отвязываем
            ClearPLCChild vsoShapePaste 'Чистим ссылки
        
        Case typePLCParent 'ПЛК родительский
            
            'Отвязываем и нумеруем
            ClearPLCParent vsoShapePaste 'Чистим ссылки
            AutoNum vsoShapePaste 'Автонумерация
            
        Case Else 'Катушки реле, кнопки, переключатели, контакоры, лампочки,  ...
            
            ClearRelayParent vsoShapePaste 'Чистим старые ссылки в Scratch
            AutoNum vsoShapePaste 'Автонумерация
            
    End Select
End Sub

'Чистим события
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    Set vsoPagesEvent = Nothing
    Set vsoWindowEvent = Nothing
End Sub

'Активация событий по кнопке в меню/на пенели
Sub InitEvent()
    Set vsoPagesEvent = ActiveDocument.Pages
End Sub




