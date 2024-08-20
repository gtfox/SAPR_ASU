Option Explicit


Dim WithEvents vsoPagesEvent As Visio.Pages 'события
Dim WithEvents vsoWindowEvent As Visio.Window 'мышь

Dim vsoShapePaste As Visio.Shape
Public MouseClick As Boolean
Public SelectionMoreOne As Boolean
Public bVpisatVList As Boolean
'Public BlockMacros As Boolean

'Вставка выделенного
Private Sub vsoPagesEvent_SelectionAdded(ByVal Selection As IVSelection)
    Dim vsoShape As Visio.Shape
    Dim colShape As Collection
    Dim bThumbExist As Boolean
    
    If Selection.Count > 1 Then
        SelectionMoreOne = True 'Если в выделении больше 1 элемента - привязку к курсору не делаем
        'Чистим миниатюры контактов после вставки
        Set colShape = New Collection
        For Each vsoShape In Selection
            If ShapeSATypeIs(vsoShape, typeCxemaThumb) Then
                colShape.Add vsoShape
            Else 'Чистим связи, т.к. для эл-тов в Selection дальше первого не происходят события EventDrop/EventMultiDrop
                If vsoShape.CellExists("User.Dropped", 0) Then
                    If vsoShape.Cells("User.Dropped").Result(0) = 1 Then
                        ClearAndAutoNum vsoShape
                    End If
                End If
            End If
        Next
        If colShape.Count > 0 Then
            For Each vsoShape In colShape
                Application.EventsEnabled = False
                vsoShape.Delete
                Application.EventsEnabled = True
            Next
        End If
    End If
End Sub

''Удаление выделенного
'Private Function vsoPagesEvent_QueryCancelSelectionDelete(ByVal Selection As IVSelection) As Boolean
'    Dim vsoShape As Visio.Shape
'    Dim colShape As Collection
'    Dim colThumb As Collection
'
'    If Selection.Count > 1 Then
'        Set colShape = New Collection
'        Set colThumb = New Collection
'        For Each vsoShape In Selection
'            If ShapeSATypeIs(vsoShape, typeCxemaThumb) Then
'                colThumb.Add vsoShape
'            Else
'                colShape.Add vsoShape
'            End If
'        Next
'        'Выделяем всё без миниатюр
'        If colThumb.Count > 0 And colShape.Count > 0 Then
'            ActiveWindow.DeselectAll
'            For Each vsoShape In colShape
'                ActiveWindow.Select vsoShape, visSelect
'            Next
'        End If
'        'Удаляем миниатюры
''        If colThumb.Count > 0 Then
''            For Each vsoShape In colThumb
''                Application.EventsEnabled = False
''                colThumb.Remove vsoShape
''                vsoShape.Delete
''                Application.EventsEnabled = True
''            Next
''        End If
'    End If
''    vsoPagesEvent_QueryCancelSelectionDelete = True
'End Function

'Перед удалением шейпа чистим что-либо
Private Sub vsoPagesEvent_BeforeShapeDelete(ByVal vsoShape As IVShape)
    If vsoShape.CellExists("User.SAType", 0) Then   'Если в шейпе есть тип, то он элемент SAPR ASU
        Select Case ShapeSAType(vsoShape) 'В зависимости от типа выбираем способ удаления
            Case typeCxemaNO, typeCxemaNC 'Контакт реле NO,NC (дочерний)
                DeleteRelayChild vsoShape
            Case typeCxemaCoil, typeCxemaParent, typeCxemaElement, typeCxemaTerm 'Катушка реле KL (родительский)
            'Добавить все остальные, которые соединяются проводами
                DeleteRelayParent vsoShape
            Case typeCxemaWireLinkR  'Разрыв провода (дочерний)
                DeleteWireLinkChild vsoShape
            Case typeCxemaWireLinkS 'Разрыв провода (родительский)
                DeleteWireLinkParent vsoShape
            Case typeCxemaWire   'Провод
                DeleteWire vsoShape
            Case typeCxemaCable   'Кабель на эл. схеме
                DeleteCableSH vsoShape
            Case typePlanSensor, typePlanActuator  'Датчик на ПЛАНЕ
                DeleteSensorChildPlan vsoShape
            Case typeCxemaShkafMesto 'Шкаф/место на эл. схеме
                DeleteShkafMesto vsoShape
            Case typeFSASensor, typeFSAActuator  'Датчик ФСА
                DeleteSensorChild vsoShape
            Case typeFSAPodval 'Подвал на ФСА
                DeleteFSAPodvalChild vsoShape
            Case typeCxemaSensor, typeCxemaActuator   'Датчик/Привод на эл. схеме
                If Not (vsoShape.ContainingPage.NameU Like cListNameSVP & "*") Then
                    DeleteSensorParent vsoShape
                End If
            Case typePLCParent   'ПЛК (родительский)
                DeletePLCParent vsoShape
            Case typePLCChild   'ПЛК (дочерний)
                DeletePLCChild vsoShape
            Case typePLCModParent   'Модуль ПЛК (родительский)
                DeletePLCModParent vsoShape
            Case typePLCModChild   'Модуль ПЛК (дочерний)
                DeletePLCModChild vsoShape
            Case typePLCIOLParent, typePLCIORParent  'Вход ПЛК (родительский)
                DeletePLCIOParent vsoShape
            Case typePLCIOChild   'Вход ПЛК (дочерний)
                DeletePLCIOChild vsoShape
        End Select
    End If
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
    
    ClearAndAutoNum vsoShapeEvent
    
    If vsoShapeEvent.Cells("User.Dropped").Result(0) = 0 Then 'Вбросили из набора элементов
        vsoShapeEvent.Cells("User.Dropped").FormulaU = 1
    ElseIf Not SelectionMoreOne And ActiveWindow.Selection.Count = 1 Then 'Если в выделении больше 1 элемента - привязку к курсору не делаем
        Set vsoWindowEvent = ActiveWindow 'Вбросили при копировании - привязываем к курсору
    Else    'Запрет привязки делаем только 1 раз после SelectionAdded
        SelectionMoreOne = False 'Разрешаем привязку
    End If

    MouseClick = False 'Начинаем ждать клика
End Sub

Sub ClearAndAutoNum(vsoShapeEvent As Visio.Shape)

    'AutoNum без мыши но с очисткой - для применения в MultiDrop

    Select Case ShapeSAType(vsoShapeEvent)
    
        Case typeCxemaNO, typeCxemaNC 'Контакты
        
            ClearRelayChild vsoShapeEvent 'Чистим ссылки в дочернем при его копировании.
            
        Case typeCxemaWireLinkS, typeCxemaWireLinkR 'Разрывы проводов
            
            If vsoShapeEvent.Cells("User.Dropped").Result(0) = 1 Then 'Если не вбросили из набора элементов
                ClearReferenceWireLink vsoShapeEvent 'Чистим ссылки в при копировании разрыва провода.
            End If
            
        Case typeCxemaWire 'Провода

            If vsoShapeEvent.Cells("User.Dropped").Result(0) = 1 Then 'Если не вбросили из набора элементов
                'Не нумеруем, т.к. нумеруется в процессе соединения
                ClearWire vsoShapeEvent
            End If
        
        Case typeCxemaCable 'Кабели на схеме электрической
        
            'Чистим ссылку на план
            ClearCableSH vsoShapeEvent 'Чистим ссылку
            AutoNum vsoShapeEvent 'Автонумерация
            
        Case typeSVPCable, typePlanCable, typePlanDuct, typeVidShkafaDIN, typeVidShkafaDver, typeVidShkafaShkaf, typePlanBox
        
            'Не нумеруем при вбросе
        
        Case typeFSASensor, typeFSAActuator 'Датчик на ФСА
        
            'Отвязываем и нумеруем
            ClearSensorChild vsoShapeEvent 'Чистим ссылки
            AutoNumFSA vsoShapeEvent 'Автонумерация
            
        Case typeFSAPodval 'Канал в подвале ФСА
            
            'Отвязываем и нумеруем
            ClearFSAPodvalChild vsoShapeEvent 'Чистим ссылки
            AutoNumFSA vsoShapeEvent 'Автонумерация
        
        Case typeCxemaSensor, typeCxemaActuator 'датчики, двигатели, приводы вне шкафа
            
            'Отвязываем и нумеруем
            ClearSensorParent vsoShapeEvent 'Чистим ссылки
            AutoNum vsoShapeEvent 'Автонумерация
            
        Case typePLCChild 'ПЛК дочерний
        
            'Отвязываем
            ClearPLCChild vsoShapeEvent 'Чистим ссылки
        
        Case typePLCParent 'ПЛК родительский
            
            'Отвязываем и нумеруем
            ClearPLCParent vsoShapeEvent 'Чистим ссылки
            AutoNum vsoShapeEvent 'Автонумерация
            
        Case Else 'Катушки реле, кнопки, переключатели, контакоры, лампочки,  ...
            
            ClearRelayParent vsoShapeEvent 'Чистим старые ссылки в Scratch
            AutoNum vsoShapeEvent 'Автонумерация
            
    End Select
End Sub

'Соединение шейпов
Private Sub vsoPagesEvent_ConnectionsAdded(ByVal Connects As IVConnects)
    If Connects.FromSheet.CellExistsU("User.SAType", 0) Then 'То что цепляем - объект SAPR_ASU
       If Not Connects.ToSheet Is Nothing Then
           If Connects.ToSheet.CellExistsU("User.SAType", 0) Then 'То к чему цепляем - объект SAPR_ASU
                Select Case ShapeSAType(Connects.FromSheet) 'То что цепляем - это...
                    Case typeCxemaWire   'Цепляем провод
                        ConnectWire Connects
                    Case typePlanVynoska, typePlanVynoska2 'Цепляем выноску
                        VynoskaPlan Connects
    '                Case typePlanPodem 'Цепляем подъём на отметку
    '                    PodemPlan Connects
                End Select
            End If
        End If
    End If
End Sub

'Отсоединение шейпов
Private Sub vsoPagesEvent_ConnectionsDeleted(ByVal Connects As IVConnects)
    If Connects.FromSheet.CellExistsU("User.SAType", 0) Then 'То что отцепляем - объект SAPR_ASU
        If Not Connects.ToSheet Is Nothing Then
           If Connects.ToSheet.CellExistsU("User.SAType", 0) Then 'То от чего отцепляем - объект SAPR_ASU
                Select Case ShapeSAType(Connects.FromSheet) 'То что отцепляем - это...
                    Case typeCxemaWire   'Отцепляем провод
                        DisconnectWire Connects
                    Case typePlanVynoska, typePlanVynoska2 'Отцепляем выноску
                        VynoskaPlan Connects
    '                Case typePlanPodem 'Отцепляем подъём на отметку
    '                    PodemPlan Connects
                End Select
            End If
        End If
    End If
End Sub

'Активация событий при старте
Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    Set vsoPagesEvent = ActiveDocument.Pages
    AddToolBar
End Sub

'Масштабирование на листе ВИД
Private Sub vsoPagesEvent_ShapeAdded(ByVal Shape As IVShape)
    If bVpisatVList Then
        frmMenuVpisatVList.run Shape
        bVpisatVList = False
    End If
End Sub

'Таскаем фируру за мышкой
Private Sub vsoWindowEvent_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal X As Double, ByVal Y As Double, CancelDefault As Boolean)
    Dim blsPrevious As Boolean
    
    If Not MouseClick Then
        If vsoShapePaste.OneD <> -1 Then
'            blsPrevious = Application.UndoEnabled
'            Application.UndoEnabled = False
            On Error Resume Next
            vsoShapePaste.Cells("PinX") = X
            vsoShapePaste.Cells("PinY") = Y
'            Application.UndoEnabled = blsPrevious
        End If
    End If
End Sub

'Закончили таскать фигуру кликом мыши
Private Sub vsoWindowEvent_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal X As Double, ByVal Y As Double, CancelDefault As Boolean)
    MouseClick = True
    Set vsoWindowEvent = Nothing
'    Application.DoCmd visCmdSelectionModeRect 'Возврат мыши
End Sub

'Чистим события
Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    Set vsoPagesEvent = Nothing
    Set vsoWindowEvent = Nothing
End Sub

'Активация событий по кнопке в меню/на пенели
Public Sub InitEvent()
    Set vsoPagesEvent = ActiveDocument.Pages
End Sub




