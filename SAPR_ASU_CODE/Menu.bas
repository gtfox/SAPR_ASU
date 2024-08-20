Sub AddToolBar()
    Dim Bar As CommandBar
    
    'Меню САПР АСУ
    For Each Bar In Application.CommandBars
        If Bar.name = "САПР АСУ" Then Bar.Delete 'Exit Sub
    Next
    
    Set Bar = Application.CommandBars.Add(Position:=msoBarTop, Temporary:=True) 'msoBarTop msoBarFloating
    
    With Bar
        .name = "САПР АСУ"
        .Visible = True
        .RowIndex = 7
        .Left = 944
        .Top = 104
    End With
    
    AddButtons

    'Меню САПР АСУ ВИД
    For Each Bar In Application.CommandBars
        If Bar.name = "САПР АСУ ВИД" Then Bar.Delete 'Exit Sub
    Next
    
    Set Bar = Application.CommandBars.Add(Position:=msoBarTop, Temporary:=True) 'msoBarTop msoBarFloating
    
    With Bar
        .name = "САПР АСУ ВИД"
        .Visible = True
        .RowIndex = 7
        .Left = 944
        .Top = 104
    End With
    
    AddButtonsVID

    'Меню САПР АСУ СХЕМА
    For Each Bar In Application.CommandBars
        If Bar.name = "САПР АСУ СХЕМА" Then Bar.Delete 'Exit Sub
    Next
    
    Set Bar = Application.CommandBars.Add(Position:=msoBarTop, Temporary:=True) 'msoBarTop msoBarFloating
    
    With Bar
        .name = "САПР АСУ СХЕМА"
        .Visible = True
        .RowIndex = 7
        .Left = 944
        .Top = 104
    End With
    
    AddButtonsCXEMA

End Sub


Private Sub AddButtons()

    Dim Bar As CommandBar
    Dim Button As CommandBarButton

    Set Bar = Application.CommandBars("САПР АСУ")
    

    
'    '---Кнопка Формат->Специальный
'    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1) '33841
'    With Button
'        .Caption = "ФорматСпециальный"
'        .Tag = "FormatSpecial"
'        .style = msoButtonAutomatic
'        '.OnAction = "LockTitleBlock"
'        .TooltipText = "Формат->Специальный"
'        .FaceID = 274
'    End With
    
        '---Кнопка ObjInfo Формат->Специальный +
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1)
    With Button
        .Caption = "ФорматСпециальныйNameU"
        .Tag = "ObjInfo"
        .style = msoButtonAutomatic
        .OnAction = "ObjInfo"
        .TooltipText = "Формат->Специальный+NameU"
        .FaceId = 487
    End With
    
        '---Кнопка Экспорта на GitHub
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=2)
    With Button
        .Caption = "ЭкспортGitHub"
        .Tag = "ExportGit"
        .style = msoButtonAutomatic
        .OnAction = "ExportGitHub"
        .TooltipText = "Экспорт кода для GitHub"
        .FaceId = 521 '3
    End With
    
        '---Кнопка Сохранить копию проекта
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=3)
    With Button
        .Caption = "СохранитьПроект"
        .Tag = "SaveFileAs"
        .OnAction = "SaveProjectFileAs"
        .TooltipText = "Сохранить копию проекта"
        .FaceId = 3
'        .BeginGroup = True
    End With
    
        '---Кнопка Блокировки рамки
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=4)
    With Button
        .Caption = "БлокРамки"
        .Tag = "LockTitle"
        .OnAction = "LockTitleBlock"
        .TooltipText = "Блокировка рамки"
        .FaceId = 894 '519
        .BeginGroup = True
    End With

        '---Кнопка Добавить лист
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=5)
    With Button
        .Caption = "ДобавитьЛист"
        .Tag = "AddPage"
        .style = msoButtonAutomatic
        .OnAction = "AddSAPageNext"
        .TooltipText = "Добавить лист"
        .FaceId = 535 '18
        .BeginGroup = True
    End With
    
        '---Кнопка Удалить лист
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=6)
    With Button
        .Caption = "УдалитьЛист"
        .Tag = "DelPage"
        .style = msoButtonAutomatic
        .OnAction = "DelSAPage"
        .TooltipText = "Удалить лист"
        .FaceId = 536 '305
    End With
    
        '---Кнопка Создать раздел
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=7)
    With Button
        .Caption = "СоздатьРаздел"
        .Tag = "AddRazdel"
        .style = msoButtonAutomatic
        .OnAction = "ShowSAPageRazdel"
        .TooltipText = "Создать раздел"
        .FaceId = 533 '786
    End With

        '---Кнопка Копировать лист
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=8)
    With Button
        .Caption = "КопироватьЛист"
        .Tag = "CopyList"
        .style = msoButtonAutomatic
        .OnAction = "CopySAPage"
        .TooltipText = "Копировать лист"
        .FaceId = 531 '585
    End With

        '---Кнопка Перенумерация
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=9)
    With Button
        .Caption = "ПеренумерацияЭлементов"
        .Tag = "ReNumber"
        .style = msoButtonAutomatic
        .OnAction = "ShowReNumber"
        .TooltipText = "Перенумерация элементов"
        .FaceId = 2476 '786
        .BeginGroup = True
    End With
    
        '---Кнопка Данные для спецификации
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=10)
    With Button
        .Caption = "ДанныеСпецификации"
        .Tag = "Specifikaciya"
        .style = msoButtonAutomatic
        .OnAction = "ShowSpecifikaciya"
        .TooltipText = "Перечень оборудования из Visio в Excel"
        .FaceId = 263 '5897
        .BeginGroup = True
    End With

        '---Кнопка Сохранить в PDF
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=11)
    With Button
        .Caption = "СохранитьвPDF"
        .Tag = "SavePDF"
        .style = msoButtonAutomatic
        .OnAction = "SavePDF"
        .TooltipText = "Сохранить в PDF"
        .FaceId = 267
        .BeginGroup = True
    End With
    
        '---Кнопка Сохранить в PDF Цветное
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=12)
    With Button
        .Caption = "СохранитьвPDFЦветное"
        .Tag = "SavePDFColor"
        .style = msoButtonAutomatic
        .OnAction = "SavePDFColor"
        .TooltipText = "Сохранить в PDF в цвете"
        .FaceId = 508
        .BeginGroup = True
    End With
    
        '---Кнопка Настройки
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=13)
    With Button
        .Caption = "НастройкиПроекта"
        .Tag = "SettingsProject"
        .style = msoButtonAutomatic
        .OnAction = "ShowSettingsProject"
        .TooltipText = "Настройки Проекта"
        .FaceId = 642
        .BeginGroup = True
    End With

        '---Кнопка 0
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=14)
    With Button
        .Caption = "0"
        .Tag = "SettingsProject"
        .style = msoButtonAutomatic
        .OnAction = "SetUserSAType_0"
        .TooltipText = "0"
        .FaceId = 70
        .BeginGroup = True
    End With

        '---Кнопка 132
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=15)
    With Button
        .Caption = "132"
        .Tag = "SettingsProject"
        .style = msoButtonAutomatic
        .OnAction = "SetUserSAType_132"
        .TooltipText = "132"
        .FaceId = 59
        .BeginGroup = True
    End With



    Set Button = Nothing
           
End Sub

Private Sub AddButtonsVID()

    Dim Bar As CommandBar
    Dim Button As CommandBarButton

    Set Bar = Application.CommandBars("САПР АСУ ВИД")

        '---Кнопка "Вписать в лист"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1)
    With Button
        .Caption = "Вписатьвлист"
        .Tag = "VpisatVList"
        .style = msoButtonAutomatic
        .OnAction = "VpisatVList"
        .TooltipText = "Вписать в лист"
        .FaceId = 25 '1796
    End With
    
        '---Кнопка "Распределить на двери"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=2)
    With Button
        .Caption = "Распределитьнадвери"
        .Tag = "RaspredelitGorizont"
        .style = msoButtonAutomatic
        .OnAction = "RaspredelitGorizont"
        .TooltipText = "Распределить на двери"
        .FaceId = 1650 '408 '669 '2067
    End With
    
        '---Кнопка "Вертикальные размеры"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=3)
    With Button
        .Caption = "Вертикальныеразмеры"
        .Tag = "VertRazmery"
        .OnAction = "VertRazmery"
        .TooltipText = "Вертикальные размеры"
        .FaceId = 1647 '1137=1258 2068
'        .BeginGroup = True
    End With

    Set Button = Nothing
           
End Sub

Private Sub AddButtonsCXEMA()

    Dim Bar As CommandBar
    Dim Button As CommandBarButton

    Set Bar = Application.CommandBars("САПР АСУ СХЕМА")

        '---Кнопка "Дубликат 2х"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1)
    With Button
        .Caption = "Дубликат2х"
        .Tag = "Duplicate"
        .style = msoButtonAutomatic
        .OnAction = "Duplicate"
        .TooltipText = "Дубликат 2х"
        .FaceId = 72 '1836 '19 '1774 1807 1950 523
    End With
    
        '---Кнопка "Сначала группа"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=2)
    With Button
        .Caption = "Сначалагруппа"
        .Tag = "BeginGroup"
        .style = msoButtonAutomatic
        .OnAction = "BeginGroup"
        .TooltipText = "Сначала группа"
        .FaceId = 623 '2761
        .BeginGroup = True
    End With

        '---Кнопка "Только группа"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=3)
    With Button
        .Caption = "Толькогруппа"
        .Tag = "OnlyGroup"
        .style = msoButtonAutomatic
        .OnAction = "OnlyGroup"
        .TooltipText = "Только группа"
        .FaceId = 572
'        .BeginGroup = True
    End With
    
        '---Кнопка "Показать дочерние номера проводов"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=4)
    With Button
        .Caption = "Показатьдочерниеномерапроводов"
        .Tag = "ShowWireNumChildInDoc"
        .OnAction = "ShowWireNumChildInDoc"
        .TooltipText = "Показать дочерние номера проводов"
        .FaceId = 291 '2810 2805
'        .BeginGroup = True
    End With
    
        '---Кнопка "Скрыть дочерние номера проводов"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=5)
    With Button
        .Caption = "Скрытьдочерниеномерапроводов"
        .Tag = "HideWireNumChildInDoc"
        .OnAction = "HideWireNumChildInDoc"
        .TooltipText = "Скрыть дочерние номера проводов"
        .FaceId = 290 '2810 2805
        .BeginGroup = True
    End With

        '---Кнопка "Вставить миниатюры контактов"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=6)
    With Button
        .Caption = "Вставитьминиатюрыконтактов"
        .Tag = "AddLocThumbAllInDoc"
        .OnAction = "AddLocThumbAllInDoc"
        .TooltipText = "Вставить миниатюры контактов"
        .FaceId = 2871
        .BeginGroup = True
    End With

        '---Кнопка "Удалить миниатюры контактов"
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=7)
    With Button
        .Caption = "Удалитьминиатюрыконтактов"
        .Tag = "DelLocThumbAllInDoc"
        .OnAction = "DelLocThumbAllInDoc"
        .TooltipText = "Удалить миниатюры контактов"
        .FaceId = 2164
'        .BeginGroup = True
    End With

            '---Кнопка Создать шаблон схемы
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=8)
    With Button
        .Caption = "Создатьшаблонсхемы"
        .Tag = "LockSelect"
        .style = msoButtonAutomatic
        .OnAction = "MenuAddToStencilFrm"
        .TooltipText = "Создать шаблон схемы"
        .FaceId = 516 '2135 '2134 '1807 '1672 '1048 516 582
        .BeginGroup = True
    End With

    Set Button = Nothing
    
        '---Кнопка Блокировки выделенного объекта
    Set Button = Bar.Controls.Add(Type:=msoControlButton, id:=1, Before:=9)
    With Button
        .Caption = "БлокировкаВыделенного"
        .Tag = "LockSelect"
        .style = msoButtonAutomatic
        .OnAction = "LockSelected"
        .TooltipText = "Блокировка выделенных объектов"
        .FaceId = 519
        .BeginGroup = True
    End With

    Set Button = Nothing
           
End Sub
