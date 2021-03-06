Sub AddToolBar()
    Dim Bar As CommandBar
    
    'Меню существует?
    For Each Bar In Application.CommandBars
        If Bar.Name = "САПР АСУ" Then Bar.Delete 'Exit Sub
    Next
    
    Set Bar = Application.CommandBars.Add(Position:=msoBarTop, Temporary:=True) 'msoBarTop msoBarFloating
    
    With Bar
        .Name = "САПР АСУ"
        .Visible = True
        .RowIndex = 7
        .Left = 944
        .Top = 104
    End With
    
    AddButtons

End Sub


Private Sub AddButtons()

    Dim Bar As CommandBar
    Dim Button As CommandBarButton

    Set Bar = Application.CommandBars("САПР АСУ")
    

    
    '---Кнопка Формат->Специальный
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=33841)
    With Button
        .Caption = "ФорматСпециальный"
        .Tag = "FormatSpecial"
        .style = msoButtonAutomatic
        '.OnAction = "LockTitleBlock"
        .TooltipText = "Формат->Специальный"
        .FaceID = 274
    End With
    
        '---Кнопка ObjInfo Формат->Специальный +
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=2)
    With Button
        .Caption = "ФорматСпециальныйNameU"
        .Tag = "ObjInfo"
        .style = msoButtonAutomatic
        .OnAction = "ObjInfo"
        .TooltipText = "Формат->Специальный+NameU"
        .FaceID = 487
    End With
    
        '---Кнопка Экспорта на GitHub
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=3)
    With Button
        .Caption = "ЭкспортGitHub"
        .Tag = "ExportGit"
        .style = msoButtonAutomatic
        .OnAction = "ExportGitHub"
        .TooltipText = "Экспорт кода для GitHub"
        .FaceID = 521 '3
    End With
    
        '---Кнопка Блокировки рамки
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=4)
    With Button
        .Caption = "БлокРамки"
        .Tag = "LockTitle"
        .OnAction = "LockTitleBlock"
        .TooltipText = "Блокировка рамки"
        .FaceID = 894 '519
        .BeginGroup = True
    End With
    
            '---Кнопка Сохранить копию проекта
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=5)
    With Button
        .Caption = "СохранитьПроект"
        .Tag = "SaveFileAs"
        .OnAction = "SaveProjectFileAs"
        .TooltipText = "Сохранить копию проекта"
        .FaceID = 3
'        .BeginGroup = True
    End With

        '---Кнопка Добавить лист
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=6)
    With Button
        .Caption = "ДобавитьЛист"
        .Tag = "AddPage"
        .style = msoButtonAutomatic
        .OnAction = "AddSAPageNext"
        .TooltipText = "Добавить лист"
        .FaceID = 535 '18
        .BeginGroup = True
    End With
    
        '---Кнопка Удалить лист
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=7)
    With Button
        .Caption = "УдалитьЛист"
        .Tag = "DelPage"
        .style = msoButtonAutomatic
        .OnAction = "DelSAPage"
        .TooltipText = "Удалить лист"
        .FaceID = 536 '305
    End With
    
        '---Кнопка Создать раздел
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=8)
    With Button
        .Caption = "СоздатьРаздел"
        .Tag = "AddRazdel"
        .style = msoButtonAutomatic
        .OnAction = "ShowSAPageRazdel"
        .TooltipText = "Создать раздел"
        .FaceID = 533 '786
    End With

        '---Кнопка Копировать лист
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=9)
    With Button
        .Caption = "КопироватьЛист"
        .Tag = "CopyList"
        .style = msoButtonAutomatic
        .OnAction = "CopySAPage"
        .TooltipText = "Копировать лист"
        .FaceID = 531 '585
    End With

        '---Кнопка Перенумерация
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=10)
    With Button
        .Caption = "ПеренумерацияЭлементов"
        .Tag = "ReNumber"
        .style = msoButtonAutomatic
        .OnAction = "ShowReNumber"
        .TooltipText = "Перенумерация элементов"
        .FaceID = 2476 '786
        .BeginGroup = True
    End With
    
        '---Кнопка Настройки
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=11)
    With Button
        .Caption = "НастройкиПроекта"
        .Tag = "SettingsProject"
        .style = msoButtonAutomatic
        .OnAction = "ShowSettingsProject"
        .TooltipText = "Настройки Проекта"
        .FaceID = 642
        .BeginGroup = True
    End With
    
    Set Button = Nothing
           
End Sub