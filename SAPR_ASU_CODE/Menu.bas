Sub AddToolBar()
    Dim Bar As CommandBar
    
    'Меню существует?
    For Each Bar In Application.CommandBars
        If Bar.Name = "САПР АСУ" Then Exit Sub
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
    
    '---Кнопка Блокировки рамки
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "БлокРамки"
        .Tag = "LockTitle"
        .OnAction = "LockTitleBlock"
        .TooltipText = "Блокировка рамки"
        .FaceID = 894 '519
    End With
    
    '---Кнопка Формат->Специальный
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=33841, Before:=2)
    With Button
        .Caption = "ФорматСпециальный"
        .Tag = "FormatSpecial"
        .style = msoButtonAutomatic
        '.OnAction = "LockTitleBlock"
        .TooltipText = "Формат->Специальный"
        .FaceID = 274
    End With
    
        '---Кнопка ObjInfo Формат->Специальный +
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=3)
    With Button
        .Caption = "ФорматСпециальныйNameU"
        .Tag = "ObjInfo"
        .style = msoButtonAutomatic
        .OnAction = "ObjInfo"
        .TooltipText = "Формат->Специальный+NameU"
        .FaceID = 487
    End With
    
        '---Кнопка Экспорта на GitHub
    Set Button = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Before:=4)
    With Button
        .Caption = "ЭкспортGitHub"
        .Tag = "ExportGit"
        .style = msoButtonAutomatic
        .OnAction = "ExportGitHub"
        .TooltipText = "Экспорт кода для GitHub"
        .FaceID = 3
    End With
    
    
    Set Button = Nothing
           
End Sub