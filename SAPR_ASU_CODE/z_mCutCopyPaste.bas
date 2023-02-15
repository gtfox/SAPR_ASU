'---------------------------------------------------------------------------------------
' Module : mCutCopyPaste
'---------------------------------------------------------------------------------------
' Author : Щербаков Дмитрий(The_Prist)
'          Профессиональная разработка приложений для MS Office любой сложности
'          Проведение тренингов по MS Excel
'          https://www.excel-vba.ru
'          info@excel-vba.ru
'          WebMoney - R298726502453; Яндекс.Деньги - 41001332272872
' Purpose: https://www.excel-vba.ru/chto-umeet-excel/sobstvennoe-menyu-v-textbox
'---------------------------------------------------------------------------------------
'        Применение:
'
'        Private Sub UserForm_Initialize()
'            InitCustomCCPMenu Me 'Контекстное меню для TextBox
'        End Sub
'
'        Private Sub UserForm_Terminate()
'            DelCustomCCPMenu 'Удаления контекстного меню для TextBox
'        End Sub


Option Explicit
Public Const sPopupMenuName$ = "CustomCCP_PopupMenu"    'имя собственного меню
Public tbxAct As MSForms.TextBox                        'переменная уровня проекта для запоминания текущего TextBox
                                                        'того, из которого вызвали собственное меню
Public colTbx() As New clsCustomMenu                    'массив хранения TextBox-ов для обработки через модуль класса

Public Sub InitCustomCCPMenu(UserForm As MSForms.UserForm)
'при вызове формы
'   создаем свое меню "Вырезать-Копировать-Вставить" для вызова из TextBox
'   и для всех TextBox-ов назначаем обработку событий через модуль класса clsCustomMenu
'   Подробнее про модули классов: https://www.excel-vba.ru/chto-umeet-excel/rabota-s-modulyami-klassov/

    Call DelCustomCCPMenu 'на всякий случай удаляем меню, если вдруг оно уже есть
    With Application.CommandBars.Add(sPopupMenuName, msoBarPopup)
        With .Controls.Add(msoControlButton)
            .Caption = "Копировать"                 'текст кнопки
            .FaceId = "19"                          'код иконки для кнопки
            .OnAction = "MyPopupMenuButtonCClick"   'имя процедуры, которая будет выполнена при нажатии кнопки
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "Вставить"                   'текст кнопки
            .FaceId = "22"                          'код иконки для кнопки
            .OnAction = "MyPopupMenuButtonPClick"   'имя процедуры, которая будет выполнена при нажатии кнопки
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "Вырезать"                   'текст кнопки
            .FaceId = "21"                          'код иконки для кнопки
            .OnAction = "MyPopupMenuButtonCutClick" 'имя процедуры, которая будет выполнена при нажатии кнопки
        End With
    End With
    'для каждого TextBox назначаем отслеживание нажатия для вызова меню
    Dim oCt As Object
    Dim li As Long
    For Each oCt In UserForm.Controls
        If TypeOf oCt Is MSForms.TextBox Then
            'расширяем массив текстбоксов еще на один элемент
            ReDim Preserve colTbx(li)
            'записываем текстбокс в массив
            Set colTbx(li).oTbx = oCt
            li = li + 1
        End If
    Next
End Sub

'функция удаления созданного ранее меню
Public Sub DelCustomCCPMenu()
    On Error Resume Next 'пропускаем ошибки, если вдруг меню было удалено ранее
    Application.CommandBars(sPopupMenuName).Delete
End Sub
Public Sub MyPopupMenuButtonCClick()
    tbxAct.Copy
    Set tbxAct = Nothing
End Sub
Public Sub MyPopupMenuButtonPClick()
    tbxAct.Paste
    Set tbxAct = Nothing
End Sub
Public Sub MyPopupMenuButtonCutClick()
    tbxAct.Cut
    Set tbxAct = Nothing
End Sub
