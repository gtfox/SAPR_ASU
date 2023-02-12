'---------------------------------------------------------------------------------------
' Author : Щербаков Дмитрий(The_Prist)
'          Профессиональная разработка приложений для MS Office любой сложности
'          Проведение тренингов по MS Excel
'          https://www.excel-vba.ru
'          info@excel-vba.ru
'          WebMoney - R298726502453; Яндекс.Деньги - 41001332272872
' Purpose: https://www.excel-vba.ru/chto-umeet-excel/sobstvennoe-menyu-v-textbox
'---------------------------------------------------------------------------------------
'   при вызове формы
'   создаем свое меню "Вырезать-Копировать-Вставить" для вызова из TextBox
'   и для всех TextBox-ов назначаем обработку событий через модуль класса clsCustomMenu
'   Подробнее про модули классов: https://www.excel-vba.ru/chto-umeet-excel/rabota-s-modulyami-klassov/

Option Explicit

Public WithEvents oTbx As MSForms.TextBox

Private Sub oTbx_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        Set tbxAct = oTbx
        Application.CommandBars(sPopupMenuName).ShowPopup
    End If
End Sub