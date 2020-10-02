'-----------------------------------------------------------------------------
'                       Константы для маркировки шейпов
'                                SAPR_ASU_Type
'                Числа располагаются в User.SAType каждой фигуры
'-----------------------------------------------------------------------------

Public Const typeNO As Integer = 0 'Контакт NO (Дочерний)(Не нумеруется)
Public Const typeNC As Integer = 1 'Контакт NC (Дочерний)(Не нумеруется)

Public Const typeCoil As Integer = 2 'Катушка реле (Родительский) KL, KM, KT, KV, KK
Public Const typeParent As Integer = 3 'Нумеруемый элемент схемы без катушки (Родительский) SA, SB, QF, SF, QS, QA, FU, RU, DD (ТРМ, ПЛК-моноблок)
Public Const typeElement As Integer = 4 'Нумеруемый элемент схемы без контактов (НЕ Родительский) HL, XS

'ПЛК - разнесенное отображение
Public Const typePLCTerm As Integer = 10 'Клемма внутри ПЛК. (НЕ Родительский) (Не нумеруется)
Public Const typePLCIOChild As Integer = 11 'Вход/Выход внутри дочернего модуля дочернего ПЛК. Содержит несколько клемм (НЕ Родительский) (Не нумеруется)
Public Const typePLCIOParent As Integer = 12 'Вход/Выход внутри родительского модуля родительского ПЛК. Содержит несколько клемм (НЕ Родительский) (Не нумеруется)
Public Const typePLCModChild As Integer = 13 'Модуль внутри ПЛК. Содержит несколько Входов/Выходов (Дочерний)(Не нумеруется)
Public Const typePLCChild As Integer = 14 'Кусок ПЛК при разнесенном отображении. Содержит несколько Модулей (Дочерний)(Не нумеруется)
Public Const typePLCModParent As Integer = 15 'Модуль при разнесенном ПЛК. Содержит описание Входов/Выходов в виде монтажного отбражения (Родительский)(Нумеруется вручную) AI DI AO DO
Public Const typePLCParent As Integer = 16 'ПЛК при разнесенном отображении. Содержит ВСЕ модули в виде монтажного отбражения (Родительский)  DD

Public Const typeThumb As Integer = 30 'Миниатюры контактов (Не нумеруется)

Public Const typeWireLinkS As Integer = 40 'Разрыв провода Источник (Родительский)(Не нумеруется)
Public Const typeWireLinkR As Integer = 45 'Разрыв провода Приемник (Дочерний)(Не нумеруется)

Public Const typeTerminal As Integer = 50 'Клеммы в шкафу, в распределительной коробке

Public Const typeWire As Integer = 60 'Провода внутри шкафа

Public Const typeCableSH As Integer = 70 'Кабель вне шкафа на схеме электрической принципиальной
Public Const typeCableVP As Integer = 80 'Кабель вне шкафа на схеме внешних проводок (Не нумеруется)
Public Const typeCablePL As Integer = 90 'Кабель вне шкафа на ПЛАНЕ оборудования и КИП (Не нумеруется)

Public Const typeVynoskaPL As Integer = 95 'Выноска на ПЛАНЕ оборудования и КИП (Не нумеруется)

Public Const typeActuator As Integer = 100 'Привод вне шкафа
Public Const typeSensor As Integer = 110 'Датчик вне шкафа
Public Const typeFSASensor As Integer = 120 'Датчик на ФСА
Public Const typeFSAPodval As Integer = 130 'Датчик на ФСА в подвале

Public Const typeBox As Integer = 140 'Шкафы, распределительные коробки (Не нумеруется)

Public Const typeElectroOneWire As Integer = 150 'Однолинейная схема

Public Const typeVidShkafa As Integer = 160 'Внешний вид шкафа (Не нумеруется)

Public Const typeDuctPlan As Integer = 170 'Лотки на ПЛАНЕ, кабельные трассы (Не нумеруется)
Public Const typeElectroPlan As Integer = 180 'ЭС ЭО на плане
Public Const typeOPSPlan As Integer = 190 'ОПС на плане



