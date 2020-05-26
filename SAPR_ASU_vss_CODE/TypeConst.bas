Attribute VB_Name = "TypeConst"
'-----------------------------------------------------------------------------
'                       Константы для маркировки шейпов
'-----------------------------------------------------------------------------

Public Const typeNO As Integer = 0 'Контакт NO (Дочерний)(Не нумеруется)

Public Const typeNC As Integer = 1 'Контакт NC (Дочерний)(Не нумеруется)

Public Const typeCoil As Integer = 2 'Катушка реле (Родительский) KL, KM, KT, KV, KK
Public Const typeParent As Integer = 3 'Нумеруемый элемент схемы без катушки (Родительский) SA, SB, QF, SF, QS, QA, FU, RU
Public Const typeElement As Integer = 4 'Нумеруемый элемент схемы без контактов (НЕ Родительский) HL, DD (ПЛК, ТРМ), XS

Public Const typeThumb As Integer = 30 'Миниатюры контактов (Не нумеруется)

Public Const typeWireLinkS As Integer = 40 'Разрыв провода Источник (Родительский)(Не нумеруется)
Public Const typeWireLinkR As Integer = 45 'Разрыв провода Приемник (Дочерний)(Не нумеруется)

Public Const typeTerminal As Integer = 50 'Клеммы в шкафу, в распределительной коробке

Public Const typeWire As Integer = 60 'Провода внутри шкафа

Public Const typeSensor As Integer = 70 'Датчик вне шкафа
Public Const typeActuator As Integer = 75 'Привод вне шкафа

Public Const typeCableSH As Integer = 80 'Кабель вне шкафа на схеме электрической принципиальной
Public Const typeCableVP As Integer = 85 'Кабель вне шкафа на схеме внешних проводок (Не нумеруется)

Public Const typeCablePL As Integer = 90 'Кабель вне шкафа на ПЛАНЕ оборудования и КИП (Не нумеруется)

Public Const typeFSASensor As Integer = 100 'Датчик на ФСА

Public Const typeFSAPodval As Integer = 110 'Датчик на ФСА в подвале

Public Const typeDuctPlan As Integer = 120 'Лотки на ПЛАНЕ, кабельные трассы (Не нумеруется)

Public Const typeElectroPlan As Integer = 130 'ЭС ЭО на плане

Public Const typeElectroOneWire As Integer = 140 'Однолинейная схема

Public Const typeOPSPlan As Integer = 150 'ОПС на плане

Public Const typeVidShkafa As Integer = 160 'Внешний вид шкафа (Не нумеруется)

Public Const typeBox As Integer = 170 'Шкафы, распределительные коробки (Не нумеруется)


