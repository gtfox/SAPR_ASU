'------------------------------------------------------------------------------------------------------------
' Module        : Dijkstra_Levit - Нахождение кратчайших путей из одной точки в другую
' Author        : gtfox
' Date          : 2022.02.17
' Description   : Автопрокладка кабелей по лоткам, подсчет длины, выноски кабелей
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе:
                '------------------------------------------------------------------------------------------------------------
                ' Module        : MyGraph - Нахождение кратчайших путей из одной точки в другую
                ' Author        : MCH (Михаил Ч.)
                ' Date          : 10.2013
                ' Description   : Реализация на VBA  алгоритма Дейкстры и алгоритма Левита, по нахождению кратчайших путей
                ' Links         : http://www.excelworld.ru/forum/3-6656-1, https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=read&FID=1&TID=51651&MESSAGE_TYPE=EDIT&sessid=cbfd1f4b9a600cb73acef75989396ff9&result=edit
                '------------------------------------------------------------------------------------------------------------
Option Explicit

Private Const INF As Double = 1E+100 'значение бесконечности
Private Const maxEdge As Long = 20 'максимальное кол-во ребер для каждой вершины

Private Type Vertex 'тип для описания вершин
    name As String 'наименование вершины
    d As Double 'дистанция до текущей вершины
    p As Long '"предок" до текущей вершины
    u As Boolean 'метка о прохождении вершины, используется в алгоритме Дейкстры
    id As Long 'принадлежность к множествам, используется в алгоритме Левита
    edgeCount As Long 'количество ребер
    nGraph(1 To maxEdge) As Long 'массив смежных вершин
    dGraph(1 To maxEdge) As Double 'массив дистанций до смежных вершин
End Type

Sub Main() 'алгоритмом Дейкстры
    Dim out
    out = MyDijkstra(Range("Graph"), Range("Vertex"), [b3].Value, [c3].Value)
    Range("B6:E22").ClearContents
    If IsArray(out) Then
        Range("B6").ReSize(UBound(out), 4) = out
    End If
End Sub

Sub Main2() 'алгоритмом Левита
    Dim out
    out = MyLevit(Range("Graph"), Range("Vertex"), [b3].Value, [c3].Value)
    Range("B6:E22").ClearContents
    If IsArray(out) Then
        Range("B6").ReSize(UBound(out), 4) = out
    End If
End Sub

Function MyDijkstra(rngGraph As Range, rngVertex As Range, s1 As Long, s2 As Long)
    'реализовано в виде функции, возвращающей массив
    'rngGraph - исходный граф
    'rngVertex - перечень вершин
    's1 - начальная вершина
    's2 - конечная вершина
    
    Dim g() As Vertex
    Dim n As Long, i As Long, j As Long, v As Long, t As Long, d As Double
    
    Call MakeGraph(g, rngGraph, rngVertex) 'создаем граф
    n = UBound(g) 'количество вершин
    g(s1).d = 0 'дистанция до начальной точки равна нулю
    g(s1).p = s1 'предок отсутствует
    
    For i = 1 To n
        v = -1
        For j = 1 To n 'ищем вершину с минимальной дистанцией
            If Not g(j).u Then If v = -1 Then v = j Else If g(j).d < g(v).d Then v = j
        Next j
        If g(v).d = INF Or v = s2 Then Exit For 'если дистанция бесконечность, либо вершина равна конечной, то прекращаем поиск
        g(v).u = True 'вершина просмотрена
        For j = 1 To g(v).edgeCount 'проходим по всем смежным вершинам
            t = g(v).nGraph(j) 'смежная вершина
            d = g(v).dGraph(j) 'дистанция до нее
            If g(v).d + d < g(t).d Then 'если расстояние короче, чем уже посчитано
                g(t).d = g(v).d + d 'улучшаем расстояние
                g(t).p = v 'запоминаем предка
            End If
        Next j
    Next i
    
    If g(s2).p <> 0 Then 'если путь найден
        ReDim tmp(1 To n, 1 To 3)
        j = s2
        i = 0
        Do 'заносим элементы пути во временный массив
           i = i + 1
           tmp(i, 1) = g(g(j).p).name
           tmp(i, 2) = g(j).name
           tmp(i, 3) = g(j).d
           j = g(j).p
        Loop While j <> s1
        
        ReDim out(1 To i, 1 To 4) 'создаем выходной массив, переворачивая временный
        d = 0
        For j = 1 To i
            out(j, 1) = tmp(i - j + 1, 1)
            out(j, 2) = tmp(i - j + 1, 2)
            out(j, 4) = tmp(i - j + 1, 3)
            out(j, 3) = out(j, 4) - d
            d = d + out(j, 3)
        Next j
        MyDijkstra = out
    End If
End Function

Function MyLevit(rngGraph As Range, rngVertex As Range, s1 As Long, s2 As Long)
    'реализовано в виде функции, возвращающей массив
    'rngGraph - исходный граф
    'rngVertex - перечень вершин
    's1 - начальная вершина
    's2 - конечная вершина
    
    Dim g() As Vertex
    Dim n As Long, i As Long, j As Long, v As Long, t As Long, d As Double
    Dim qh As Long, qt As Long 'индексы в очереди
    
    Call MakeGraph(g, rngGraph, rngVertex) 'создаем граф
    n = UBound(g) 'количество вершин
    ReDim q(1 To n) As Long 'массив индексов очереди
    g(s1).d = 0 'дистанция до начальной точки равна нулю
    g(s1).p = s1 'предок отсутствует
    qh = 1 'начало очереди
    q(qh) = s1 'сохраняем в очередь начальную вершину
    qt = qh + 1 'индекс на последующий элемент
    
    While qh <> qt 'пока очередь не пуста
        v = q(qh) 'вершина из начала очереди
        qh = qh Mod n + 1 'удаляем элемент (сдвигаем начало очереди на единицу)
        g(v).id = 2 'изменяем принадлежность множества
        For j = 1 To g(v).edgeCount 'проходим по всем ребрам данной вершины
            t = g(v).nGraph(j) 'смежная вершина
            d = g(v).dGraph(j) 'дистанция до нее
            If g(v).d + d < g(t).d Then 'если расстояние короче, чем уже посчитано
                g(t).d = g(v).d + d 'улучшаем расстояние
                g(t).p = v 'запоминаем предка
                If g(t).id = 0 Then 'если вершина еще не обрабатывалась
                    q(qt) = t 'помещаем ее в конец очереди
                    qt = qt Mod n + 1 'вычисляем конец очереди
                    g(t).id = 1 'меняем статус вершины
                ElseIf g(t).id = 2 Then 'если вершина уже считалась
                    qh = IIf(qh = 1, n, qh - 1) 'сдвигаем начало очереди
                    q(qh) = t 'помещаем вершину в начало очереди
                    g(t).id = 1 'меняем статус вершины
                End If
            End If
        Next j
    Wend
        
    If g(s2).p <> 0 Then 'если путь найден
        ReDim tmp(1 To n, 1 To 3)
        j = s2
        i = 0
        Do 'заносим элементы пути во временный массив
           i = i + 1
           tmp(i, 1) = g(g(j).p).name
           tmp(i, 2) = g(j).name
           tmp(i, 3) = g(j).d
           j = g(j).p
        Loop While j <> s1
        
        ReDim out(1 To i, 1 To 4) 'создаем выходной массив, переворачивая временный
        d = 0
        For j = 1 To i
            out(j, 1) = tmp(i - j + 1, 1)
            out(j, 2) = tmp(i - j + 1, 2)
            out(j, 4) = tmp(i - j + 1, 3)
            out(j, 3) = out(j, 4) - d
            d = d + out(j, 3)
        Next j
        MyLevit = out
    End If
End Function

Sub MakeGraph(graph() As Vertex, rngGraph As Range, rngVertex As Range) 'процедура создания графа
    Dim DataGraph, DataVertex, i As Long, j As Long, k As Long
    
    DataGraph = rngGraph.Value 'двумерный массив с путями
    DataVertex = rngVertex.Value 'массив с наименованием вершин
    
    ReDim graph(1 To UBound(DataVertex)) As Vertex
    For i = 1 To UBound(DataGraph)
        If i <= UBound(DataVertex) Then graph(i).name = DataVertex(i, 1)
        graph(i).d = INF
        k = 0
        For j = 1 To UBound(DataGraph, 2)
            If DataGraph(i, j) > 0 Then
                k = k + 1
                graph(i).edgeCount = k
                graph(i).nGraph(k) = j
                graph(i).dGraph(k) = DataGraph(i, j)
            End If
    Next j, i
End Sub