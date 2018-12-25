Attribute VB_Name = "mod_QuickSortNonRecursive"

 Option Explicit

 Private Type QuickStack
     'тип для QuickSort
     Low As Long
     High As Long
 End Type


 Public Sub QuickSortNonRecursive(SortArray() As Variant)

 Dim i As Long, j As Long, lb As Long, ub As Long
 Dim stack() As QuickStack, stackpos As Long, ppos As Long, pivot As Variant, swp

     ReDim stack(1 To 1024)
     stackpos = 1

     stack(1).Low = LBound(SortArray)
     stack(1).High = UBound(SortArray)

     Do
         'Взять границы lb и ub текущего массива из стека.
         lb = stack(stackpos).Low
         ub = stack(stackpos).High
         stackpos = stackpos - 1
         Do
             'Шаг 1. Разделение по элементу pivot
             ppos = (lb + ub) \ 2
             i = lb: j = ub: pivot = SortArray(ppos)

             Do
                 Do While SortArray(i) < pivot: i = i + 1: Loop
                 Do While pivot < SortArray(j): j = j - 1: Loop
                 If i <= j Then
                     swp = SortArray(i): SortArray(i) = SortArray(j): SortArray(j) = swp
 '                    Swap SortArray(i), SortArray(j)
                     i = i + 1
                     j = j - 1
                 End If
             Loop While i <= j

             'Сейчас указатель i указывает на начало правого подмассива,
             'j - на конец левого lb ? j ? i ? ub.
             'Возможен случай, когда указатель i или j выходит за границу массива
             'Шаги 2, 3. Отправляем большую часть в стек  и двигаем lb,ub

             If i < ppos Then   'правая часть больше
                 If i < ub Then
                     stackpos = stackpos + 1
                     stack(stackpos).Low = i
                     stack(stackpos).High = ub
                 End If
                 ub = j        'следующая итерация разделения будет работать с левой частью
             Else
                 If j > lb Then
                     stackpos = stackpos + 1
                     stack(stackpos).Low = lb
                     stack(stackpos).High = j
                 End If
                 lb = i
             End If
         Loop While lb < ub
     Loop While stackpos
 End Sub
