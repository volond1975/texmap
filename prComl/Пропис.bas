Attribute VB_Name = "Пропис"
'
' Преобразование даты в Украинский Формат прописью
' Макрос записан 25.08.96 (Корольков Владислав)
'
Sub Пропис_Д(n, s$)
Dim naim(1 To 12) As String
naim(1) = " січня "
naim(2) = " лютого "
naim(3) = " березня "
naim(4) = " квітня "
naim(5) = " травня "
naim(6) = " червня "
naim(7) = " липня "
naim(8) = " серпня "
naim(9) = " вересня "
naim(10) = " жовтня "
naim(11) = " листопада "
naim(12) = " грудня "
I = Month(n)
s$ = " "" " + str(Day(n)) + " "" " + naim(I) + str(Year(n)) + " р."
End Sub
'
' Преобразование числа в сумму прописью
' Макрос записан 25.08.96 (Корольков Владислав)
' Макрос может быть применен в другом макросе как подпрограмма
' его вызов Call Пропис(N, s$)
' В параметрах указывается исходнае число и поле для текстового результата
Sub Пропис(n, s$)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' Копейки
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' Миллиарды
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' Миллионы
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' Тысячи
   r = r - ts * 1000
   eee = Int(r)                  ' до тысячи
   s$ = f$(MLRD, 3) + f$(mln, 2) + f$(ts, 1) + f$(eee, 0) + " грн. " + Kop + " коп."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "Ноль грн. " + Kop + " коп."
End If
End Sub
' Функция для преобразование числа в сумму прописью
Function грн$(n)
Dim Kop, MLRD, r, mln, ts, eee, s$
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' Копейки
Kop = VBA.Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' Миллиарды
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' Миллионы
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' Тысячи
   r = r - ts * 1000
   eee = Int(r)                  ' до тысячи
   s$ = f$(MLRD, 3) + f$(mln, 2) + f$(ts, 1) + f$(eee, 0) + " грн. " + Kop + " коп."
   s$ = VBA.UCase(VBA.Mid(s$, 2, 1)) + VBA.Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "Ноль грн. " + Kop + " коп."
End If
грн$ = s$
End Function
Function грн1(n)
z = Int(Application.Round(n, 2))
Kop = (n - Int(n)) * 100  ' Копейки
Kop = Format(Kop, "#00")
If n >= 1 Then
  
   s = z & " грн. " & Kop & " коп."
   
          Else
   s = "0 грн. " & Kop & " коп."
End If
грн1 = s
End Function












'
' Преобразование числа до 999 в сумму прописью
' Макрос записан 25.08.96 (Корольков Владислав)
' до тысячи R = 0  0-999
' тысячи    R = 1  1-999 тысяч
' миллионы  R = 2  1-999 миллионов
' миллиарды R = 3  1-999 миллиардов
Function f$(n, r)
Dim s$, ed, I, des, sot
    s$ = ""              ' Установка начального значения результата
   If n = 0 Then GoTo Kon
' Вычисление количества единиц, десятков, сотен
   ed = n Mod 10
   I = Int(n / 10)
   des = I Mod 10
   I = Int(I / 10)
   sot = I Mod 10
' Формирование строки прописью
  If des = 1 Then
     Select Case ed
            Case 1
                 s$ = " одинадцять"
            Case 2
                 s$ = " дванадцять"
            Case 3
                 s$ = " тринадцять"
            Case 4
                 s$ = " чотирнадцять"
            Case 5
                 s$ = " п'ятнадцять"
            Case 6
                 s$ = " шістьнадцять"
            Case 7
                 s$ = " сімнадцять"
            Case 8
                 s$ = " вісімнадцять"
            Case 9
                 s$ = " дев'ятнадцять"
            Case Else
                 s$ = " десять"
     End Select
            Else
         Select Case ed
            Case 1
                 s$ = " одна"
                 If r > 1 Then s$ = " один"
            Case 2
                 s$ = " дві"
                 If r > 1 Then s$ = " два"
            Case 3
                 s$ = " три"
            Case 4
                 s$ = " чотири"
            Case 5
                 s$ = " п'ять"
            Case 6
                 s$ = " шість"
            Case 7
                 s$ = " сім"
            Case 8
                 s$ = " вісім"
            Case 9
                 s$ = " дев'ять"
            Case Else
                 s$ = ""
     End Select
   End If
     Select Case des
            Case 2
                 s$ = " двадцять" + s$
            Case 3
                 s$ = " тридцять" + s$
            Case 4
                 s$ = " сорок" + s$
            Case 5
                 s$ = " п'ятдесят" + s$
            Case 6
                 s$ = " шістдесят" + s$
            Case 7
                 s$ = " сімдесят" + s$
            Case 8
                 s$ = " вісімдесят" + s$
            Case 9
                 s$ = " дев'яносто" + s$
            Case Else
                 s$ = s$
     End Select
         Select Case sot
            Case 1
                 s$ = " сто" + s$
            Case 2
                 s$ = " двісті" + s$
            Case 3
                 s$ = " триста" + s$
            Case 4
                 s$ = " чотириста" + s$
            Case 5
                 s$ = " п'ятсот" + s$
            Case 6
                 s$ = " шістсот" + s$
            Case 7
                 s$ = " сімсот" + s$
            Case 8
                 s$ = " вісімсот" + s$
            Case 9
                 s$ = " дев'ятсот" + s$
            Case Else
                 s$ = s$
     End Select
' формирование наименований по группам тысяча-миллиард
     If des = 1 Then       ' наименование для диапазона 11-19
                 Select Case r
                        Case 0
                             s$ = s$     '+ " гривень"
                        Case 1
                             s$ = s$ + " тисяч"
                        Case 2
                             s$ = s$ + " мільйонів"
                        Case 3
                             s$ = s$ + " мільярдів"
                        Case Else
                             s$ = s$
                  End Select
' наименование по последней цифре
            Else
                Select Case ed
                       Case 1  ' один
                          Select Case r
                                 Case 0
                                      s$ = s$     '+ " гривня"
                                 Case 1
                                      s$ = s$ + " тисяча"
                                 Case 2
                                      s$ = s$ + " мільйон"
                                 Case 3
                                      s$ = s$ + " мільярд"
                                 Case Else
                                      s$ = s$
                          End Select
                       Case 2 To 4 ' два - четыре
                          Select Case r
                                 Case 0
                                      s$ = s$    ' + " гривні"
                                 Case 1
                                      s$ = s$ + " тисячі"
                                 Case 2
                                      s$ = s$ + " мільйона"
                                 Case 3
                                      s$ = s$ + " мільярда"
                                 Case Else
                                      s$ = s$
                           End Select
                       Case Else  ' остальные
                          Select Case r
                                 Case 0
                                      s$ = s$     ' + " гривень"
                                 Case 1
                                      s$ = s$ + " тисяч"
                                 Case 2
                                      s$ = s$ + " мільйонів"
                                 Case 3
                                      s$ = s$ + " мільярдів"
                                 Case Else
                                      s$ = s$
                           End Select
                End Select
      End If
Kon:
     f$ = s$
End Function
'
' Преобразование даты в Русский Формат прописью
' Макрос записан 25.08.96 (Корольков Владислав)
'
Sub Пропис_Д_Р(n, s$)
Dim naim(1 To 12) As String
naim(1) = " января "
naim(2) = " февраля "
naim(3) = " марта "
naim(4) = " апреля "
naim(5) = " мая "
naim(6) = " июня "
naim(7) = " июля "
naim(8) = " августа "
naim(9) = " сентября "
naim(10) = " октября "
naim(11) = " ноября "
naim(12) = " декабря "
I = Month(n)
s$ = " "" " + str(Day(n)) + " "" " + naim(I) + str(Year(n)) + " р."
End Sub
'
' Преобразование числа в сумму прописью
' Макрос записан 25.08.96 (Корольков Владислав)
'
Sub Пропис_РГ(n, s$)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' Копейки
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' Миллиарды
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' Миллионы
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' Тысячи
   r = r - ts * 1000
   eee = Int(r)                  ' до тысячи
   s$ = FR$(MLRD, 3) + FR$(mln, 2) + FR$(ts, 1) + FR$(eee, 0) + " гривн. " + Kop + " коп."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "Ноль грн. " + Kop + " коп."
End If
End Sub
Function грн_р$(n)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' Копейки
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' Миллиарды
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' Миллионы
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' Тысячи
   r = r - ts * 1000
   eee = Int(r)                  ' до тысячи
   s$ = FR$(MLRD, 3) + FR$(mln, 2) + FR$(ts, 1) + FR$(eee, 0) + " грн. " + Kop + " коп."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "Ноль грн. " + Kop + " коп."
End If
грн_р$ = s$
End Function
Function руб$(n)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' Копейки
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' Миллиарды
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' Миллионы
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' Тысячи
   r = r - ts * 1000
   eee = Int(r)                  ' до тысячи
   s$ = FR$(MLRD, 3) + FR$(mln, 2) + FR$(ts, 1) + FR$(eee, 0) + " руб. " + Kop + " коп."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "Ноль руб. " + Kop + " коп."
End If
руб$ = s$

End Function
'
' Преобразование числа до 999 в сумму прописью
' Макрос записан 25.08.96 (Корольков Владислав)
' до тысячи R = 0  0-999
' тысячи    R = 1  1-999 тысяч
' миллионы  R = 2  1-999 миллионов
' миллиарды R = 3  1-999 миллиардов
Function FR$(n, r)
   s$ = ""              ' Установка начального значения результата
   If n = 0 Then GoTo Kon
' Вычисление количества единиц, десятков, сотен
   ed = n Mod 10
   I = Int(n / 10)
   des = I Mod 10
   I = Int(I / 10)
   sot = I Mod 10
' Формирование строки прописью
  If des = 1 Then
     Select Case ed
            Case 1
                 s$ = " одиннадцать"
            Case 2
                 s$ = " двенадцать"
            Case 3
                 s$ = " тринадцать"
            Case 4
                 s$ = " четырнадцать"
            Case 5
                 s$ = " пятнадцать"
            Case 6
                 s$ = " шестнадцать"
            Case 7
                 s$ = " семнадцать"
            Case 8
                 s$ = " восемнадцать"
            Case 9
                 s$ = " девятнадцать"
            Case Else
                 s$ = " десять"
     End Select
            Else
         Select Case ed
            Case 1
                 s$ = " одна"
                 If r > 1 Then s$ = " один"
            Case 2
                 s$ = " две"
                 If r > 1 Then s$ = " два"
            Case 3
                 s$ = " три"
            Case 4
                 s$ = " четыре"
            Case 5
                 s$ = " пять"
            Case 6
                 s$ = " шесть"
            Case 7
                 s$ = " семь"
            Case 8
                 s$ = " восемь"
            Case 9
                 s$ = " девять"
            Case Else
                 s$ = ""
     End Select
   End If
     Select Case des
            Case 2
                 s$ = " двадцать" + s$
            Case 3
                 s$ = " тридцать" + s$
            Case 4
                 s$ = " сорок" + s$
            Case 5
                 s$ = " пятьдесят" + s$
            Case 6
                 s$ = " шестьдесят" + s$
            Case 7
                 s$ = " семьдесят" + s$
            Case 8
                 s$ = " восемьдесят" + s$
            Case 9
                 s$ = " девяносто" + s$
            Case Else
                 s$ = s$
     End Select
         Select Case sot
            Case 1
                 s$ = " сто" + s$
            Case 2
                 s$ = " двести" + s$
            Case 3
                 s$ = " триста" + s$
            Case 4
                 s$ = " четыреста" + s$
            Case 5
                 s$ = " пятьсот" + s$
            Case 6
                 s$ = " шестьсот" + s$
            Case 7
                 s$ = " семьсот" + s$
            Case 8
                 s$ = " восемьсот" + s$
            Case 9
                 s$ = " девятьсот" + s$
            Case Else
                 s$ = s$
     End Select
' формирование наименований по группам тысяча-миллиард
     If des = 1 Then       ' наименование для диапазона 11-19
                 Select Case r
                        Case 0
                             s$ = s$     '+ " гривень"
                        Case 1
                             s$ = s$ + " тысяч"
                        Case 2
                             s$ = s$ + " миллионов"
                        Case 3
                             s$ = s$ + " милиардов"
                        Case Else
                             s$ = s$
                  End Select
' наименование по последне цифре
            Else
                Select Case ed
                       Case 1  ' один
                          Select Case r
                                 Case 0
                                      s$ = s$     '+ " гривня"
                                 Case 1
                                      s$ = s$ + " тысяча"
                                 Case 2
                                      s$ = s$ + " миллион"
                                 Case 3
                                      s$ = s$ + " милиард"
                                 Case Else
                                      s$ = s$
                          End Select
                       Case 2 To 4 ' два - четыре
                          Select Case r
                                 Case 0
                                      s$ = s$    ' + " гривні"
                                 Case 1
                                      s$ = s$ + " тысячи"
                                 Case 2
                                      s$ = s$ + " миллиона"
                                 Case 3
                                      s$ = s$ + " милиарда"
                                 Case Else
                                      s$ = s$
                           End Select
                       Case Else  ' остальные
                          Select Case r
                                 Case 0
                                      s$ = s$     ' + " гривень"
                                 Case 1
                                      s$ = s$ + " тысяч"
                                 Case 2
                                      s$ = s$ + " миллионов"
                                 Case 3
                                      s$ = s$ + " милиардов"
                                 Case Else
                                      s$ = s$
                           End Select
                End Select
      End If
Kon:
     FR$ = s$
End Function
'
' Преобразование числа в сумму прописью в долларах
' Макрос записан 25.08.96 (Корольков Владислав)
'
Sub Пропис_Дол(n, s$, pr1$, pr2$)
к$ = str(n)
m = InStr(к$, ".")
dr$ = "0"
If m > 0 Then
   dr$ = Mid(к$, m + 1)
End If
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' Миллиарды
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' Миллионы
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' Тысячи
   r = r - ts * 1000
   eee = Int(r)                  ' до тысячи
   s$ = FRD$(MLRD, 3) + FRD$(mln, 2) + FRD$(ts, 1) + FRD$(eee, 0) + pr1$ + dr$ + pr2$
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "Ноль " + pr1$ + dr$ + pr2$
End If
End Sub
Function ЕдИзм$(n, pr1$, pr2$)
к$ = str(n)
m = InStr(к$, ".")
dr$ = "0"
If m > 0 Then
   dr$ = Mid(к$, m + 1)
End If
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' Миллиарды
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' Миллионы
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' Тысячи
   r = r - ts * 1000
   eee = Int(r)                  ' до тысячи
   s$ = FRD$(MLRD, 3) + FRD$(mln, 2) + FRD$(ts, 1) + FRD$(eee, 0) + pr1$ + dr$ + pr2$
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "Ноль " + pr1$ + dr$ + pr2$
End If
ЕдИзм$ = s$
End Function
'
' Преобразование числа до 999 в сумму прописью
' Макрос записан 25.08.96 (Корольков Владислав)
' до тысячи R = 0  0-999
' тысячи    R = 1  1-999 тысяч
' миллионы  R = 2  1-999 миллионов
' миллиарды R = 3  1-999 миллиардов
Function FRD$(n, r)
   s$ = ""              ' Установка начального значения результата
   If n = 0 Then GoTo Kon
' Вычисление количества единиц, десятков, сотен
   ed = n Mod 10
   I = Int(n / 10)
   des = I Mod 10
   I = Int(I / 10)
   sot = I Mod 10
' Формирование строки прописью
  If des = 1 Then
     Select Case ed
            Case 1
                 s$ = " одиннадцать"
            Case 2
                 s$ = " двенадцать"
            Case 3
                 s$ = " тринадцать"
            Case 4
                 s$ = " четырнадцать"
            Case 5
                 s$ = " пятнадцать"
            Case 6
                 s$ = " шестнадцать"
            Case 7
                 s$ = " семнадцать"
            Case 8
                 s$ = " восемнадцать"
            Case 9
                 s$ = " девятнадцать"
            Case Else
                 s$ = " десять"
     End Select
            Else
         Select Case ed
            Case 1
                 s$ = " один"
            Case 2
                 s$ = " два"
            Case 3
                 s$ = " три"
            Case 4
                 s$ = " четыре"
            Case 5
                 s$ = " пять"
            Case 6
                 s$ = " шесть"
            Case 7
                 s$ = " семь"
            Case 8
                 s$ = " восемь"
            Case 9
                 s$ = " девять"
            Case Else
                 s$ = ""
     End Select
   End If
     Select Case des
            Case 2
                 s$ = " двадцать" + s$
            Case 3
                 s$ = " тридцать" + s$
            Case 4
                 s$ = " сорок" + s$
            Case 5
                 s$ = " пятьдесят" + s$
            Case 6
                 s$ = " шестьдесят" + s$
            Case 7
                 s$ = " семьдесят" + s$
            Case 8
                 s$ = " восемьдесят" + s$
            Case 9
                 s$ = " девяносто" + s$
            Case Else
                 s$ = s$
     End Select
         Select Case sot
            Case 1
                 s$ = " сто" + s$
            Case 2
                 s$ = " двести" + s$
            Case 3
                 s$ = " триста" + s$
            Case 4
                 s$ = " четыреста" + s$
            Case 5
                 s$ = " пятьсот" + s$
            Case 6
                 s$ = " шестьсот" + s$
            Case 7
                 s$ = " семьсот" + s$
            Case 8
                 s$ = " восемьсот" + s$
            Case 9
                 s$ = " девятьсот" + s$
            Case Else
                 s$ = s$
     End Select
' формирование наименований по группам тысяча-миллиард
     If des = 1 Then       ' наименование для диапазона 11-19
                 Select Case r
                        Case 0
                             s$ = s$     '+ " гривень"
                        Case 1
                             s$ = s$ + " тысяч"
                        Case 2
                             s$ = s$ + " миллионов"
                        Case 3
                             s$ = s$ + " милиардов"
                        Case Else
                             s$ = s$
                  End Select
' наименование по последне цифре
            Else
                Select Case ed
                       Case 1  ' один
                          Select Case r
                                 Case 0
                                      s$ = s$     '+ " гривня"
                                 Case 1
                                      s$ = s$ + " тысяча"
                                 Case 2
                                      s$ = s$ + " миллион"
                                 Case 3
                                      s$ = s$ + " милиард"
                                 Case Else
                                      s$ = s$
                          End Select
                       Case 2 To 4 ' два - четыре
                          Select Case r
                                 Case 0
                                      s$ = s$    ' + " гривні"
                                 Case 1
                                      s$ = s$ + " тысячи"
                                 Case 2
                                      s$ = s$ + " миллиона"
                                 Case 3
                                      s$ = s$ + " милиарда"
                                 Case Else
                                      s$ = s$
                           End Select
                       Case Else  ' остальные
                          Select Case r
                                 Case 0
                                      s$ = s$     ' + " гривень"
                                 Case 1
                                      s$ = s$ + " тысяч"
                                 Case 2
                                      s$ = s$ + " миллионов"
                                 Case 3
                                      s$ = s$ + " милиардов"
                                 Case Else
                                      s$ = s$
                           End Select
                End Select
      End If
Kon:
     FRD$ = s$
End Function

Function НазваМесяца(Дата)
Dim naim(1 To 12) As String
naim(1) = "Січень"
naim(2) = "Лютий"
naim(3) = "Березень"
naim(4) = "Квітень"
naim(5) = "Травень"
naim(6) = "Червень"
naim(7) = "Липень"
naim(8) = "Серпень"
naim(9) = "Вересень"
naim(10) = "Жовтень"
naim(11) = "Листопад"
naim(12) = "Грудень"

I = Month(Дата)
НазваМесяца = naim(I)
End Function


'DV-7kiR8DMzxy0H2fnITKABIR # Do not ove this line; required for DocVerse merge.
