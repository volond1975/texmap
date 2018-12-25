Attribute VB_Name = "������"
'
' �������������� ���� � ���������� ������ ��������
' ������ ������� 25.08.96 (��������� ���������)
'
Sub ������_�(n, s$)
Dim naim(1 To 12) As String
naim(1) = " ���� "
naim(2) = " ������ "
naim(3) = " ������� "
naim(4) = " ����� "
naim(5) = " ������ "
naim(6) = " ������ "
naim(7) = " ����� "
naim(8) = " ������ "
naim(9) = " ������� "
naim(10) = " ������ "
naim(11) = " ��������� "
naim(12) = " ������ "
I = Month(n)
s$ = " "" " + str(Day(n)) + " "" " + naim(I) + str(Year(n)) + " �."
End Sub
'
' �������������� ����� � ����� ��������
' ������ ������� 25.08.96 (��������� ���������)
' ������ ����� ���� �������� � ������ ������� ��� ������������
' ��� ����� Call ������(N, s$)
' � ���������� ����������� �������� ����� � ���� ��� ���������� ����������
Sub ������(n, s$)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' �������
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' ���������
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' ��������
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' ������
   r = r - ts * 1000
   eee = Int(r)                  ' �� ������
   s$ = f$(MLRD, 3) + f$(mln, 2) + f$(ts, 1) + f$(eee, 0) + " ���. " + Kop + " ���."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "���� ���. " + Kop + " ���."
End If
End Sub
' ������� ��� �������������� ����� � ����� ��������
Function ���$(n)
Dim Kop, MLRD, r, mln, ts, eee, s$
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' �������
Kop = VBA.Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' ���������
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' ��������
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' ������
   r = r - ts * 1000
   eee = Int(r)                  ' �� ������
   s$ = f$(MLRD, 3) + f$(mln, 2) + f$(ts, 1) + f$(eee, 0) + " ���. " + Kop + " ���."
   s$ = VBA.UCase(VBA.Mid(s$, 2, 1)) + VBA.Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "���� ���. " + Kop + " ���."
End If
���$ = s$
End Function
Function ���1(n)
z = Int(Application.Round(n, 2))
Kop = (n - Int(n)) * 100  ' �������
Kop = Format(Kop, "#00")
If n >= 1 Then
  
   s = z & " ���. " & Kop & " ���."
   
          Else
   s = "0 ���. " & Kop & " ���."
End If
���1 = s
End Function












'
' �������������� ����� �� 999 � ����� ��������
' ������ ������� 25.08.96 (��������� ���������)
' �� ������ R = 0  0-999
' ������    R = 1  1-999 �����
' ��������  R = 2  1-999 ���������
' ��������� R = 3  1-999 ����������
Function f$(n, r)
Dim s$, ed, I, des, sot
    s$ = ""              ' ��������� ���������� �������� ����������
   If n = 0 Then GoTo Kon
' ���������� ���������� ������, ��������, �����
   ed = n Mod 10
   I = Int(n / 10)
   des = I Mod 10
   I = Int(I / 10)
   sot = I Mod 10
' ������������ ������ ��������
  If des = 1 Then
     Select Case ed
            Case 1
                 s$ = " ����������"
            Case 2
                 s$ = " ����������"
            Case 3
                 s$ = " ����������"
            Case 4
                 s$ = " ������������"
            Case 5
                 s$ = " �'���������"
            Case 6
                 s$ = " ������������"
            Case 7
                 s$ = " ���������"
            Case 8
                 s$ = " ����������"
            Case 9
                 s$ = " ���'���������"
            Case Else
                 s$ = " ������"
     End Select
            Else
         Select Case ed
            Case 1
                 s$ = " ����"
                 If r > 1 Then s$ = " ����"
            Case 2
                 s$ = " ��"
                 If r > 1 Then s$ = " ���"
            Case 3
                 s$ = " ���"
            Case 4
                 s$ = " ������"
            Case 5
                 s$ = " �'���"
            Case 6
                 s$ = " �����"
            Case 7
                 s$ = " ��"
            Case 8
                 s$ = " ���"
            Case 9
                 s$ = " ���'���"
            Case Else
                 s$ = ""
     End Select
   End If
     Select Case des
            Case 2
                 s$ = " ��������" + s$
            Case 3
                 s$ = " ��������" + s$
            Case 4
                 s$ = " �����" + s$
            Case 5
                 s$ = " �'�������" + s$
            Case 6
                 s$ = " ���������" + s$
            Case 7
                 s$ = " �������" + s$
            Case 8
                 s$ = " ��������" + s$
            Case 9
                 s$ = " ���'������" + s$
            Case Else
                 s$ = s$
     End Select
         Select Case sot
            Case 1
                 s$ = " ���" + s$
            Case 2
                 s$ = " ����" + s$
            Case 3
                 s$ = " ������" + s$
            Case 4
                 s$ = " ���������" + s$
            Case 5
                 s$ = " �'�����" + s$
            Case 6
                 s$ = " �������" + s$
            Case 7
                 s$ = " �����" + s$
            Case 8
                 s$ = " ������" + s$
            Case 9
                 s$ = " ���'�����" + s$
            Case Else
                 s$ = s$
     End Select
' ������������ ������������ �� ������� ������-��������
     If des = 1 Then       ' ������������ ��� ��������� 11-19
                 Select Case r
                        Case 0
                             s$ = s$     '+ " �������"
                        Case 1
                             s$ = s$ + " �����"
                        Case 2
                             s$ = s$ + " �������"
                        Case 3
                             s$ = s$ + " �������"
                        Case Else
                             s$ = s$
                  End Select
' ������������ �� ��������� �����
            Else
                Select Case ed
                       Case 1  ' ����
                          Select Case r
                                 Case 0
                                      s$ = s$     '+ " ������"
                                 Case 1
                                      s$ = s$ + " ������"
                                 Case 2
                                      s$ = s$ + " ������"
                                 Case 3
                                      s$ = s$ + " ������"
                                 Case Else
                                      s$ = s$
                          End Select
                       Case 2 To 4 ' ��� - ������
                          Select Case r
                                 Case 0
                                      s$ = s$    ' + " �����"
                                 Case 1
                                      s$ = s$ + " ������"
                                 Case 2
                                      s$ = s$ + " �������"
                                 Case 3
                                      s$ = s$ + " �������"
                                 Case Else
                                      s$ = s$
                           End Select
                       Case Else  ' ���������
                          Select Case r
                                 Case 0
                                      s$ = s$     ' + " �������"
                                 Case 1
                                      s$ = s$ + " �����"
                                 Case 2
                                      s$ = s$ + " �������"
                                 Case 3
                                      s$ = s$ + " �������"
                                 Case Else
                                      s$ = s$
                           End Select
                End Select
      End If
Kon:
     f$ = s$
End Function
'
' �������������� ���� � ������� ������ ��������
' ������ ������� 25.08.96 (��������� ���������)
'
Sub ������_�_�(n, s$)
Dim naim(1 To 12) As String
naim(1) = " ������ "
naim(2) = " ������� "
naim(3) = " ����� "
naim(4) = " ������ "
naim(5) = " ��� "
naim(6) = " ���� "
naim(7) = " ���� "
naim(8) = " ������� "
naim(9) = " �������� "
naim(10) = " ������� "
naim(11) = " ������ "
naim(12) = " ������� "
I = Month(n)
s$ = " "" " + str(Day(n)) + " "" " + naim(I) + str(Year(n)) + " �."
End Sub
'
' �������������� ����� � ����� ��������
' ������ ������� 25.08.96 (��������� ���������)
'
Sub ������_��(n, s$)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' �������
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' ���������
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' ��������
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' ������
   r = r - ts * 1000
   eee = Int(r)                  ' �� ������
   s$ = FR$(MLRD, 3) + FR$(mln, 2) + FR$(ts, 1) + FR$(eee, 0) + " �����. " + Kop + " ���."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "���� ���. " + Kop + " ���."
End If
End Sub
Function ���_�$(n)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' �������
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' ���������
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' ��������
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' ������
   r = r - ts * 1000
   eee = Int(r)                  ' �� ������
   s$ = FR$(MLRD, 3) + FR$(mln, 2) + FR$(ts, 1) + FR$(eee, 0) + " ���. " + Kop + " ���."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "���� ���. " + Kop + " ���."
End If
���_�$ = s$
End Function
Function ���$(n)
n = Application.Round(n, 2)
Kop = (n - Int(n)) * 100  ' �������
Kop = Format(Kop, "#00")
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' ���������
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' ��������
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' ������
   r = r - ts * 1000
   eee = Int(r)                  ' �� ������
   s$ = FR$(MLRD, 3) + FR$(mln, 2) + FR$(ts, 1) + FR$(eee, 0) + " ���. " + Kop + " ���."
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "���� ���. " + Kop + " ���."
End If
���$ = s$

End Function
'
' �������������� ����� �� 999 � ����� ��������
' ������ ������� 25.08.96 (��������� ���������)
' �� ������ R = 0  0-999
' ������    R = 1  1-999 �����
' ��������  R = 2  1-999 ���������
' ��������� R = 3  1-999 ����������
Function FR$(n, r)
   s$ = ""              ' ��������� ���������� �������� ����������
   If n = 0 Then GoTo Kon
' ���������� ���������� ������, ��������, �����
   ed = n Mod 10
   I = Int(n / 10)
   des = I Mod 10
   I = Int(I / 10)
   sot = I Mod 10
' ������������ ������ ��������
  If des = 1 Then
     Select Case ed
            Case 1
                 s$ = " �����������"
            Case 2
                 s$ = " ����������"
            Case 3
                 s$ = " ����������"
            Case 4
                 s$ = " ������������"
            Case 5
                 s$ = " ����������"
            Case 6
                 s$ = " �����������"
            Case 7
                 s$ = " ����������"
            Case 8
                 s$ = " ������������"
            Case 9
                 s$ = " ������������"
            Case Else
                 s$ = " ������"
     End Select
            Else
         Select Case ed
            Case 1
                 s$ = " ����"
                 If r > 1 Then s$ = " ����"
            Case 2
                 s$ = " ���"
                 If r > 1 Then s$ = " ���"
            Case 3
                 s$ = " ���"
            Case 4
                 s$ = " ������"
            Case 5
                 s$ = " ����"
            Case 6
                 s$ = " �����"
            Case 7
                 s$ = " ����"
            Case 8
                 s$ = " ������"
            Case 9
                 s$ = " ������"
            Case Else
                 s$ = ""
     End Select
   End If
     Select Case des
            Case 2
                 s$ = " ��������" + s$
            Case 3
                 s$ = " ��������" + s$
            Case 4
                 s$ = " �����" + s$
            Case 5
                 s$ = " ���������" + s$
            Case 6
                 s$ = " ����������" + s$
            Case 7
                 s$ = " ���������" + s$
            Case 8
                 s$ = " �����������" + s$
            Case 9
                 s$ = " ���������" + s$
            Case Else
                 s$ = s$
     End Select
         Select Case sot
            Case 1
                 s$ = " ���" + s$
            Case 2
                 s$ = " ������" + s$
            Case 3
                 s$ = " ������" + s$
            Case 4
                 s$ = " ���������" + s$
            Case 5
                 s$ = " �������" + s$
            Case 6
                 s$ = " ��������" + s$
            Case 7
                 s$ = " �������" + s$
            Case 8
                 s$ = " ���������" + s$
            Case 9
                 s$ = " ���������" + s$
            Case Else
                 s$ = s$
     End Select
' ������������ ������������ �� ������� ������-��������
     If des = 1 Then       ' ������������ ��� ��������� 11-19
                 Select Case r
                        Case 0
                             s$ = s$     '+ " �������"
                        Case 1
                             s$ = s$ + " �����"
                        Case 2
                             s$ = s$ + " ���������"
                        Case 3
                             s$ = s$ + " ���������"
                        Case Else
                             s$ = s$
                  End Select
' ������������ �� �������� �����
            Else
                Select Case ed
                       Case 1  ' ����
                          Select Case r
                                 Case 0
                                      s$ = s$     '+ " ������"
                                 Case 1
                                      s$ = s$ + " ������"
                                 Case 2
                                      s$ = s$ + " �������"
                                 Case 3
                                      s$ = s$ + " �������"
                                 Case Else
                                      s$ = s$
                          End Select
                       Case 2 To 4 ' ��� - ������
                          Select Case r
                                 Case 0
                                      s$ = s$    ' + " �����"
                                 Case 1
                                      s$ = s$ + " ������"
                                 Case 2
                                      s$ = s$ + " ��������"
                                 Case 3
                                      s$ = s$ + " ��������"
                                 Case Else
                                      s$ = s$
                           End Select
                       Case Else  ' ���������
                          Select Case r
                                 Case 0
                                      s$ = s$     ' + " �������"
                                 Case 1
                                      s$ = s$ + " �����"
                                 Case 2
                                      s$ = s$ + " ���������"
                                 Case 3
                                      s$ = s$ + " ���������"
                                 Case Else
                                      s$ = s$
                           End Select
                End Select
      End If
Kon:
     FR$ = s$
End Function
'
' �������������� ����� � ����� �������� � ��������
' ������ ������� 25.08.96 (��������� ���������)
'
Sub ������_���(n, s$, pr1$, pr2$)
�$ = str(n)
m = InStr(�$, ".")
dr$ = "0"
If m > 0 Then
   dr$ = Mid(�$, m + 1)
End If
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' ���������
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' ��������
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' ������
   r = r - ts * 1000
   eee = Int(r)                  ' �� ������
   s$ = FRD$(MLRD, 3) + FRD$(mln, 2) + FRD$(ts, 1) + FRD$(eee, 0) + pr1$ + dr$ + pr2$
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "���� " + pr1$ + dr$ + pr2$
End If
End Sub
Function �����$(n, pr1$, pr2$)
�$ = str(n)
m = InStr(�$, ".")
dr$ = "0"
If m > 0 Then
   dr$ = Mid(�$, m + 1)
End If
If n >= 1 Then
   MLRD = Int(n / 1000000000)    ' ���������
   r = n - MLRD * 1000000000
   mln = Int(r / 1000000)        ' ��������
   r = r - mln * 1000000
   ts = Int(r / 1000)            ' ������
   r = r - ts * 1000
   eee = Int(r)                  ' �� ������
   s$ = FRD$(MLRD, 3) + FRD$(mln, 2) + FRD$(ts, 1) + FRD$(eee, 0) + pr1$ + dr$ + pr2$
   s$ = UCase(Mid(s$, 2, 1)) + Mid(s$, 3, Len(s$) - 2)
          Else
   s$ = "���� " + pr1$ + dr$ + pr2$
End If
�����$ = s$
End Function
'
' �������������� ����� �� 999 � ����� ��������
' ������ ������� 25.08.96 (��������� ���������)
' �� ������ R = 0  0-999
' ������    R = 1  1-999 �����
' ��������  R = 2  1-999 ���������
' ��������� R = 3  1-999 ����������
Function FRD$(n, r)
   s$ = ""              ' ��������� ���������� �������� ����������
   If n = 0 Then GoTo Kon
' ���������� ���������� ������, ��������, �����
   ed = n Mod 10
   I = Int(n / 10)
   des = I Mod 10
   I = Int(I / 10)
   sot = I Mod 10
' ������������ ������ ��������
  If des = 1 Then
     Select Case ed
            Case 1
                 s$ = " �����������"
            Case 2
                 s$ = " ����������"
            Case 3
                 s$ = " ����������"
            Case 4
                 s$ = " ������������"
            Case 5
                 s$ = " ����������"
            Case 6
                 s$ = " �����������"
            Case 7
                 s$ = " ����������"
            Case 8
                 s$ = " ������������"
            Case 9
                 s$ = " ������������"
            Case Else
                 s$ = " ������"
     End Select
            Else
         Select Case ed
            Case 1
                 s$ = " ����"
            Case 2
                 s$ = " ���"
            Case 3
                 s$ = " ���"
            Case 4
                 s$ = " ������"
            Case 5
                 s$ = " ����"
            Case 6
                 s$ = " �����"
            Case 7
                 s$ = " ����"
            Case 8
                 s$ = " ������"
            Case 9
                 s$ = " ������"
            Case Else
                 s$ = ""
     End Select
   End If
     Select Case des
            Case 2
                 s$ = " ��������" + s$
            Case 3
                 s$ = " ��������" + s$
            Case 4
                 s$ = " �����" + s$
            Case 5
                 s$ = " ���������" + s$
            Case 6
                 s$ = " ����������" + s$
            Case 7
                 s$ = " ���������" + s$
            Case 8
                 s$ = " �����������" + s$
            Case 9
                 s$ = " ���������" + s$
            Case Else
                 s$ = s$
     End Select
         Select Case sot
            Case 1
                 s$ = " ���" + s$
            Case 2
                 s$ = " ������" + s$
            Case 3
                 s$ = " ������" + s$
            Case 4
                 s$ = " ���������" + s$
            Case 5
                 s$ = " �������" + s$
            Case 6
                 s$ = " ��������" + s$
            Case 7
                 s$ = " �������" + s$
            Case 8
                 s$ = " ���������" + s$
            Case 9
                 s$ = " ���������" + s$
            Case Else
                 s$ = s$
     End Select
' ������������ ������������ �� ������� ������-��������
     If des = 1 Then       ' ������������ ��� ��������� 11-19
                 Select Case r
                        Case 0
                             s$ = s$     '+ " �������"
                        Case 1
                             s$ = s$ + " �����"
                        Case 2
                             s$ = s$ + " ���������"
                        Case 3
                             s$ = s$ + " ���������"
                        Case Else
                             s$ = s$
                  End Select
' ������������ �� �������� �����
            Else
                Select Case ed
                       Case 1  ' ����
                          Select Case r
                                 Case 0
                                      s$ = s$     '+ " ������"
                                 Case 1
                                      s$ = s$ + " ������"
                                 Case 2
                                      s$ = s$ + " �������"
                                 Case 3
                                      s$ = s$ + " �������"
                                 Case Else
                                      s$ = s$
                          End Select
                       Case 2 To 4 ' ��� - ������
                          Select Case r
                                 Case 0
                                      s$ = s$    ' + " �����"
                                 Case 1
                                      s$ = s$ + " ������"
                                 Case 2
                                      s$ = s$ + " ��������"
                                 Case 3
                                      s$ = s$ + " ��������"
                                 Case Else
                                      s$ = s$
                           End Select
                       Case Else  ' ���������
                          Select Case r
                                 Case 0
                                      s$ = s$     ' + " �������"
                                 Case 1
                                      s$ = s$ + " �����"
                                 Case 2
                                      s$ = s$ + " ���������"
                                 Case 3
                                      s$ = s$ + " ���������"
                                 Case Else
                                      s$ = s$
                           End Select
                End Select
      End If
Kon:
     FRD$ = s$
End Function

Function �����������(����)
Dim naim(1 To 12) As String
naim(1) = "ѳ����"
naim(2) = "�����"
naim(3) = "��������"
naim(4) = "������"
naim(5) = "�������"
naim(6) = "�������"
naim(7) = "������"
naim(8) = "�������"
naim(9) = "��������"
naim(10) = "�������"
naim(11) = "��������"
naim(12) = "�������"

I = Month(����)
����������� = naim(I)
End Function


'DV-7kiR8DMzxy0H2fnITKABIR # Do not ove this line; required for DocVerse merge.
