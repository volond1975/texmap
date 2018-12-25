Attribute VB_Name = "nerv_SetFormPosition"

 #If Win64 Then
     Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
     Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
     Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As Rect) As Long
     Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
 #Else
     Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
     Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
     Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
     Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
 #End If

 Private Type Rect: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type

 '=========================================================
 ' Author: nerv            | E-mail: nerv-net@yandex.ru
 ' Last Update: 02/12/2011 | ίνδεκρ.Δενόγθ: 41001156540584
 '=========================================================

 Public Sub SetFormPosition(ByRef uf As Object, ByRef r As Range)
 Dim RFP As Object, hWnd, Frm As Rect, X As Integer, Y As Single, xn As Integer, i As Single, h As Long, w As Long
 On Error Resume Next
 w = GetSystemMetrics(16&)
 h = GetSystemMetrics(17&)
 i = 1 / (w / h)
 With ActiveWindow
     Do While X < w And Y < h
         X = X + 1: Y = Y + i
         Set RFP = .RangeFromPoint(X, Y)
         If Not RFP Is Nothing Then
             If TypeOf RFP Is Range Then
                 i = 1: Y = CInt(Y)
                 While RFP.Column <> r.Column
                     Set RFP = .RangeFromPoint(X, Y): X = X + i
                     If X <= 0 Then Exit Sub
                     If X >= w Then i = -1
                 Wend
                 i = 1: If RFP.Row = r.Row Then Y = Y + 1
                 While RFP.Row <> r.Row
                     Set RFP = .RangeFromPoint(X, Y): Y = Y + i
                     If Y <= 0 Then Exit Sub
                     If Y >= h Then i = -1
                 Wend
                 Do
                     If RFP.Column <> r.Column Or RFP Is Nothing Then Exit Do
                     Set RFP = .RangeFromPoint(X, Y): X = X - 1
                 Loop
                 X = X + 3: xn = X: Set RFP = .RangeFromPoint(X, Y)
                 With r.MergeArea: i = .Column + .Columns.Count - 1: End With
                 Do
                     If RFP.Column > i Or RFP Is Nothing Then Exit Do
                     Set RFP = .RangeFromPoint(xn, Y): xn = xn + 1
                 Loop
                 Exit Do
             End If
         End If
     Loop
 End With
 hWnd = FindWindow("ThunderDframe", uf.Caption)
 GetWindowRect hWnd, Frm: Frm.Left = Frm.Right - Frm.Left
 If Frm.Left + xn < w Then X = xn Else X = X - Frm.Left - 3: If X <= 0 Then Exit Sub
 If Frm.Bottom - Frm.Top + Y >= h Then Y = Y - Frm.Bottom - Frm.Top - 1
 SetWindowPos hWnd, 0&, X, Y, 0&, 0&, 1&
 End Sub
