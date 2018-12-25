Attribute VB_Name = "mFuncValOfFormula"
'---------------------------------------------------------------------------------------
' Module    : mFuncValOfFormula
' DateTime  : 13.09.2012 10:19
' Author    : The_Prist(ўербаков ƒмитрий)
'             ѕрофессиональна€ разработка приложений дл€ MS Office любой сложности
'             ѕроведение тренингов по MS Excel
'             WebMoney - R298726502453; яндекс.ƒеньги - 41001332272872
'             http://www.excel-vba.ru
' Purpose   : http://www.excel-vba.ru/chto-umeet-excel/otobrazit-v-formulax-vmesto-ssylok-na-yachejki-znacheniya-yacheek/
'             ѕроцедура:
'             1) копирует формулу одной €чейки в другую;
'             2) преобразует скопированную формулу другой €чейки в текст таким образом,
'                что вместо ссылок на €чейки проставл€ютс€ их значени€
'---------------------------------------------------------------------------------------
Option Explicit
Dim wsParentSheet As Worksheet, bCell As Boolean
'#mFuncValOfFormula.Get_Val_of_Formula
Sub Get_Val_of_Formula()
    Dim rRange As Range, rCell As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    On Error Resume Next
    'определ€ем диапазон €чеек с формулами
    If Selection.Count = 1 Then
        If ActiveCell.HasFormula Then Set rRange = ActiveCell
    Else
        Set rRange = Selection.SpecialCells(xlFormulas)
    End If
    If rRange Is Nothing Then MsgBox "¬ выделенном диапазоне отсутствуют €чейки с формулами", vbCritical, "Get_Val_of_Formula": Exit Sub
    bCell = (MsgBox("ќтобразить переведенную формулу в примечании к €чейке?", vbYesNo, "ќпределение метода отображени€") = vbNo)
    
    'запоминаем лист с формулами - это пригодитс€ дл€ записи значений
    Set wsParentSheet = ActiveSheet

    Application.ScreenUpdating = False: Application.EnableEvents = False
    ActiveSheet.copy
    'выставл€ем дл€ новой книги "“очность как на экране" - дл€ вставки значений в таком же виде, в каком они видны пользователю на листе
    ActiveWorkbook.PrecisionAsDisplayed = True
    'определ€ем значени€ ссылок на €чейки дл€ каждой формулы
    For Each rCell In rRange
        Call Val_Of_Formula(Range(rCell.Address))
    Next rCell
    ActiveWorkbook.Close False
    wsParentSheet.Activate
    Application.ScreenUpdating = True: Application.EnableEvents = True
End Sub

Function Val_Of_Formula(ByVal rCell As Range)
    Const sArrSepRows = ":", sArrSepCols = ";", sArgSep = ";"
    Dim sFormLocalStr As String, sTmpAddr As String, sExternalRng As String, sTmpStr As String, sRes As String
    Dim rSel As Object, oVal, objRegEx As Object, objMatshes As Object
    Dim avArr, avTmp, oCmnt
    Dim iRefMatch_Cnt As Long, lPresedCnt As Long, iSelCnt As Long, iR_Cnt As Long, iC_Cnt As Long, iLen As Long
    Dim iTmpLen As Long, iAddrBeginPos As Long, iAddrLen As Long
    Dim sTmpFormStr As String, sPattern As String
    Dim li As Long
'   //получаем формулы на локализованном €зыке
    sFormLocalStr = rCell.FormulaLocal
    
    ReDim avArr(10000, 4)
    Dim i As Long, ir As Long, ic As Long
    lPresedCnt = 1
    Set objRegEx = CreateObject("VBScript.Regexp")
    objRegEx.IgnoreCase = True
    objRegEx.MultiLine = True
    objRegEx.Global = True

    sPattern = "(('?\[[^\\\/\:\*\?\""\<\>\|]+?\]([^\:\\\/\?\*\[\]']{1,31}'!|" & _
        "[^\:\\\/\?\*\[\]'!\@\#\$\%\^\&\(\)\+\-\=\|\""\<\>\{\},`~\; ]{1,31}!)" & _
        "(" & "([\:]?\$?[a-z]{1,3}\$?\d{1,7}){1,2}(?=[\+\-\=\*\^\/\\;,\)\s])|" & _
        "\$?[a-z]{1,3}\:\$?[a-z]{1,3}(?=[\+\-\=\*\^\/\\;,\)\s])|(\$?\d{1,7}\:\$?\d{1,7})(?=[\+\-\=\*\^\/\\;,\)\s])))|(" & _
        "(('[^\:\\\/\?\*\[\]']{1,31}'|" & "[^\:\\\/\?\*\[\]'!\@\#\$\%\^\&\(\)\+\-\=\|\""\<\>\{\},`~\; ]{1,31})!)(([\:]?\$?[a-z]{1,3}\$?\d{1,7}){1,2}" & _
        "(?=[\+\-\=\*\^\/\\;,\)\s])|\$?[a-z]{1,3}\:\$?[a-z]{1,3}(?=[\+\-\=\*\^\/\\;,\)\s])" & _
        "|(\$?\d{1,7}\:\$?\d{1,7})(?=[\+\-\=\*\^\/\\;,\)\s])))|(" & _
        "([\:]?\$?[a-z]{1,3}\$?\d{1,7}){1,2}(?=[\+\-\=\*\^\/\\;,\)\s])|\$?[a-z]{1,3}\:\$?[a-z]{1,3}" & _
        "(?=[\+\-\=\*\^\/\\;,\)\s])|(\$?\d{1,7}\:\$?\d{1,7})(?=[\+\-\=\*\^\/\\;,\)\s])))"
    objRegEx.pattern = sPattern
    Set objMatshes = objRegEx.Execute(sFormLocalStr & "+")
    iRefMatch_Cnt = objMatshes.Count
'   //ѕеребираем все разбитые ссылки
    For i = iRefMatch_Cnt - 1 To 0 Step -1
        sTmpAddr = objMatshes.Item(i).value
        If IsRange(sTmpAddr) Then
            sExternalRng = Range(sTmpAddr).Address(, , , True)
            Set rSel = Range(sExternalRng)
            sTmpAddr = rSel.Address(, , , True)
            iSelCnt = rSel.Count
            If iSelCnt > 1 Then
                avTmp = rSel.value
                iR_Cnt = UBound(avTmp, 1)
                iC_Cnt = UBound(avTmp, 2)
                For ir = 1 To iR_Cnt
                    For ic = 1 To iC_Cnt
                        oVal = avTmp(ir, ic)
                        If (oVal = "") Then oVal = 0
                        If IsNumeric(oVal) = False Then
                            oVal = Chr(34) & oVal & Chr(34)
                        End If
                        avArr(lPresedCnt - 1, 0) = avArr(lPresedCnt - 1, 0) & oVal & sArrSepCols
                    Next ic
                    sTmpStr = avArr(lPresedCnt - 1, 0)
                    If Right(sTmpStr, Len(sArrSepCols)) = sArrSepCols Then
                        avArr(lPresedCnt - 1, 0) = Mid(sTmpStr, 1, Len(sTmpStr) - 1)
                    End If
                    sTmpStr = avArr(lPresedCnt - 1, 0)
                    If Right(sTmpStr, Len(sArrSepRows)) <> sArrSepRows Then
                        avArr(lPresedCnt - 1, 0) = sTmpStr & sArrSepRows
                    End If
                Next ir
                sTmpStr = avArr(lPresedCnt - 1, 0)
                If Right(sTmpStr, Len(sArrSepCols)) = sArrSepCols Or Right(sTmpStr, Len(sArrSepRows)) = sArrSepRows Then
                    sTmpStr = Mid(sTmpStr, 1, Len(sTmpStr) - 1)
                    avArr(lPresedCnt - 1, 0) = "{" & sTmpStr & "}"
                End If
            Else
                avArr(lPresedCnt - 1, 0) = rSel.value
'                //если значение €чейки текст(и это не массив) - обрамл€ем в кавычки
                If IsNumeric(avArr(lPresedCnt - 1, 0)) = False Then
                    avArr(lPresedCnt - 1, 0) = Chr(34) & avArr(lPresedCnt - 1, 0) & Chr(34)
                End If
            End If 'If iSelCnt > 1 Then
            avArr(lPresedCnt - 1, 1) = objMatshes(i)
        Else
            avArr(lPresedCnt - 1, 0) = "[«начение недоступно]"
            avArr(lPresedCnt - 1, 1) = objMatshes(i)
        End If
        avArr(lPresedCnt - 1, 2) = objMatshes(i).FirstIndex
        avArr(lPresedCnt - 1, 3) = objMatshes(i).Length
        lPresedCnt = lPresedCnt + 1
    Next i
    'если ссылки на другие €чейки есть
    If lPresedCnt > 1 Then
        For li = 0 To lPresedCnt - 1
            iLen = Len(sFormLocalStr)
            iAddrBeginPos = Val(avArr(li, 2))
            iAddrLen = Val(avArr(li, 3))
            If iAddrBeginPos + iAddrLen >= iLen Then
                sTmpFormStr = ""
            Else
                sTmpFormStr = Mid(sFormLocalStr, iAddrBeginPos + iAddrLen + 1, iLen - (iAddrBeginPos + iAddrLen))
            End If
            sFormLocalStr = Mid(sFormLocalStr, 1, iAddrBeginPos) & avArr(li, 0) & sTmpFormStr
        Next li
        sRes = "'" & sFormLocalStr
    Else
        sRes = "'" & sFormLocalStr & " [ссылок на другие €чейки нет]"
    End If
    'записываем значение формулы в €чейку или создаем дл€ неЄ примечание
    If bCell Then
        wsParentSheet.Range(rCell.Address).Offset(, 1).value = sRes
    Else
        Set oCmnt = wsParentSheet.Range(rCell.Address).Comment
        If Not oCmnt Is Nothing Then
            wsParentSheet.Range(rCell.Address).Comment.Delete
        End If
        Dim sRest
        sRest = Mid(sRes, 2) & vbCrLf
        sRest = sRest & sFormLocalStr
        wsParentSheet.Range(rCell.Address).AddComment Mid(sRes, 2)
        wsParentSheet.Range(rCell.Address).value = wsParentSheet.Range(rCell.Address).value
    End If
End Function
Function IsRange(s As String)
    Dim rr As Range
    On Error Resume Next
    Set rr = Range(s)
    IsRange = Not rr Is Nothing
End Function
