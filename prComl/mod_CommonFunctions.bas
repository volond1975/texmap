Attribute VB_Name = "mod_CommonFunctions"


 '---------------------------------------------------------------------------------------
 ' Module        : mod_CommonFunctions
 ' Автор     : EducatedFool  (Игорь)                    Дата: 26.03.2012
 ' Разработка макросов для Excel, Word, CorelDRAW. Быстро, профессионально, недорого.
 ' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
 ' Реквизиты для оплаты работы: http://ExcelVBA.ru/payments
 '---------------------------------------------------------------------------------------
 Option Compare Text
 Option Private Module

 Function Chars(ByVal txt As String) As Variant
     On Error Resume Next: ReDim Arr(0 To Len(txt) - 1)
     For I = LBound(Arr) To UBound(Arr): Arr(I) = Mid(txt, I + 1, 1): Next I
     If Err Then Chars = Array() Else Chars = Arr
 End Function

 Function SafeText(ByVal txt As String) As String
     For I = 1 To Len(txt)
         SafeText = SafeText & IIf(I = 1, "", "-") & AscW(Mid(txt, I, 1))
     Next I
 End Function

 Function RestoreText(ByVal txt As String) As String
     On Error Resume Next: Arr = Split(txt, "-")
     For I = LBound(Arr) To UBound(Arr): Arr(I) = ChrW(Val(Arr(I))): Next I
     RestoreText = Join(Arr, "")
 End Function

 Function TemplatesInfo(ByVal files As Collection)
     For Each Item In files
         TemplatesInfo = TemplatesInfo & ";" & TemplateType(Item)
     Next
     TemplatesInfo = Left(Mid(TemplatesInfo, 2), 100)
 End Function

 Function ColunmNameByColumnNumber(ByVal col As Long) As String
     resA1 = Application.ConvertFormula("=r1c" & col, xlR1C1, xlA1)
     ColunmNameByColumnNumber = col & " «" & Split(resA1, "$")(1) & "»"
 End Function

 Function FindAll(SearchRange As Range, _
                  FindWhat As Variant, _
                  Optional LookIn As XlFindLookIn = xlValues, _
                  Optional LookAt As XlLookAt = xlWhole, _
                  Optional SearchOrder As XlSearchOrder = xlByRows, _
                  Optional MatchCase As Boolean = False, _
                  Optional BeginsWith As String = vbNullString, _
                  Optional EndsWith As String = vbNullString, _
                  Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Range
'
              
     Dim FoundCell As Range, FirstFound As Range, LastCell As Range, rngResultRange As Range
     Dim XLookAt As XlLookAt, Include As Boolean, CompMode As VbCompareMethod
     Dim Area As Range, MaxRow As Long, MaxCol As Long, BeginB As Boolean, EndB As Boolean

     CompMode = BeginEndCompare
     XLookAt = LookAt: If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then XLookAt = xlPart

    For Each Area In SearchRange.Areas
         With Area
             If .Cells(.Cells.Count).Row > MaxRow Then MaxRow = .Cells(.Cells.Count).Row
             If .Cells(.Cells.Count).Column > MaxCol Then MaxCol = .Cells(.Cells.Count).Column
         End With
     Next Area
     Set LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)
     Set FoundCell = SearchRange.Find(What:=FindWhat, after:=LastCell, _
                                      LookIn:=LookIn, LookAt:=XLookAt, _
                                      SearchOrder:=SearchOrder, MatchCase:=MatchCase)

     If Not FoundCell Is Nothing Then
         Set FirstFound = FoundCell
         Do Until False    ' Loop forever. We'll "Exit Do" when necessary.
            Include = False
             If BeginsWith = vbNullString And EndsWith = vbNullString Then
                 Include = True
             Else
                 If BeginsWith <> vbNullString Then
                     If StrComp(Left(FoundCell.Text, Len(BeginsWith)), _
                                BeginsWith, BeginEndCompare) = 0 Then Include = True
                 End If
                 If EndsWith <> vbNullString Then
                     If StrComp(Right(FoundCell.Text, Len(EndsWith)), _
                                EndsWith, BeginEndCompare) = 0 Then Include = True
                 End If
             End If
             If Include = True Then
                 If rngResultRange Is Nothing Then
                     Set rngResultRange = FoundCell
                 Else
                     Set rngResultRange = Application.Union(rngResultRange, FoundCell)
                 End If
             End If
             Set FoundCell = SearchRange.FindNext(after:=FoundCell)
             If (FoundCell Is Nothing) Then Exit Do
             If (FoundCell.address = FirstFound.address) Then Exit Do
         Loop
     End If
     Set FindAll = rngResultRange
 End Function
