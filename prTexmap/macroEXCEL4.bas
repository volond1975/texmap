Attribute VB_Name = "macroEXCEL4"
'GET.FORMULA () XLM
Function getFormula(r As Range)
getFormula = r.Formula
End Function
Function getFormulaArray(r As Range)
getFormulaArray = r.FormulaArray
End Function
Function getFormulaR1C1(r As Range)
getFormulaR1C1 = r.FormulaR1C1
End Function
Function getFormulaLocal(r As Range)
getFormulaLocal = r.FormulaLocal
End Function
Function getFormulaR1C1Local(r As Range)
getFormulaR1C1Local = r.FormulaR1C1Local
End Function
Function getNamedRangeByName(name As String) As name
Dim twb As Workbook
Set twb = ThisWorkbook
Set getNamedRangeByName = twb.Names(name)
End Function

Function getReferToNamedRangeByName(name As String) As String

Set rname = getNamedRangeByName(name)
getReferToNamedRangeByName = rname.RefersTo
End Function
Function getNamedRangeByAddress(addr As String)
Dim twb As Workbook
Dim r As Range
Set r = Range(addr)
getNamedRangeByAddress = r.Address(External:=True)
End Function
Function Intersect_Name(Optional addr)
    Dim nn As name, rRange As Range
    Set r = Range(addr)
    For Each nn In ActiveWorkbook.Names
        On Error Resume Next: Set rRange = nn.RefersToRange
        On Error GoTo 0
        If Not rRange Is Nothing Then
        
            If rRange.parent.name = r.parent.name Then
            
                If Not Intersect(rRange, r) Is Nothing Then
                Intersect_Name = nn.name
                Exit Function
                End If
            End If
            Set rRange = Nothing
        End If
    Next
'   // MsgBox "?????????? ?????? ?? ?????? ?? ? ???? ??????????? ????????", vbInformation, "Hay_from_The_Prist"
End Function
Sub addNamedFormula()

Application.Names.Add ActiveCell.Offset(0, -7).value, ActiveCell.value
'MsgBox Evaluate("x-2")
'Application.Names.Add name = ActiveCell.Offset(0, -7).value, _
                      RefersTo:="""" & ActiveCell.value & """"
End Sub

Function CellAddress()
CellAddress = Application.ThisCell.Address
End Function
Function CellParentName()
CellParentName = Application.ThisCell.parent.name

End Function
Function REGPLACE(myRangeValue As String, matchPattern As String, outputPattern As String) As Variant
    Dim regex As New VBScript_RegExp_55.RegExp
    Dim strInput As String

    strInput = myRangeValue

    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = matchPattern
    End With

    REGPLACE = regex.Replace(strInput, outputPattern)

End Function
Function REGPLACEFORMULA(myRangeFormula As String, matchPattern As String, outputPattern As String) As Variant
    Dim regex As New VBScript_RegExp_55.RegExp
    Dim strInput As String

 strInput = myRangeFormula

    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = matchPattern
    End With

    REGPLACEFORMULA = regex.Replace(strInput, outputPattern)

End Function
'http://qaru.site/questions/19369/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops


Function regex(strInput As String, matchPattern As String, Optional ByVal outputPattern As String = "$0") As Variant
    Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
    Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
    Dim replaceNumber As Integer

    With inputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = matchPattern
    End With
    With outputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "\$(\d+)"
    End With
    With outReplaceRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With

    Set inputMatches = inputRegexObj.Execute(strInput)
    If inputMatches.Count = 0 Then
        regex = False
    Else
        Set replaceMatches = outputRegexObj.Execute(outputPattern)
        For Each replaceMatch In replaceMatches
            replaceNumber = replaceMatch.SubMatches(0)
            outReplaceRegexObj.pattern = "\$" & replaceNumber

            If replaceNumber = 0 Then
                outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).value)
            Else
                If replaceNumber > inputMatches(0).SubMatches.Count Then
                    'regex = "A to high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
                    regex = CVErr(xlErrValue)
                    Exit Function
                Else
                    outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
                End If
            End If
        Next
        regex = outputPattern
    End If
End Function
Public Function CellReflist(Optional r As Range)  ' single cell
Dim result As Object: Dim testExpression As String: Dim objRegEx As Object
If r Is Nothing Then Set r = ActiveCell ' Cells(1, 2)  ' INPUT THE CELL HERE , e.g.    RANGE("A1")
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.IgnoreCase = True: objRegEx.Global = True: objRegEx.pattern = """.*?"""    ' remove expressions
testExpression = CStr(r.Formula)
testExpression = objRegEx.Replace(testExpression, "")
'objRegEx.Pattern = "(([A-Z])+(\d)+)"  'grab the address

objRegEx.pattern = "(['\[].*?['!])?([[A-Z0-9_]+[!])?(\$?[A-Z]+\$?(\d)+(:\$?[A-Z]+\$?(\d)+)?|\$?[A-Z]+:\$?[A-Z]+|(\$?[A-Z]+\$?(\d)+))"
If objRegEx.Test(testExpression) Then
    Set result = objRegEx.Execute(testExpression)
    If result.Count > 0 Then CellReflist = result(0).value
    If result.Count > 1 Then
        For i = 1 To result.Count - 1 'Each Match In result
            dbl = False ' poistetaan tuplaesiintymiset
            For j = 0 To i - 1
                If result(i).value = result(j).value Then dbl = True
            Next j
            If Not dbl Then CellReflist = CellReflist & "," & result(i).value 'Match.Value
        Next i 'Match
    End If
End If
End Function

Sub ReturnFormulaReferences()

    Dim objRegExp As New VBScript_RegExp_55.RegExp
    Dim objCell As Range
    Dim objStringMatches As Object
    Dim objReferenceMatches As Object
    Dim objMatch As Object
    Dim intReferenceCount As Integer
    Dim intIndex As Integer
    Dim booIsReference As Boolean
    Dim objName As name
    Dim booNameFound As Boolean

    With objRegExp
        .MultiLine = True
        .Global = True
        .IgnoreCase = True
    End With

    For Each objCell In Selection.Cells
        If Left(objCell.Formula, 1) = "=" Then

            objRegExp.pattern = "\"".*\"""
            Set objStringMatches = objRegExp.Execute(objCell.Formula)

            objRegExp.pattern = "(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
            & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)?" _
            & "(\$?[a-z]{1,3}\$?[0-9]{1,7}(\:\$?[a-z]{1,3}\$?[0-9]{1,7})?" _
            & "|\$[a-z]{1,3}\:\$[a-z]{1,3}" _
            & "|[a-z]{1,3}\:[a-z]{1,3}" _
            & "|\$[0-9]{1,7}\:\$[0-9]{1,7}" _
            & "|[0-9]{1,7}\:[0-9]{1,7}" _
            & "|[a-z_\\][a-z0-9_\.]{0,254})"
            Set objReferenceMatches = objRegExp.Execute(objCell.Formula)

            intReferenceCount = 0
            For Each objMatch In objReferenceMatches
                intReferenceCount = intReferenceCount + 1
            Next

            Debug.Print objCell.Formula
            For intIndex = intReferenceCount - 1 To 0 Step -1
                booIsReference = True
                For Each objMatch In objStringMatches
                    If objReferenceMatches(intIndex).FirstIndex > objMatch.FirstIndex _
                    And objReferenceMatches(intIndex).FirstIndex < objMatch.FirstIndex + objMatch.Length Then
                        booIsReference = False
                        Exit For
                    End If
                Next

                If booIsReference Then
                    objRegExp.pattern = "(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
                    & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)?" _
                    & "(\$?[a-z]{1,3}\$?[0-9]{1,7}(\:\$?[a-z]{1,3}\$?[0-9]{1,7})?" _
                    & "|\$[a-z]{1,3}\:\$[a-z]{1,3}" _
                    & "|[a-z]{1,3}\:[a-z]{1,3}" _
                    & "|\$[0-9]{1,7}\:\$[0-9]{1,7}" _
                    & "|[0-9]{1,7}\:[0-9]{1,7})"
                    If Not objRegExp.Test(objReferenceMatches(intIndex).value) Then 'reference is not A1
                        objRegExp.pattern = "^(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
                        & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)" _
                        & "[a-z_\\][a-z0-9_\.]{0,254}$"
                        If Not objRegExp.Test(objReferenceMatches(intIndex).value) Then 'name is not external
                            booNameFound = False
                            For Each objName In objCell.Worksheet.parent.Names
                                If objReferenceMatches(intIndex).value = objName.name Then
                                    booNameFound = True
                                    Exit For
                                End If
                            Next
                            If Not booNameFound Then
                                objRegExp.pattern = "^(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
                                & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)"
                                For Each objName In objCell.Worksheet.Names
                                    If objReferenceMatches(intIndex).value = objRegExp.Replace(objName.name, "") Then
                                        booNameFound = True
                                        Exit For
                                    End If
                                Next
                            End If
                            booIsReference = booNameFound
                        End If
                    End If
                End If

                If booIsReference Then
                    Debug.Print "  " & objReferenceMatches(intIndex).value _
                    & " (" & objReferenceMatches(intIndex).FirstIndex & ", " _
                    & objReferenceMatches(intIndex).Length & ")"
                End If
            Next intIndex
            Debug.Print

        End If
    Next

    Set objRegExp = Nothing
    Set objStringMatches = Nothing
    Set objReferenceMatches = Nothing
    Set objMatch = Nothing
    Set objCell = Nothing
    Set objName = Nothing

End Sub
 Sub replaceAllAC()
 Dim r As Range
 Dim f As String
 For Each r In Selection.Cells
 Call replaceAll(r)
 f = getFormula(r)
 Debug.Print f
r.Formula = REGPLACE(f, "1.215", "g  «")

 Next

 End Sub
 
Sub replaceAll(Optional r As Range)
If r Is Nothing Then Set r = ActiveCell
Dim rStr As String
Dim v As Variant
Dim strInput As String
Dim testExpression As String
Dim newz
Dim z As String
testExpression = CStr(r.Formula)
 rStr = CellReflist(r)
  v = Split(rStr, ",")
 For i = 0 To UBound(v)
 strInput = v(i)
 z = Val(regex(strInput, "\d+")) + 0
 leter = regex(strInput, "[A-Z]+")
 newz = z + 65
 Debug.Print v(i)
testExpression = REGPLACEFORMULA(testExpression, z, CStr(newz))
Debug.Print testExpression
Next i
r.Formula = testExpression
End Sub
Function BankersRound(num, precision)
BankersRound = Round(num, precision)
End Function
Function JOIN_Range(r As Range)
'????????? ????????? ????
  
Dim myArr As Variant
Dim Val As String
  v = r
   Dim rngRectangle As Range, rngRows As Range, rngColumns As Range
    Set rngRectangle = r 'Selection
'    С ќпредел€ет вертикальный вектор массива
    Set rngRows = rngRectangle.Resize(, 1)
'    С ќпредел€ет горизонтальный вектор массива
    Set rngColumns = rngRectangle.Resize(1)
'    rngRectangle = Evaluate("IF(ROW(" & rngRows.Address & "), _
        IF(COLUMN(" & rngColumns.Address &"),UPPER(" & rngRectangle.Address & ")))")
myArr = Application.Transpose(Application.Transpose(r))
   
 Val = Join(myArr, "_")
  
JOIN_Range = Val
  
End Function
Sub Join_example()
'????????? ????????? ????
  
Dim valArr(1 To 3)
Dim Val As String
  
 valArr(1) = "???"
 valArr(2) = "MS"
 valArr(3) = "Excel"
   
 Val = Join(valArr, "_")
  
MsgBox Val
  
End Sub
Sub RectangularUpper()
'    С ѕреобразует все €чейки в выделенном диапазоне в верхний регистр
    Dim rngRectangle As Range, rngRows As Range, rngColumns As Range
    Set rngRectangle = Selection
'    С ќпредел€ет вертикальный вектор массива
    Set rngRows = rngRectangle.Resize(, 1)
'    С ќпредел€ет горизонтальный вектор массива
    Set rngColumns = rngRectangle.Resize(1)
'    rngRectangle = Evaluate("IF(ROW(" & rngRows.Address & "), _
        IF(COLUMN(" & rngColumns.Address &"),UPPER(" & rngRectangle.Address & ")))")
End Sub

Sub AddNamesOffsetRange()
 Dim r As Range
 Dim NameMe As String
 For Each r In Selection.Cells
 If r <> "" Then
' r.value
 NameMe = r.value
 ActiveWorkbook.Names.Add name = NameMe, _
                      RefersTo:="'" & r.Offset(0, r.Offset(0, 1)).parent.name & "'!" & r.Offset(0, r.Offset(0, 1)).Address
 
 End If

 Next

 End Sub
