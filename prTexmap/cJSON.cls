VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJSON"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Base 0

'Version 1.0

'Tools > References... > Microsoft Scripting Runtime must be checked

'members
Private Log_ As String
Private PositionInJSONString_ As Long
Private CharacterAtPositionInJSONString_ As String
Private LengthJSONString_ As Long 'VBA: limit of the code
Private DecimalSign_ As String
Private HowToConvertJSONNumberToVBA_ As eHowToConvertJSONNumberToVBA
Private HowToConformToJSONGrammar_ As eHowToConformToJSONGrammar
Private HowToConformToJSONNull_ As eHowToConformToJSONNull
Private CustomStringNull_ As String

'emumerations
Public Enum eHowToConvertJSONNumberToVBA
    TryLongTryDoubleElseUseString
    TryLongTryDoubleElseError
    TryCurrencyTryDoubleElseError
    TryLongElseError
    TryCurrencyElseError
    TryDoubleElseError
    UseLong
    UseCurrency
    UseDouble
    UseString
End Enum
Public Enum eHowToConformToJSONGrammar
    BeStrict 'Source: The JSON Data Interchange Format, Standard ECMA-404, 1st Edition / October 2013
    BeForgiving
End Enum
Public Enum eHowToConformToJSONNull
    UseNull
    UseNullString
    UseEmptyString
    UseCustomString
    UseZero
End Enum
Public Enum eHowToReset
    Full
    KeepHowToSettings
End Enum
Private Enum eCharacterCompare
    Both
    LowerCase
    UpperCase
End Enum
'methods
Public Sub Class_Initialize()
    Reset eHowToReset.Full
End Sub
Private Sub Reset(Optional ByVal HowToReset As eHowToReset = eHowToReset.KeepHowToSettings)
    If (HowToReset = eHowToReset.KeepHowToSettings) Then
        Log_ = vbNullString
        PositionInJSONString_ = 0
        CharacterAtPositionInJSONString_ = vbNullString
        LengthJSONString_ = 0
    Else
        Log_ = vbNullString
        PositionInJSONString_ = 0
        CharacterAtPositionInJSONString_ = vbNullString
        LengthJSONString_ = 0
        GuessDecimalSign
        HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.TryLongTryDoubleElseUseString
        HowToConformToJSONGrammar_ = eHowToConformToJSONGrammar.BeStrict
        HowToConformToJSONNull_ = eHowToConformToJSONNull.UseNull
    End If
End Sub
Public Sub SetHowToConformToJSONGrammar(ByVal HowToConformToJSONGrammar As eHowToConformToJSONGrammar)
    HowToConformToJSONGrammar_ = HowToConformToJSONGrammar
End Sub
Public Sub SetHowToConvertJSONNumberToVBA(ByVal HowToConvertJSONNumberToVBA As eHowToConvertJSONNumberToVBA)
    HowToConvertJSONNumberToVBA_ = HowToConvertJSONNumberToVBA
End Sub
Public Sub SetHowToConvertJSONNullToVBA(ByVal HowToConformToJSONNull As eHowToConformToJSONNull, Optional ByVal CustomStringNull As String = vbNullString)
    HowToConformToJSONNull_ = HowToConformToJSONNull
    If (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseCustomString) Then
        CustomStringNull_ = CustomStringNull
    End If
End Sub
Public Function ShowWhyNotOk() As String
    If (Not IsOk()) Then
        ShowWhyNotOk = Log_
    End If
End Function
Public Function IsOk() As Boolean
    IsOk = (Len(Log_) = 0)
End Function
Public Function Deserialize(ByRef JSONText As String) As Dictionary
    On Error GoTo DeserializeError
    'ready?
    Dim D As Dictionary
    Set D = New Dictionary
    'set?
    Set Deserialize = Nothing
    'go!
    Reset eHowToReset.KeepHowToSettings
    LengthJSONString_ = Len(JSONText) 'VBA: limit of the code
    If (LengthJSONString_ > 0) Then
        ChangePositionInJSONStringForward JSONText, 1, True
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""JSON only consists of white space.""}"
            Exit Function
        End If
        '[{]"Name":Value... or [[]Value,...
        If (AscW(CharacterAtPositionInJSONString_) = 123) Then '{
            If (ExtractObject(JSONText, D)) Then
                Set Deserialize = D
            Else
                Set D = Nothing
                Exit Function
            End If
        ElseIf (AscW(CharacterAtPositionInJSONString_) = 91) Then '[
            If (ExtractArray(JSONText, D)) Then
                Set Deserialize = D
            Else
                Set D = Nothing
                Exit Function
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""First (non white space) character is incorrect. Expecting { or [."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Set D = Nothing
            Exit Function
        End If
        ChangePositionInJSONStringForward JSONText, 1, True
        If (PositionInJSONString_ <= LengthJSONString_) Then
            If (HowToConformToJSONGrammar_ = eHowToConformToJSONGrammar.BeStrict) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Unnecessary characters after closing JSON."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Set Deserialize = Nothing
                Exit Function
            End If
        End If
    Else
        AddToLog "{""Type"":""JSON grammar"",""Description"":""JSON cannot be empty.""}"
    End If
    'finish...
    Set D = Nothing
    Exit Function
DeserializeError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in Deserialize().""}"
    Set D = Nothing
End Function
Private Sub AddToLog(ByRef Description As String)
    Log_ = "{""Error"":" & Description & "}"
End Sub
Private Function ExtractObject(ByRef JSONText As String, ByRef D As Dictionary) As Boolean
    On Error GoTo ExtractObjectError
    '...,[{]"Name":Value,...
    'ready?
    Dim IsFirstNameValuePair As Boolean
    IsFirstNameValuePair = True
    'set?
    ExtractObject = False 'default value
    'go!
    ChangePositionInJSONStringForward JSONText, 1, True
    If (PositionInJSONString_ > LengthJSONString_) Then
        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting object."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        Exit Function
    End If
    '...,{["]Key":Value,... or ...,{[}]
    Do While ((Not IsRightCurlyBracket(CharacterAtPositionInJSONString_)) And (PositionInJSONString_ <= LengthJSONString_))
        If (Not IsFirstNameValuePair) Then
            '...,{"Name1":Value1[,]"Name2":Value2,...
            If (IsComma(CharacterAtPositionInJSONString_)) Then
                ChangePositionInJSONStringForward JSONText, 1, True
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting object."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting , to separate name from value in object."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        End If
        If (PositionInJSONString_ <= LengthJSONString_) Then
            '...["]Key":Value,...
            If (ExtractNameValuePair(JSONText, D)) Then
                IsFirstNameValuePair = False
            Else
                Exit Function
            End If
            ChangePositionInJSONStringForward JSONText, 1, True
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting object."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
            '..."Name":Value[,]... or ..."Name":Value[}]...
        End If
    Loop
    If (PositionInJSONString_ > LengthJSONString_) Then
        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting object."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        Exit Function
    Else
        ExtractObject = True
    End If
    'finish...
    Exit Function
ExtractObjectError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractObject().""}"
End Function
Private Function ExtractArray(ByRef JSONText As String, ByRef D As Dictionary) As Boolean
    On Error GoTo ExtractArrayError
    '...,[[]Value,...
    'ready?
    Dim IndexOfValue As Long
    IndexOfValue = 0 'VBA: limit of the code
    Dim IsFirstValue As Boolean
    IsFirstValue = True
    'set?
    ExtractArray = False 'default value
    'go!
    ChangePositionInJSONStringForward JSONText, 1, True
    If (PositionInJSONString_ > LengthJSONString_) Then
        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting array."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        Exit Function
    End If
    '...[[V]alue,... or ...[[]]
    Do While ((Not IsRightSquareBracket(CharacterAtPositionInJSONString_)) And (PositionInJSONString_ <= LengthJSONString_))
        If (Not IsFirstValue) Then
            '...[Value1[,]Value2,...
            If (IsComma(CharacterAtPositionInJSONString_)) Then
                ChangePositionInJSONStringForward JSONText, 1, True
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting array."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting , to separate values in array."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        End If
        '...[V]alue,...
        If (ExtractValue(JSONText, D, CStr(IndexOfValue))) Then
            IsFirstValue = False
        Else
            Exit Function
        End If
        IndexOfValue = IndexOfValue + 1
        ChangePositionInJSONStringForward JSONText, 1, True
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting array."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Exit Function
        End If
        '...Value[,]... or ...Value[]]...
    Loop
    If (PositionInJSONString_ > LengthJSONString_) Then
        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting array."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        Exit Function
    Else
        ExtractArray = True
    End If
    'finish...
    Exit Function
ExtractArrayError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractArray().""}"
End Function
Private Function ExtractNameValuePair(ByRef JSONText As String, ByRef D As Dictionary) As Boolean
    On Error GoTo ExtractNameValuePairError
    '...["]Name":Value,...
    'ready?
    Dim Name As String
    Name = "" 'default value
    'set?
    ExtractNameValuePair = False 'default value
    'go!
    If (IsDoubleQuote(CharacterAtPositionInJSONString_)) Then
        If (ExtractName(JSONText, Name)) Then
            '...,"Name["]:Value,...
            ChangePositionInJSONStringForward JSONText, 1, True
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting name/value pair."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
            '...,"Name"[:]Value,...
            If (IsColon(CharacterAtPositionInJSONString_)) Then
                ChangePositionInJSONStringForward JSONText, 1, True
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting name/value pair."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
                '...,{"Name":[V]alue,...
                ExtractNameValuePair = ExtractValue(JSONText, D, Name)
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting : to separate name and value in name/value pair of object."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        Else
            Exit Function
        End If
    Else
        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting new name/value pair in object."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        Exit Function
    End If
    'finish...
    Exit Function
ExtractNameValuePairError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractNameValuePair().""}"
End Function
Private Function ExtractName(ByRef JSONText As String, ByRef Name As String) As Boolean
    ExtractName = ExtractString(JSONText, Name, True)
End Function
Private Function ExtractValue(ByRef JSONText As String, ByRef D As Dictionary, ByRef Key As String) As Boolean
    On Error GoTo ExtractValueError
    '...,{"Name":[V]alue,... or ...,[[V]alue,... or ...[V]alue,...
    'ready?
    Dim S As String
    Dim DSub As Dictionary
    Dim CSub As Dictionary
    'set?
    ExtractValue = False 'default value
    'go!
    '...,{"Name":[V]alue,...
    If (HowToConformToJSONGrammar_ = eHowToConformToJSONGrammar.BeStrict) Then
        If (IsN(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
            If (ExtractNull(JSONText)) Then
                If (Not D.Exists(Key)) Then
                    If (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseNull) Then
                        D.Add Key, Null
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseNullString) Then
                        D.Add Key, vbNullString
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseEmptyString) Then
                        D.Add Key, ""
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseCustomString) Then
                        D.Add Key, CustomStringNull_
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseZero) Then
                        D.Add Key, 0
                    Else
                        AddToLog "{""Type"":""Code"",""Description"":""Impossible path in ExtractValue().""}"
                        Exit Function
                    End If
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsF(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
            If (ExtractFalse(JSONText)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, False
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsT(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
            If (ExtractTrue(JSONText)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, True
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsSignMinus(CharacterAtPositionInJSONString_) Or IsDigit(CharacterAtPositionInJSONString_)) Then
            If (ExtractNumber(JSONText, S)) Then
                If (Not D.Exists(Key)) Then
                    If (AddNumber(JSONText, D, Key, S)) Then
                        ExtractValue = True
                    End If
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsDoubleQuote(CharacterAtPositionInJSONString_)) Then
            If (ExtractString(JSONText, S)) Then 'also used to extract name, so other function signature
                If (Not D.Exists(Key)) Then
                    D.Add Key, S
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsLeftCurlyBracket(CharacterAtPositionInJSONString_)) Then
            Set DSub = New Dictionary
            If (ExtractObject(JSONText, DSub)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, DSub
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsLeftSquareBracket(CharacterAtPositionInJSONString_)) Then
            Set CSub = New Dictionary
            If (ExtractArray(JSONText, CSub)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, CSub
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting value."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    ElseIf (HowToConformToJSONGrammar_ = eHowToConformToJSONGrammar.BeForgiving) Then
        If (IsN(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
            If (ExtractNull(JSONText, eCharacterCompare.Both)) Then
                If (Not D.Exists(Key)) Then
                    If (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseNull) Then
                        D.Add Key, Null
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseNullString) Then
                        D.Add Key, vbNullString
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseEmptyString) Then
                        D.Add Key, ""
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseCustomString) Then
                        D.Add Key, CustomStringNull_
                    ElseIf (HowToConformToJSONNull_ = eHowToConformToJSONNull.UseZero) Then
                        D.Add Key, 0
                    Else
                        AddToLog "{""Type"":""Code"",""Description"":""Impossible path in ExtractValue().""}"
                        Exit Function
                    End If
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsF(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
            If (ExtractFalse(JSONText, eCharacterCompare.Both)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, False
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsT(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
            If (ExtractTrue(JSONText, eCharacterCompare.Both)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, True
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsSignMinus(CharacterAtPositionInJSONString_) Or IsSignPlus(CharacterAtPositionInJSONString_) Or IsPoint(CharacterAtPositionInJSONString_) Or IsDigit(CharacterAtPositionInJSONString_)) Then
            If (ExtractNumber(JSONText, S)) Then
                If (Not D.Exists(Key)) Then
                    AddNumber JSONText, D, Key, S
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsDoubleQuote(CharacterAtPositionInJSONString_)) Then
            If (ExtractString(JSONText, S)) Then 'also used to extract name, so other function signature
                If (Not D.Exists(Key)) Then
                    D.Add Key, S
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsLeftCurlyBracket(CharacterAtPositionInJSONString_)) Then
            Set DSub = New Dictionary
            If (ExtractObject(JSONText, DSub)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, DSub
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (IsLeftSquareBracket(CharacterAtPositionInJSONString_)) Then
            Set CSub = New Dictionary
            If (ExtractArray(JSONText, CSub)) Then
                If (Not D.Exists(Key)) Then
                    D.Add Key, CSub
                    ExtractValue = True
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Name """ & MapTextToJSONGrammar(Key) & """ already exists in dictionary."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting value."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    End If
    'finish...
    Exit Function
ExtractValueError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractValue().""}"
End Function
Private Sub ChangePositionInJSONStringForward(ByRef JSONText As String, Optional ByVal NumberOfPositionsToMove As Long = 1, Optional SkipWhiteSpace As Boolean = True)
    If (SkipWhiteSpace) Then
        Dim NumberOfPositionsMoved As Long
        NumberOfPositionsMoved = 0
        Dim Character As String
        Character = ""
        CharacterAtPositionInJSONString_ = ""
        Do While (NumberOfPositionsMoved < NumberOfPositionsToMove)
            PositionInJSONString_ = PositionInJSONString_ + 1
            If (PositionInJSONString_ <= LengthJSONString_) Then
                Character = Mid$(JSONText, PositionInJSONString_, 1)
                If (Not IsWhiteSpaceCharacter(Character)) Then
                    CharacterAtPositionInJSONString_ = CharacterAtPositionInJSONString_ & Character
                    NumberOfPositionsMoved = NumberOfPositionsMoved + 1
                End If
            Else
                Exit Sub
            End If
        Loop
    Else
        PositionInJSONString_ = PositionInJSONString_ + NumberOfPositionsToMove
        CharacterAtPositionInJSONString_ = Mid$(JSONText, PositionInJSONString_ - NumberOfPositionsToMove + 1, NumberOfPositionsToMove)
    End If
End Sub
Private Sub ChangePositionInJSONStringBackward(ByRef JSONText As String, Optional ByVal NumberOfPositionsToMove As Long = 1, Optional SkipWhiteSpace As Boolean = True)
    If (SkipWhiteSpace) Then
        Dim NumberOfPositionsMoved As Long
        NumberOfPositionsMoved = 0
        Dim Character As String
        Character = ""
        CharacterAtPositionInJSONString_ = ""
        Do While (NumberOfPositionsMoved < NumberOfPositionsToMove)
            PositionInJSONString_ = PositionInJSONString_ - 1
            If (PositionInJSONString_ > 0) Then
                Character = Mid$(JSONText, PositionInJSONString_, 1)
                If (Not IsWhiteSpaceCharacter(Character)) Then
                    CharacterAtPositionInJSONString_ = CharacterAtPositionInJSONString_ & Character
                    NumberOfPositionsMoved = NumberOfPositionsMoved + 1
                End If
            Else
                Exit Sub
            End If
        Loop
    Else
        PositionInJSONString_ = PositionInJSONString_ - NumberOfPositionsToMove
        CharacterAtPositionInJSONString_ = Mid$(JSONText, PositionInJSONString_ - NumberOfPositionsToMove + 1, NumberOfPositionsToMove)
    End If
End Sub
Private Function ExtractString(ByRef JSONText As String, ByRef S As String, Optional ByVal IsName As Boolean = False) As Boolean
    On Error GoTo ExtractStringError
    '...,["]Name":Value,... or '...,"Name":["]Value",...
    'ready?
    'set?
    ExtractString = False
    S = ""
    'go!
    ChangePositionInJSONStringForward JSONText, 1, False
    If (PositionInJSONString_ > LengthJSONString_) Then
        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting string."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        Exit Function
    End If
    '...,"Name":"[V]alue",...
    Do While (Not IsDoubleQuote(CharacterAtPositionInJSONString_)) And (PositionInJSONString_ <= LengthJSONString_)
        If (IsSlashBackward(CharacterAtPositionInJSONString_)) Then
            'escape sequences
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting string."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
            '\[S]equence
            If (IsDoubleQuote(CharacterAtPositionInJSONString_)) Then
                S = S & """"
            ElseIf (IsSlashBackward(CharacterAtPositionInJSONString_)) Then
                S = S & "\"
            ElseIf (IsSlashForward(CharacterAtPositionInJSONString_)) Then
                S = S & "/"
            ElseIf (IsB(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                S = S & vbBack
            ElseIf (IsF(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                S = S & vbFormFeed
            ElseIf (IsN(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                S = S & vbLf
            ElseIf (IsR(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                S = S & vbCr
            ElseIf (IsT(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                S = S & vbTab
            ElseIf (IsU(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                Dim Index As Byte
                Index = 1
                ChangePositionInJSONStringForward JSONText, 1, False
                Do While ((Index <= 4) And (PositionInJSONString_ <= LengthJSONString_))
                    If (IsHexadecimalCharacter(CharacterAtPositionInJSONString_)) Then
                        Index = Index + 1
                        ChangePositionInJSONStringForward JSONText, 1, False
                    Else
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting 4 hexadecimal characters in escaped sequence in string."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Loop
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting string."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting string."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
                '\U[0000]
                S = S & ChrW$("&H" & Mid$(JSONText, PositionInJSONString_ - 4, 4))
                ChangePositionInJSONStringBackward JSONText, , False
            Else
                If (IsName) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting escape sequence in name of name/value pair."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting escape sequence in string of name/value pair."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                End If
            End If
        Else
            S = S & CharacterAtPositionInJSONString_
        End If
        ChangePositionInJSONStringForward JSONText, 1, False
        If (PositionInJSONString_ > LengthJSONString_) Then
            If (IsName) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting name."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting value."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            End If
            Exit Function
        End If
    Loop
    If (PositionInJSONString_ > LengthJSONString_) Then
        If (IsName) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting name."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting value."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
        Exit Function
    Else
        ExtractString = True
    End If
    'finish...
    Exit Function
ExtractStringError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractString().""}"
End Function
Private Function ExtractNull(ByRef JSONText As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.LowerCase) As Boolean
    On Error GoTo ExtractNullError
    '...,"Name":[n]ull,...
    'ready?
    'set?
    ExtractNull = False
    'go!
    If (CharacterCompare = eCharacterCompare.LowerCase) Then
        ChangePositionInJSONStringForward JSONText, 1, False
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        '...,"Name":n[u]ll,...
        If (IsU(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            '...,"Name":nu[l]ll,...
            If (IsL(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                '...,"Name":nul[l],...
                If (IsL(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                    ExtractNull = True
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting l in value null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting l in value null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting u in value null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    ElseIf (CharacterCompare = eCharacterCompare.Both) Then
        ChangePositionInJSONStringForward JSONText, 1, False
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        '...,"Name":n[u]ll,...
        If (IsU(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            '...,"Name":nu[l]ll,...
            If (IsL(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                '...,"Name":nul[l],...
                If (IsL(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
                    ExtractNull = True
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting l in value null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting l in value null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting u in value null."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    End If
    'finish...
    Exit Function
ExtractNullError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractNull().""}"
End Function
Private Function ExtractFalse(ByRef JSONText As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.LowerCase) As Boolean
    On Error GoTo ExtractFalseError
    '...,"Name":[f]alse,...
    'ready?
    'set?
    ExtractFalse = False
    'go!
    If (CharacterCompare = eCharacterCompare.LowerCase) Then
        ChangePositionInJSONStringForward JSONText, 1, False
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        '...,"Name":f[a]lse,...
        If (IsA(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            '...,"Name":fa[l]se,...
            If (IsL(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                '...,"Name":fal[s]e,...
                If (IsS(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                        Exit Function
                    End If
                    '...,"Name":fals[e],...
                    If (IsE(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                        ExtractFalse = True
                    Else
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting e in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    End If
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting s in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting l in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting a in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    ElseIf (CharacterCompare = eCharacterCompare.Both) Then
        ChangePositionInJSONStringForward JSONText, 1, False
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        '...,"Name":f[a]lse,...
        If (IsA(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            '...,"Name":fa[l]se,...
            If (IsL(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                '...,"Name":fal[s]e,...
                If (IsS(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                        Exit Function
                    End If
                    '...,"Name":fals[e],...
                    If (IsE(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
                        ExtractFalse = True
                    Else
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting e or E in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    End If
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting s or S in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting l or L in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting a or A in value false."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    End If
    'finish...
    Exit Function
ExtractFalseError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractFalse().""}"
End Function
Private Function ExtractTrue(ByRef JSONText As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.LowerCase) As Boolean
    On Error GoTo ExtractTrueError
    '...,"Name":[t]rue,...
    'ready?
    'set?
    ExtractTrue = False
    'go!
    If (CharacterCompare = eCharacterCompare.LowerCase) Then
        ChangePositionInJSONStringForward JSONText, 1, False
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        '...,"Name":t[r]ue,...
        If (IsR(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            '...,"Name":tr[u]e,...
            If (IsU(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                '...,"Name":tru[e],...
                If (IsE(CharacterAtPositionInJSONString_, eCharacterCompare.LowerCase)) Then
                    ExtractTrue = True
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting e in value true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting u in value true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting r in value true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    ElseIf (CharacterCompare = eCharacterCompare.Both) Then
        ChangePositionInJSONStringForward JSONText, 1, True
        If (PositionInJSONString_ > LengthJSONString_) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        '...,"Name":t[r]ue,...
        If (IsR(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            '...,"Name":tr[u]e,...
            If (IsU(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                '...,"Name":tru[e],...
                If (IsE(CharacterAtPositionInJSONString_, eCharacterCompare.Both)) Then
                    ExtractTrue = True
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting e or E in value true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                End If
            Else
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting u or U in value true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            End If
        Else
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting r or R in value true."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
        End If
    End If
    'finish...
    Exit Function
ExtractTrueError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractTrue().""}"
End Function
Private Function ExtractNumber(ByRef JSONText As String, ByRef S As String) As Boolean
    On Error GoTo ExtractNumberError
    '...,"Name":[-]1234.567e+89,...
    'ready?
    Dim SignOfNumber As String
    SignOfNumber = ""
    Dim DigitsOfNumberBeforeDecimalSign As String
    DigitsOfNumberBeforeDecimalSign = ""
    Dim NumberHasDigitsAfterDecimalSign As Boolean
    NumberHasDigitsAfterDecimalSign = False
    Dim DigitsOfNumberAfterDecimalSign As String
    DigitsOfNumberAfterDecimalSign = ""
    Dim NumberHasExponent As Boolean
    NumberHasExponent = False
    Dim SignOfExponent As String
    SignOfExponent = "+"
    Dim DigitsOfExponent As String
    DigitsOfExponent = ""
    'set?
    ExtractNumber = False
    'go!
    'step1: get sign of number
    If (HowToConformToJSONGrammar_ = eHowToConformToJSONGrammar.BeStrict) Then
        '[1]234.567e+89 or [-]1234.567e+89
        If (IsSignMinus(CharacterAtPositionInJSONString_)) Then
            SignOfNumber = "-"
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
            If (Not IsDigit(CharacterAtPositionInJSONString_)) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit after minus sign."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        End If
        'step2: get digits of number before decimal sign
        If (IsZero(CharacterAtPositionInJSONString_)) Then
            '[0].12345 or -[0].12345
            DigitsOfNumberBeforeDecimalSign = DigitsOfNumberBeforeDecimalSign & CharacterAtPositionInJSONString_
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            If (Not (IsPoint(CharacterAtPositionInJSONString_) Or IsE(CharacterAtPositionInJSONString_, eCharacterCompare.Both) Or IsNumberEndingCharacter(CharacterAtPositionInJSONString_))) Then
                If (IsDigit(CharacterAtPositionInJSONString_)) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Numbers cannot have unnecessary zeroes in front."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            End If
        Else
            '[1]234.567e+89 or -[1]234.567e+89
            DigitsOfNumberBeforeDecimalSign = DigitsOfNumberBeforeDecimalSign & CharacterAtPositionInJSONString_
            ChangePositionInJSONStringForward JSONText, 1, False
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            Do While (Not (IsPoint(CharacterAtPositionInJSONString_) Or IsE(CharacterAtPositionInJSONString_, eCharacterCompare.Both) Or IsNumberEndingCharacter(CharacterAtPositionInJSONString_)))
                If (IsDigit(CharacterAtPositionInJSONString_)) Then
                    DigitsOfNumberBeforeDecimalSign = DigitsOfNumberBeforeDecimalSign & CharacterAtPositionInJSONString_
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Loop
        End If
        If (Len(DigitsOfNumberBeforeDecimalSign) = 0) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting at least one digit before decimal sign position in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        If (Not (IsNumberEndingCharacter(CharacterAtPositionInJSONString_))) Then
            If (IsPoint(CharacterAtPositionInJSONString_)) Then
                NumberHasDigitsAfterDecimalSign = True
                'step3: get digits of number after decimal sign
                '1234[.]567e+89 or -1234[.]567e+89
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
                Do While ((Not (IsE(CharacterAtPositionInJSONString_, eCharacterCompare.Both) Or IsNumberEndingCharacter(CharacterAtPositionInJSONString_))))
                    If (IsDigit(CharacterAtPositionInJSONString_)) Then
                        DigitsOfNumberAfterDecimalSign = DigitsOfNumberAfterDecimalSign & CharacterAtPositionInJSONString_
                        ChangePositionInJSONStringForward JSONText, 1, False
                        If (PositionInJSONString_ > LengthJSONString_) Then
                            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting decimal part of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                            Exit Function
                        End If
                    Else
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Loop
            End If
            If (Not (IsNumberEndingCharacter(CharacterAtPositionInJSONString_))) Then
                NumberHasExponent = True
                'step4: get sign of exponent
                '1234.567[e]+89 or -1234.567[e]+89
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                If (IsSignMinus(CharacterAtPositionInJSONString_)) Then
                    SignOfExponent = "-"
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                        Exit Function
                    End If
                ElseIf (IsSignPlus(CharacterAtPositionInJSONString_)) Then
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                        Exit Function
                    End If
                End If
                Do While (Not (IsNumberEndingCharacter(CharacterAtPositionInJSONString_)))
                    If (IsDigit(CharacterAtPositionInJSONString_)) Then
                        DigitsOfExponent = DigitsOfExponent & CharacterAtPositionInJSONString_
                        ChangePositionInJSONStringForward JSONText, 1, False
                        If (PositionInJSONString_ > LengthJSONString_) Then
                            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                            Exit Function
                        End If
                    Else
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Loop
                If (Len(DigitsOfExponent) = 0) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting at least one digit in exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            End If
        End If
    ElseIf (HowToConformToJSONGrammar_ = eHowToConformToJSONGrammar.BeForgiving) Then
        '[1]234.567e+89 or [-]1234.567e+89 or [+]1234.567e+89
        If (IsSignMinus(CharacterAtPositionInJSONString_)) Then
            SignOfNumber = "-"
            ChangePositionInJSONStringForward JSONText, 1, True
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
            If (Not (IsDigit(CharacterAtPositionInJSONString_) Or IsPoint(CharacterAtPositionInJSONString_))) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit after minus sign."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        ElseIf (IsSignPlus(CharacterAtPositionInJSONString_)) Then
            SignOfNumber = ""
            ChangePositionInJSONStringForward JSONText, 1, True
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
            If (Not (IsDigit(CharacterAtPositionInJSONString_) Or IsPoint(CharacterAtPositionInJSONString_))) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit after plus sign."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        End If
        'step2: get digits of number before decimal sign
        '[0].12345 or -[0].12345 or [1]234.567e+89 or -[1]234.567e+89 or .1234567
        If (IsPoint(CharacterAtPositionInJSONString_)) Then
            DigitsOfNumberBeforeDecimalSign = "0"
        Else
            DigitsOfNumberBeforeDecimalSign = DigitsOfNumberBeforeDecimalSign & CharacterAtPositionInJSONString_
            ChangePositionInJSONStringForward JSONText, 1, True
            If (PositionInJSONString_ > LengthJSONString_) Then
                AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                Exit Function
            End If
            Do While (Not (IsPoint(CharacterAtPositionInJSONString_) Or IsE(CharacterAtPositionInJSONString_, eCharacterCompare.Both) Or IsNumberEndingCharacter(CharacterAtPositionInJSONString_)))
                If (IsDigit(CharacterAtPositionInJSONString_)) Then
                    DigitsOfNumberBeforeDecimalSign = DigitsOfNumberBeforeDecimalSign & CharacterAtPositionInJSONString_
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Else
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Loop
        End If
        If (Len(DigitsOfNumberBeforeDecimalSign) = 0) Then
            AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting at least one digit before decimal sign position in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
            Exit Function
        End If
        If (Not (IsNumberEndingCharacter(CharacterAtPositionInJSONString_))) Then
            If (IsPoint(CharacterAtPositionInJSONString_)) Then
                NumberHasDigitsAfterDecimalSign = True
                'step3: get digits of number after decimal sign
                '1234[.]567e+89 or -1234[.]567e+89
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
                Do While ((Not (IsE(CharacterAtPositionInJSONString_, eCharacterCompare.Both) Or IsNumberEndingCharacter(CharacterAtPositionInJSONString_))))
                    If (IsDigit(CharacterAtPositionInJSONString_)) Then
                        DigitsOfNumberAfterDecimalSign = DigitsOfNumberAfterDecimalSign & CharacterAtPositionInJSONString_
                        ChangePositionInJSONStringForward JSONText, 1, False
                        If (PositionInJSONString_ > LengthJSONString_) Then
                            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting decimal part of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                            Exit Function
                        End If
                    Else
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Loop
            End If
            If (Not (IsNumberEndingCharacter(CharacterAtPositionInJSONString_))) Then
                NumberHasExponent = True
                'step4: get sign of exponent
                '1234.567[e]+89 or -1234.567[e]+89
                ChangePositionInJSONStringForward JSONText, 1, False
                If (PositionInJSONString_ > LengthJSONString_) Then
                    AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                    Exit Function
                End If
                If (IsSignMinus(CharacterAtPositionInJSONString_)) Then
                    SignOfExponent = "-"
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                        Exit Function
                    End If
                ElseIf (IsSignPlus(CharacterAtPositionInJSONString_)) Then
                    ChangePositionInJSONStringForward JSONText, 1, False
                    If (PositionInJSONString_ > LengthJSONString_) Then
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                        Exit Function
                    End If
                End If
                Do While (Not (IsNumberEndingCharacter(CharacterAtPositionInJSONString_)))
                    If (IsDigit(CharacterAtPositionInJSONString_)) Then
                        DigitsOfExponent = DigitsOfExponent & CharacterAtPositionInJSONString_
                        ChangePositionInJSONStringForward JSONText, 1, False
                        If (PositionInJSONString_ > LengthJSONString_) Then
                            AddToLog "{""Type"":""JSON grammar"",""Description"":""Out of characters while extracting exponent of number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & Mid$(JSONText, PositionInJSONString_, 1) & """}"
                            Exit Function
                        End If
                    Else
                        AddToLog "{""Type"":""JSON grammar"",""Description"":""Expecting digit in number."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Loop
                If (Len(DigitsOfExponent) = 0) Then
                    NumberHasExponent = False
                End If
            End If
        End If
    End If
    If (NumberHasDigitsAfterDecimalSign) Then
        If (NumberHasExponent) Then
            S = SignOfNumber & DigitsOfNumberBeforeDecimalSign & DecimalSign_ & DigitsOfNumberAfterDecimalSign & "e" & SignOfExponent & DigitsOfExponent
        Else
            S = SignOfNumber & DigitsOfNumberBeforeDecimalSign & DecimalSign_ & DigitsOfNumberAfterDecimalSign
        End If
    Else
        If (NumberHasExponent) Then
            S = SignOfNumber & DigitsOfNumberBeforeDecimalSign & "e" & SignOfExponent & DigitsOfExponent
        Else
            S = SignOfNumber & DigitsOfNumberBeforeDecimalSign
        End If
    End If
    ChangePositionInJSONStringBackward JSONText, 1
    'finish...
    ExtractNumber = True
    Exit Function
ExtractNumberError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in ExtractNumber().""}"
End Function
Private Function AddNumber(ByRef JSONText As String, ByRef D As Dictionary, ByRef Key As String, ByRef S As String) As Boolean
    On Error GoTo AddNumberError
    'ready?
    'set?
    AddNumber = False
    'go!
    If (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.UseString) Then
        D.Add Key, S
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.UseLong) Then
        If (CanBeConvertedToLong(S)) Then
            D.Add Key, CLng(S)
        Else
            AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of long."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Exit Function
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.UseCurrency) Then
        If (CanBeConvertedToCurrency(S)) Then
            D.Add Key, CCur(S)
        Else
            AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of currency."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Exit Function
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.UseDouble) Then
        If (CanBeConvertedToDouble(S)) Then
            D.Add Key, CDbl(S)
        Else
            AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Exit Function
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.TryDoubleElseError) Then
        If (CanBeConvertedToDouble(S)) Then
            If (IsConversionToDoubleOk(CDbl(S), S)) Then
                D.Add Key, CDbl(S)
            Else
                AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        Else
            AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Exit Function
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.TryCurrencyElseError) Then
        If (CanBeConvertedToCurrency(S)) Then
            If (IsConversionToCurrencyOk(CCur(S), S)) Then
                D.Add Key, CCur(S)
            Else
                AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of currency."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        Else
            AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of currency."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Exit Function
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.TryLongElseError) Then
        If (CanBeConvertedToLong(S)) Then
            If (IsConversionToLongOk(CLng(S), S)) Then
                D.Add Key, CLng(S)
            Else
                AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of long."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        Else
            AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of long."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
            Exit Function
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.TryCurrencyTryDoubleElseError) Then
        If (CanBeConvertedToCurrency(S)) Then
            If (IsConversionToCurrencyOk(CCur(S), S)) Then
                D.Add Key, CCur(S)
            Else
                If (CanBeConvertedToDouble(S)) Then
                    If (IsConversionToDoubleOk(CDbl(S), S)) Then
                        D.Add Key, CDbl(S)
                    Else
                        AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of currency and not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            End If
        Else
            If (CanBeConvertedToDouble(S)) Then
                If (IsConversionToDoubleOk(CDbl(S), S)) Then
                    D.Add Key, CDbl(S)
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double and not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.TryLongTryDoubleElseError) Then
        If (CanBeConvertedToLong(S)) Then
            If (IsConversionToLongOk(CLng(S), S)) Then
                D.Add Key, CLng(S)
            Else
                If (CanBeConvertedToDouble(S)) Then
                    If (IsConversionToDoubleOk(CDbl(S), S)) Then
                        D.Add Key, CDbl(S)
                    Else
                        AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                        Exit Function
                    End If
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of long and not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            End If
        Else
            If (CanBeConvertedToDouble(S)) Then
                If (IsConversionToDoubleOk(CDbl(S), S)) Then
                    D.Add Key, CDbl(S)
                Else
                    AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                    Exit Function
                End If
            Else
                AddToLog "{""Type"":""VBA restriction"",""Description"":""Number " & MapTextToJSONGrammar(S) & " is not in range of long and not in range of double."",""Position"":" & CStr(PositionInJSONString_) & ",""CharacterAtPosition"":""" & MapTextToJSONGrammar(Mid$(JSONText, PositionInJSONString_, 1)) & """}"
                Exit Function
            End If
        End If
    ElseIf (HowToConvertJSONNumberToVBA_ = eHowToConvertJSONNumberToVBA.TryLongTryDoubleElseUseString) Then
        If (CanBeConvertedToLong(S)) Then
            If (IsConversionToLongOk(CLng(S), S)) Then
                D.Add Key, CLng(S)
            Else
                If (CanBeConvertedToDouble(S)) Then
                    If (IsConversionToDoubleOk(CDbl(S), S)) Then
                        D.Add Key, CDbl(S)
                    Else
                        D.Add Key, S
                    End If
                Else
                    D.Add Key, S
                End If
            End If
        Else
            If (CanBeConvertedToDouble(S)) Then
                If (IsConversionToDoubleOk(CDbl(S), S)) Then
                    D.Add Key, CDbl(S)
                Else
                    D.Add Key, S
                End If
            Else
                D.Add Key, S
            End If
        End If
    Else
        AddToLog "{""Type"":""Code"",""Description"":""Impossible path in AddNumber().""}"
        Exit Function
    End If
    'finish...
    AddNumber = True
    Exit Function
AddNumberError:
    AddToLog "{""Type"":""Code"",""Description"":""Unexpected error in AddNumber().""}"
End Function
Private Sub GuessDecimalSign()
    If (CDbl("1.5") = 15) Then
        DecimalSign_ = ","
    Else
        DecimalSign_ = "."
    End If
End Sub
Private Function MapTextToJSONGrammar(ByRef Text As String) As String
    Dim Index As Long
    Index = 0
    Dim Character As String
    Character = ""
    For Index = 1 To Len(Text)
        Character = Mid$(Text, Index, 1)
        If (Character = """") Then
            MapTextToJSONGrammar = MapTextToJSONGrammar & "\"""
        ElseIf (Character = "\") Then
            MapTextToJSONGrammar = MapTextToJSONGrammar & "\\"
        Else
            MapTextToJSONGrammar = MapTextToJSONGrammar & Character
        End If
    Next Index
End Function
Private Function IsLeftCurlyBracket(ByRef Character As String) As Boolean
    If (AscW(Character) = 123) Then IsLeftCurlyBracket = True
End Function
Private Function IsRightCurlyBracket(ByRef Character As String) As Boolean
    If (AscW(Character) = 125) Then IsRightCurlyBracket = True
End Function
Private Function IsLeftSquareBracket(ByRef Character As String) As Boolean
    If (AscW(Character) = 91) Then IsLeftSquareBracket = True
End Function
Private Function IsRightSquareBracket(ByRef Character As String) As Boolean
    If (AscW(Character) = 93) Then IsRightSquareBracket = True
End Function
Private Function IsDoubleQuote(ByRef Character As String) As Boolean
    If (AscW(Character) = 34) Then IsDoubleQuote = True
End Function
Private Function IsComma(ByRef Character As String) As Boolean
    If (AscW(Character) = 44) Then IsComma = True
End Function
Private Function IsColon(ByRef Character As String) As Boolean
    If (AscW(Character) = 58) Then IsColon = True
End Function
Private Function IsSlashBackward(ByRef Character As String) As Boolean
    If (AscW(Character) = 92) Then IsSlashBackward = True
End Function
Private Function IsSlashForward(ByRef Character As String) As Boolean
    If (AscW(Character) = 47) Then IsSlashForward = True
End Function
Private Function IsSignMinus(ByRef Character As String) As Boolean
    If (AscW(Character) = 45) Then IsSignMinus = True
End Function
Private Function IsSignPlus(ByRef Character As String) As Boolean
    If (AscW(Character) = 43) Then IsSignPlus = True
End Function
Private Function IsZero(ByRef Character As String) As Boolean
    If (AscW(Character) = 48) Then IsZero = True
End Function
Private Function IsPoint(ByRef Character As String) As Boolean
    If (AscW(Character) = 46) Then IsPoint = True
End Function
Private Function IsA(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 65) Then IsA = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 97) Then IsA = True
    Else
        If (AscW(Character) = 65) Or (AscW(Character) = 97) Then IsA = True
    End If
End Function
Private Function IsB(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 66) Then IsB = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 98) Then IsB = True
    Else
        If (AscW(Character) = 66) Or (AscW(Character) = 98) Then IsB = True
    End If
End Function
Private Function IsE(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 69) Then IsE = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 101) Then IsE = True
    Else
        If (AscW(Character) = 69) Or (AscW(Character) = 101) Then IsE = True
    End If
End Function
Private Function IsF(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 70) Then IsF = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 102) Then IsF = True
    Else
        If (AscW(Character) = 70) Or (AscW(Character) = 102) Then IsF = True
    End If
End Function
Private Function IsL(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 76) Then IsL = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 108) Then IsL = True
    Else
        If (AscW(Character) = 76) Or (AscW(Character) = 108) Then IsL = True
    End If
End Function
Private Function IsN(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 78) Then IsN = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 110) Then IsN = True
    Else
        If (AscW(Character) = 78) Or (AscW(Character) = 110) Then IsN = True
    End If
End Function
Private Function IsR(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 82) Then IsR = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 114) Then IsR = True
    Else
        If (AscW(Character) = 82) Or (AscW(Character) = 114) Then IsR = True
    End If
End Function
Private Function IsS(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 83) Then IsS = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 115) Then IsS = True
    Else
        If (AscW(Character) = 83) Or (AscW(Character) = 115) Then IsS = True
    End If
End Function
Private Function IsT(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 84) Then IsT = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 116) Then IsT = True
    Else
        If (AscW(Character) = 84) Or (AscW(Character) = 116) Then IsT = True
    End If
End Function
Private Function IsU(ByRef Character As String, Optional ByVal CharacterCompare As eCharacterCompare = eCharacterCompare.Both) As Boolean
    If (CharacterCompare = eCharacterCompare.UpperCase) Then
        If (AscW(Character) = 85) Then IsU = True
    ElseIf (CharacterCompare = eCharacterCompare.LowerCase) Then
        If (AscW(Character) = 117) Then IsU = True
    Else
        If (AscW(Character) = 85) Or (AscW(Character) = 117) Then IsU = True
    End If
End Function
Private Function IsWhiteSpaceCharacter(ByRef Character As String) As Boolean
    Dim I As Integer
    I = AscW(Character)
    If (I = 9) Or (I = 10) Or (I = 13) Or (I = 32) Then 'horizontal tab, line feed, carriage return
        IsWhiteSpaceCharacter = True
    End If
End Function
Private Function IsHexadecimalCharacter(ByRef Character As String) As Boolean
    Dim I As Integer
    I = AscW(Character)
    If (I > 47) And (I < 58) Then '0,1,...,9
        IsHexadecimalCharacter = True
    ElseIf (I > 64) And (I < 71) Then 'A,B,...,F
        IsHexadecimalCharacter = True
    ElseIf (I > 96) And (I < 103) Then 'a,b,...,f
        IsHexadecimalCharacter = True
    End If
End Function
Private Function IsDigit(ByRef Character As String) As Boolean
    Dim I As Integer
    I = AscW(Character) 'decimal code (integer) of the first character in (unicode) string
    If (I > 47) And (I < 58) Then '0,1,...,9
        IsDigit = True
    End If
End Function
'Private Function IsDigitButNotZero(ByRef Character As String) As Boolean
'    Dim I As Integer
'    I = AscW(Character) 'decimal code (integer) of the first character in (unicode) string
'    If (I > 48) And (I < 58) Then '1,2,...,9
'        IsDigitButNotZero = True
'    End If
'End Function
Private Function IsNumberEndingCharacter(ByRef Character As String) As Boolean
    Dim I As Integer
    I = AscW(Character) 'decimal code (integer) of the first character in (unicode) string
    If (I = 44) Or (I = 93) Or (I = 125) Or IsWhiteSpaceCharacter(Character) Then 'comma, right square bracket, right curly bracket
        IsNumberEndingCharacter = True
    End If
End Function
Private Function CanBeConvertedToLong(ByVal S As String) As Boolean
    On Error GoTo CanBeConvertedToLongError
    'negative, max      CanBeConvertedToLong("-2147483648")
    'positive, max      CanBeConvertedToLong("2147483647")
    Dim L As Long
    L = CLng(S)
    CanBeConvertedToLong = True
    Exit Function
CanBeConvertedToLongError:
    CanBeConvertedToLong = False
End Function
Private Function CanBeConvertedToCurrency(ByVal S As String) As Boolean
    On Error GoTo CanBeConvertedToCurrencyError
    'negative, max      CanBeConvertedToCurrency("-922337203685477,5808")
    'positive, max      CanBeConvertedToCurrency("922337203685477,5807")
    Dim C As Currency
    C = CCur(S)
    CanBeConvertedToCurrency = True
    Exit Function
CanBeConvertedToCurrencyError:
    CanBeConvertedToCurrency = False
End Function
Private Function CanBeConvertedToDouble(ByVal S As String) As Boolean
    On Error GoTo CanBeConvertedToDoubleError
    'negative, max      CanBeConvertedToDouble("-1,79769313486231E+308")
    'negative, smallest CanBeConvertedToDouble("-4,94065645841247E-324")
    'positive, max      CanBeConvertedToDouble("1,79769313486231E+308")
    'positive, smallest CanBeConvertedToDouble("4,94065645841247E-324")
    Dim D As Double
    D = CDbl(S)
    CanBeConvertedToDouble = True
    Exit Function
CanBeConvertedToDoubleError:
    CanBeConvertedToDouble = False
End Function
Private Function IsConversionToLongOk(ByVal L As Long, ByVal S As String) As Boolean
    On Error GoTo IsConversionToLongOkError
    'CLng(1.234) = 1, check for this
    If (ConvertToStandardLook(S, ",") = ConvertToStandardLook(CStr(L), ",")) Then
        IsConversionToLongOk = True
        Exit Function
    End If
    If (ConvertToStandardLook(S, ".") = ConvertToStandardLook(CStr(L), ".")) Then
        IsConversionToLongOk = True
        Exit Function
    End If
    Exit Function
IsConversionToLongOkError:
    IsConversionToLongOk = False
End Function
Private Function IsConversionToCurrencyOk(ByVal C As Currency, ByVal S As String) As Boolean
    On Error GoTo IsConversionToCurrencyOkError
    'CCurr(0.12345678) = 0.1235, check for this
    If (ConvertToStandardLook(S, ",") = ConvertToStandardLook(CStr(C), ",")) Then
        IsConversionToCurrencyOk = True
        Exit Function
    End If
    If (ConvertToStandardLook(S, ".") = ConvertToStandardLook(CStr(C), ".")) Then
        IsConversionToCurrencyOk = True
        Exit Function
    End If
    Exit Function
IsConversionToCurrencyOkError:
    IsConversionToCurrencyOk = False
End Function
Private Function IsConversionToDoubleOk(ByVal D As Double, ByVal S As String) As Boolean
    On Error GoTo IsConversionToDoubleOkError
    'CDbl(1234512345e123) = 1,234512345E+132, check for this
    If (ConvertToStandardLook(S, ",") = ConvertToStandardLook(CStr(D), ",")) Then
        IsConversionToDoubleOk = True
        Exit Function
    End If
    If (ConvertToStandardLook(S, ".") = ConvertToStandardLook(CStr(D), ".")) Then
        IsConversionToDoubleOk = True
        Exit Function
    End If
    Exit Function
IsConversionToDoubleOkError:
    IsConversionToDoubleOk = False
End Function
Private Function ConvertToStandardLook(ByRef S As String, ByRef DecimalSign As String) As String
    On Error GoTo ConvertToStandardLookError
    Dim PositionOfSignOfNumber As Long
    PositionOfSignOfNumber = 0
    Dim SignOfNumber As String
    SignOfNumber = ""
    Dim DigitsOfNumberBeforeDecimalSign As String
    DigitsOfNumberBeforeDecimalSign = ""
    Dim PositionOfDecimalSign As Long
    PositionOfDecimalSign = 0
    Dim NumberHasDigitsAfterDecimalSign As Boolean
    NumberHasDigitsAfterDecimalSign = False
    Dim DigitsOfNumberAfterDecimalSign As String
    DigitsOfNumberAfterDecimalSign = ""
    Dim PositionOfExponent As Long
    PositionOfExponent = 0
    Dim PositionOfSignOfExponent As Long
    PositionOfSignOfExponent = 0
    Dim SignOfExponent As String
    SignOfExponent = ""
    Dim DigitsOfExponent As String
    DigitsOfExponent = ""
    Dim Index As Long
    Index = 0
    'step 1: decompose number
    If (Mid$(S, 1, 1) = "-") Then
        SignOfNumber = "-"
        PositionOfSignOfNumber = 1
    ElseIf (Mid$(S, 1, 1) = "+") Then
        SignOfNumber = "+"
        PositionOfSignOfNumber = 1
    Else
        SignOfNumber = "+"
        PositionOfSignOfNumber = 0
    End If
    PositionOfDecimalSign = InStr(1, S, DecimalSign, vbBinaryCompare)
    PositionOfExponent = InStr(1, S, "e", vbTextCompare)
    If (PositionOfDecimalSign > 0) Then
        If (PositionOfExponent > 0) Then
            '-123.123e123
            DigitsOfNumberBeforeDecimalSign = Mid$(S, PositionOfSignOfNumber + 1, PositionOfDecimalSign - PositionOfSignOfNumber - 1)
            NumberHasDigitsAfterDecimalSign = (PositionOfExponent > PositionOfDecimalSign)
            DigitsOfNumberAfterDecimalSign = Mid$(S, PositionOfDecimalSign + 1, PositionOfExponent - PositionOfDecimalSign - 1)
            If (Mid$(S, PositionOfExponent + 1, 1) = "-") Then
                SignOfExponent = "-"
                PositionOfSignOfExponent = PositionOfExponent + 1
            ElseIf (Mid$(S, PositionOfExponent + 1, 1) = "+") Then
                SignOfExponent = "+"
                PositionOfSignOfExponent = PositionOfExponent + 1
            Else
                SignOfExponent = "+"
                PositionOfSignOfExponent = PositionOfExponent
            End If
            DigitsOfExponent = Mid$(S, PositionOfSignOfExponent + 1, Len(S))
        Else
            '-123.123
            DigitsOfNumberBeforeDecimalSign = Mid$(S, PositionOfSignOfNumber + 1, PositionOfDecimalSign - PositionOfSignOfNumber - 1)
            NumberHasDigitsAfterDecimalSign = (Len(S) > PositionOfDecimalSign)
            DigitsOfNumberAfterDecimalSign = Mid$(S, PositionOfDecimalSign + 1, Len(S))
        End If
    Else
        NumberHasDigitsAfterDecimalSign = False
        If (PositionOfExponent > 0) Then
            '-123e123
            DigitsOfNumberBeforeDecimalSign = Mid$(S, PositionOfSignOfNumber + 1, PositionOfExponent - PositionOfSignOfNumber - 1)
            If (Mid$(S, PositionOfExponent + 1, 1) = "-") Then
                SignOfExponent = "-"
                PositionOfSignOfExponent = PositionOfExponent + 1
            ElseIf (Mid$(S, PositionOfExponent + 1, 1) = "+") Then
                SignOfExponent = "+"
                PositionOfSignOfExponent = PositionOfExponent + 1
            Else
                SignOfExponent = "+"
                PositionOfSignOfExponent = PositionOfExponent
            End If
            DigitsOfExponent = Mid$(S, PositionOfSignOfExponent + 1, Len(S))
        Else
            '-123
            DigitsOfNumberBeforeDecimalSign = Mid$(S, PositionOfSignOfNumber + 1, Len(S))
        End If
    End If
    If (Len(DigitsOfNumberAfterDecimalSign) = 0) Then NumberHasDigitsAfterDecimalSign = False
    'step 2: rebuild number
    ConvertToStandardLook = SignOfNumber
    ConvertToStandardLook = ConvertToStandardLook & DigitsOfNumberBeforeDecimalSign & DecimalSign_
    If (NumberHasDigitsAfterDecimalSign) Then
        ConvertToStandardLook = ConvertToStandardLook & DigitsOfNumberAfterDecimalSign
    End If
    If (Len(DigitsOfExponent) > 0) Then
        If (CLng(DigitsOfExponent) < 325) Then 'double cannot have exponent > 324
            If (SignOfExponent = "+") Then
                For Index = 1 To CLng(DigitsOfExponent)
                    MoveDecimalSign ConvertToStandardLook, "ToTheRight", DecimalSign_
                Next Index
            Else
                For Index = 1 To CLng(DigitsOfExponent)
                    MoveDecimalSign ConvertToStandardLook, "ToTheLeft", DecimalSign_
                Next Index
            End If
        Else
            ConvertToStandardLook = S
        End If
    End If
    'step 3: adjust strange formats
    If (InStr(1, ConvertToStandardLook, DecimalSign_, vbBinaryCompare) > 0) Then
        Do While (Right$(ConvertToStandardLook, 1) = "0")
            ConvertToStandardLook = Mid$(ConvertToStandardLook, 1, Len(ConvertToStandardLook) - 1)
        Loop
    End If
    Do While (Len(ConvertToStandardLook) > 2 And (Mid$(ConvertToStandardLook, 2, 1) = "0"))
        ConvertToStandardLook = Mid$(ConvertToStandardLook, 1, 1) & Mid$(ConvertToStandardLook, 3, Len(ConvertToStandardLook))
    Loop
    If (Mid$(ConvertToStandardLook, 2, 1) = DecimalSign_) Then
        ConvertToStandardLook = Mid$(ConvertToStandardLook, 1, 1) & "0" & Mid$(ConvertToStandardLook, 2, Len(ConvertToStandardLook))
    End If
    If (ConvertToStandardLook = "-0" & DecimalSign_) Then ConvertToStandardLook = "+0" & DecimalSign_
    Exit Function
ConvertToStandardLookError:
    ConvertToStandardLook = S
End Function
Private Sub MoveDecimalSign(ByRef S As String, ByRef How As String, ByRef DecimalSign As String)
    Dim PositionOfDecimalSign As Long
    PositionOfDecimalSign = InStr(1, S, DecimalSign, vbBinaryCompare)
    If (How = "ToTheRight") Then
        If (PositionOfDecimalSign = Len(S)) Then
            '+12.
            S = Replace(S, DecimalSign, "0" & DecimalSign)
        ElseIf (PositionOfDecimalSign + 1 = Len(S)) Then
            '+12.3
            S = Left$(S, PositionOfDecimalSign - 1) & Mid$(S, PositionOfDecimalSign + 1, 1) & DecimalSign
        Else
            '+12.34
            S = Left$(S, PositionOfDecimalSign - 1) & Mid$(S, PositionOfDecimalSign + 1, 1) & DecimalSign & Mid$(S, PositionOfDecimalSign + 2, Len(S) - PositionOfDecimalSign - 1)
        End If
    ElseIf (How = "ToTheLeft") Then
        If (PositionOfDecimalSign = 3) Then
            '+1.2
            S = Left$(S, 1) & "0" & DecimalSign & Mid$(S, 2, 1) & Mid$(S, PositionOfDecimalSign + 1, Len(S))
        Else
            '+12.3
            S = Left$(S, PositionOfDecimalSign - 2) & DecimalSign & Mid$(S, PositionOfDecimalSign - 1, 1) & Mid$(S, PositionOfDecimalSign + 1, Len(S))
        End If
    End If
End Sub
