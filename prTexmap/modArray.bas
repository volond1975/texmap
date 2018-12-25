Attribute VB_Name = "modArray"
Option Explicit

#If Win64 Then
    Public Const PTR_LENGTH As Long = 8
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Public Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
#Else
    Public Const PTR_LENGTH As Long = 4
    Public Declare Function GetTickCount Lib "kernel32" () As Long
    Public Declare Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
#End If

Private Type SAFEARRAYBOUND
    cElements    As Long
    lLbound      As Long
End Type

Private Type SAFEARRAY_VECTOR
    cDims        As Integer
    fFeatures    As Integer
    cbElements   As Long
    cLocks       As Long
    pvData       As LongPtr
    rgsabound(0) As SAFEARRAYBOUND
End Type

Sub SliceColumn(ByVal idx As Long, ByRef arrayToSlice() As Variant, ByRef slicedArray As Variant)
'slicedArray can be passed as a 1d or 2d array
'sliceArray can also be part bound, eg  slicedArray(1 to 100) or slicedArray(10 to 100)
Dim ptrToArrayVar As LongPtr
Dim ptrToSafeArray As LongPtr
Dim ptrToArrayData As LongPtr
Dim ptrToArrayData2 As LongPtr
Dim uSAFEARRAY As SAFEARRAY_VECTOR
Dim ptrCursor As LongPtr
Dim cbElements As Long
Dim atsBound1 As Long
Dim elSize As Long

    'determine bound1 of source array (ie row Count)
    atsBound1 = UBound(arrayToSlice, 1)
    'get pointer to source array Safearray
    ptrToArrayVar = VarPtrArray(arrayToSlice)
    CopyMemory ptrToSafeArray, ByVal ptrToArrayVar, PTR_LENGTH
    CopyMemory uSAFEARRAY, ByVal ptrToSafeArray, LenB(uSAFEARRAY)
    ptrToArrayData = uSAFEARRAY.pvData
    'determine byte size of source elements
    cbElements = uSAFEARRAY.cbElements

    'get pointer to destination array Safearray
    ptrToArrayVar = VarPtr(slicedArray) + 8 'Variant reserves first 8bytes
    CopyMemory ptrToSafeArray, ByVal ptrToArrayVar, PTR_LENGTH
    CopyMemory uSAFEARRAY, ByVal ptrToSafeArray, LenB(uSAFEARRAY)
    ptrToArrayData2 = uSAFEARRAY.pvData

    'determine elements size
    elSize = UBound(slicedArray, 1) - LBound(slicedArray, 1) + 1
    'determine start position of data in source array
    ptrCursor = ptrToArrayData + (((idx - 1) * atsBound1 + LBound(slicedArray, 1) - 1) * cbElements)
    'Copy source array to destination array
    CopyMemory ByVal ptrToArrayData2, ByVal ptrCursor, cbElements * elSize

End Sub

Sub SliceRow(ByVal idx As Long, ByRef arrayToSlice() As Variant, ByRef slicedArray As Variant)
'slicedArray can be passed as a 1d or 2d array
'sliceArray can also be part bound, eg  slicedArray(1 to 100) or slicedArray(10 to 100)
Dim ptrToArrayVar As LongPtr
Dim ptrToSafeArray As LongPtr
Dim ptrToArrayData As LongPtr
Dim ptrToArrayData2 As LongPtr
Dim uSAFEARRAY As SAFEARRAY_VECTOR
Dim ptrCursor As LongPtr
Dim cbElements As Long
Dim atsBound1 As Long
Dim i As Long

    'determine bound1 of source array (ie row Count)
    atsBound1 = UBound(arrayToSlice, 1)
    'get pointer to source array Safearray
    ptrToArrayVar = VarPtrArray(arrayToSlice)
    CopyMemory ptrToSafeArray, ByVal ptrToArrayVar, PTR_LENGTH
    CopyMemory uSAFEARRAY, ByVal ptrToSafeArray, LenB(uSAFEARRAY)
    ptrToArrayData = uSAFEARRAY.pvData
    'determine byte size of source elements
    cbElements = uSAFEARRAY.cbElements

    'get pointer to destination array Safearray
    ptrToArrayVar = VarPtr(slicedArray) + 8 'Variant reserves first 8bytes
    CopyMemory ptrToSafeArray, ByVal ptrToArrayVar, PTR_LENGTH
    CopyMemory uSAFEARRAY, ByVal ptrToSafeArray, LenB(uSAFEARRAY)
    ptrToArrayData2 = uSAFEARRAY.pvData

    ptrCursor = ptrToArrayData + ((idx - 1) * cbElements)
    For i = LBound(slicedArray, 1) To UBound(slicedArray, 1)

        CopyMemory ByVal ptrToArrayData2, ByVal ptrCursor, cbElements
        ptrCursor = ptrCursor + (cbElements * atsBound1)
        ptrToArrayData2 = ptrToArrayData2 + cbElements
    Next i

End Sub
