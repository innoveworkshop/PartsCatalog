Attribute VB_Name = "modArrayHelper"
''' modArrayHelper
''' A collection of helpful functions for dealing with arrays.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com

Option Explicit

' Checks if an array is empty. Source: https://stackoverflow.com/a/53377717/126353
Public Function IsArrayEmpty(arr As Variant) As Boolean
    Dim lb As Long
    Dim ub As Long

    Err.Clear
    On Error Resume Next

    If IsArray(arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    ub = UBound(arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I
        ' cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occasions, LBound is 0 and
        ' UBound is -1.
        ' To accommodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        lb = LBound(arr)
        If lb > ub Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

    Err.Clear
End Function
