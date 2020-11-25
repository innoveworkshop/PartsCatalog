Attribute VB_Name = "modFarnell"
''' modFarnell
''' Farnell Portugal order parser module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Order file columns.
Private Enum Columns
    colOrderNumber
    colOrderConfirmationNumber
    colDeliveryETA
    colOrderStatus
    colTrackingCode
    colOrderDate
    colCurrency
    colTotal
    colShippingCost
    colImportTax
    colTaxes
    colOrderTotal
    colVouchers
    colOrigin
    colOrderCode
    colCustomPartNumber
    colLineNote
    colDescription
    colManufacturer
    colMfgPartNumber
    colQuantity
    colUnitPrice
    colItemTotalPrice
End Enum

' Parse the Farnell order CSV file.
Public Sub ParseFarnellOrder(strPath As String)
    Dim astrOrder() As String
    Dim lngRows As Long
    Dim lngCols As Long
    Dim hndFile As Integer
    Dim strContents As String

    ' Read the entire file into a string.
    hndFile = FreeFile()
    Open strPath For Input As #hndFile
        strContents = Input(LOF(1), #hndFile)
    Close #hndFile

    ' Parse the CSV file and initialize the components array.
    ParseCSV strContents, astrOrder, lngCols, lngRows
    InitializeComponentsArray (lngRows - 1)
    
    ' Populate components array.
    Dim idxRow As Long
    Dim strDescription As String
    Dim astrProperties() As String
    For idxRow = 1 To (lngRows - 1)
        ' Skip empty rows.
        If astrOrder(idxRow * lngCols + colQuantity) <> "" Then
            ' Separate description from properties.
            strDescription = Trim(astrOrder(idxRow * lngCols + colDescription))
            If strDescription = "" Then
                ' Looks like the description field is empty.
                ReDim astrProperties(0)
            Else
                ' Separate the properties from the description.
                strDescription = Replace(strDescription, ":", ": ")
                astrProperties = Split(Mid(strDescription, InStr(strDescription, ";") + 2), "; ")
                strDescription = Left(strDescription, InStr(strDescription, ";") - 1)
                
                ' Check if we have a description without any properties.
                If UBound(astrProperties) < 0 Then
                    ReDim astrProperties(0)
                End If
            End If
            
            ' Add component to the array.
            AddComponent astrOrder(idxRow * lngCols + colMfgPartNumber), _
                         strDescription, _
                         astrProperties, _
                         CLng(astrOrder(idxRow * lngCols + colQuantity)), _
                         astrOrder(idxRow * lngCols + colOrderCode)
        End If
    Next idxRow
End Sub
