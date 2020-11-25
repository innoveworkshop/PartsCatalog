Attribute VB_Name = "modCSVParser"
''' modCSVParser
''' Properly parses a CSV file.
'''
''' Author: Merri <https://www.vbforums.com/showthread.php?604031>
Option Explicit

Private Declare Function ArrayPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal saPtr As Long, saBound As Long) As Long

' returns one dimensional zero based string array in ResultSplit containing parsed CSV cells
' - ResultCols (in/out) number of columns; if positive on input the CSV data is fixed to given number of columns
' - ResultRows (out) number of rows
Public Sub ParseCSV(Expression As String, ResultSplit() As String, ResultCols As Long, ResultRows As Long, Optional ColumnDelimiter As String = ",", Optional RowDelimiter As String = vbNewLine, Optional Quote As String = """")
    Dim CSV() As Integer, HeaderCSV(5) As Long, lngCSV As Long
    ' general variables that we need
    Dim intColumn As Integer, intQuote As Integer, lngRow As Long, strRow As String
    Dim lngExpLen As Long, lngRowLen As Long
    Dim blnQuote As Boolean, lngA As Long, lngB As Long, lngC As Long, lngCount As Long, lngResults() As Long
    ' some dummy variables that we happen to need
    Dim Compare As VbCompareMethod, SafeArrayBound(1) As Long
    ' length information
    lngExpLen = LenB(Expression)
    lngRowLen = LenB(RowDelimiter)
    ' validate lengths
    If lngExpLen > 0 And lngRowLen > 0 Then
        ' column delimiter
        If LenB(ColumnDelimiter) Then intColumn = AscW(ColumnDelimiter): ColumnDelimiter = Left$(ColumnDelimiter, 1) Else intColumn = 44: ColumnDelimiter = ","
        ' quote character
        If LenB(Quote) Then intQuote = AscW(Quote): Quote = Left$(Quote, 1) Else intQuote = 34: Quote = """"
        ' maximum number of results
        ReDim lngResults(0 To (lngExpLen \ lngRowLen))
        ' prepare CSV array
        HeaderCSV(0) = 1
        HeaderCSV(1) = 2
        HeaderCSV(3) = StrPtr(Expression)
        HeaderCSV(4) = Len(Expression)
        ' assign Expression data to the Integer array
        lngCSV = ArrayPtr(CSV)
        PutMem4 lngCSV, VarPtr(HeaderCSV(0))
        ' find first row delimiter, see if within quote or not
        lngA = InStrB(1, Expression, RowDelimiter, Compare)
        Do Until (lngA And 1) Or (lngA = 0)
            lngA = InStrB(lngA + 1, Expression, RowDelimiter, Compare)
        Loop
        lngB = InStrB(1, Expression, Quote, Compare)
        Do Until (lngB And 1) Or (lngB = 0)
            lngB = InStrB(lngB + 1, Expression, Quote, Compare)
        Loop
        Do While lngA > 0
            If lngA + lngRowLen <= lngB Or lngB = 0 Then
                lngResults(lngCount) = lngA
                lngA = InStrB(lngA + lngRowLen, Expression, RowDelimiter, Compare)
                Do Until (lngA And 1) Or (lngA = 0)
                    lngA = InStrB(lngA + 1, Expression, RowDelimiter, Compare)
                Loop
                If lngCount Then
                    lngCount = lngCount + 1
                Else
                    ' calculate number of resulting columns if invalid number of columns
                    If ResultCols < 1 Then
                        ResultCols = 1
                        intColumn = AscW(ColumnDelimiter)
                        For lngC = 0 To (lngResults(0) - 1) \ 2
                            If blnQuote Then
                                If CSV(lngC) <> intQuote Then Else blnQuote = False
                            Else
                                Select Case CSV(lngC)
                                    Case intQuote
                                        blnQuote = True
                                    Case intColumn
                                        ResultCols = ResultCols + 1
                                End Select
                            End If
                        Next lngC
                    End If
                    lngCount = 1
                End If
            Else
                lngB = InStrB(lngB + 2, Expression, Quote, Compare)
                Do Until (lngB And 1) Or (lngB = 0)
                    lngB = InStrB(lngB + 1, Expression, Quote, Compare)
                Loop
                If lngB Then
                    lngA = InStrB(lngB + 2, Expression, RowDelimiter, Compare)
                    Do Until (lngA And 1) Or (lngA = 0)
                        lngA = InStrB(lngA + 1, Expression, RowDelimiter, Compare)
                    Loop
                    If lngA Then
                        lngB = InStrB(lngB + 2, Expression, Quote, Compare)
                        Do Until (lngB And 1) Or (lngB = 0)
                            lngB = InStrB(lngB + 1, Expression, Quote, Compare)
                        Loop
                    End If
                End If
            End If
        Loop
        lngResults(lngCount) = lngExpLen + 1
        ' number of rows
        ResultRows = lngCount + 1
        ' string array items to return
        ReDim Preserve ResultSplit(0 To ResultRows * ResultCols - 1)
        ' first row
        lngCount = 0
        strRow = LeftB$(Expression, lngResults(0) - 1)
        HeaderCSV(3) = StrPtr(strRow)
        lngC = 0
        blnQuote = False
        For lngB = 0 To (lngResults(0) - 1) \ 2
            If blnQuote Then
                Select Case CSV(lngB)
                    Case intQuote
                        If CSV(lngB + 1) = intQuote Then
                            ' skip next char (quote)
                            lngB = lngB + 1
                            ' add quote char
                            CSV(lngC) = intQuote
                            lngC = lngC + 1
                        Else
                            blnQuote = False
                        End If
                    Case Else
                        ' add this char
                        If lngB > lngC Then CSV(lngC) = CSV(lngB)
                        lngC = lngC + 1
                End Select
            Else
                Select Case CSV(lngB)
                    Case intQuote
                        blnQuote = True
                    Case intColumn
                        ' add this column
                        ResultSplit(lngCount) = Left$(strRow, lngC)
                        ' max column reached?
                        lngCount = lngCount + 1
                        If lngCount >= ResultCols Then Exit For
                        ' start filling column string buffer from start (strRow)
                        lngC = 0
                    Case Else
                        ' add this char
                        If lngB > lngC Then CSV(lngC) = CSV(lngB)
                        lngC = lngC + 1
                End Select
            End If
        Next lngB
        ' add last column item?
        If lngCount < ResultCols Then ResultSplit(lngCount) = Left$(strRow, lngC - 1)
        ' rows after first
        For lngA = 1 To ResultRows - 1
            ' start index for columns
            lngRow = lngA * ResultCols
            lngCount = 0
            strRow = MidB$(Expression, lngResults(lngA - 1) + lngRowLen, lngResults(lngA) - lngResults(lngA - 1) - lngRowLen)
            HeaderCSV(3) = StrPtr(strRow)
            lngC = 0
            blnQuote = False
            For lngB = 0 To (lngResults(lngA) - lngResults(lngA - 1) - lngRowLen) \ 2
                If blnQuote Then
                    Select Case CSV(lngB)
                        Case intQuote
                            If CSV(lngB + 1) = intQuote Then
                                ' skip next char (quote)
                                lngB = lngB + 1
                                ' add quote char
                                CSV(lngC) = intQuote
                                lngC = lngC + 1
                            Else
                                blnQuote = False
                            End If
                        Case Else
                            ' add this char
                            CSV(lngC) = CSV(lngB)
                            lngC = lngC + 1
                    End Select
                Else
                    Select Case CSV(lngB)
                        Case intQuote
                            blnQuote = True
                        Case intColumn
                            ' add this column
                            ResultSplit(lngRow + lngCount) = Left$(strRow, lngC)
                            ' max column reached?
                            lngCount = lngCount + 1
                            If lngCount >= ResultCols Then Exit For
                            ' start filling column string buffer from start (strRow)
                            lngC = 0
                        Case Else
                            ' add this char
                            If lngB > lngC Then CSV(lngC) = CSV(lngB)
                            lngC = lngC + 1
                    End Select
                End If
            Next lngB
            ' add last column item?
            If lngCount < ResultCols Then ResultSplit(lngRow + lngCount) = Left$(strRow, lngC - 1)
        Next lngA
        ' clean up CSV array
        PutMem4 lngCSV, 0
    Else
        ResultCols = 0
        ResultRows = 0
        ' clean any possible data that exists in the passed string array (like if it is multidimensional)
        If Not Not ResultSplit Then Erase ResultSplit
        ' mysterious IDE error fix
        Debug.Assert App.hInstance
        ' reset to one element, one dimension
        ReDim ResultSplit(0 To 0)
        ' custom redimension: remove the items (this duplicates the VB6 Split behavior)
        SafeArrayRedim Not Not ResultSplit, SafeArrayBound(0)
    End If
End Sub
