Attribute VB_Name = "modDebugging"
''' modDebugging
''' A simple module to make debugging this application easier.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private constants.
Private Const DBPATH As String = "\\mulberry\PartCat\DatabaseTesting\PartCat.mdb"
Private Const ORDERPATH As String = "\\mulberry\PartCat\DatabaseTesting\ORDERLINEOrderDetail.csv"

' Checks if the application is running from within the IDE.
Public Function InIDE() As Boolean
    InIDE = CBool(App.LogMode = 0)
End Function

' Automatically fill in and setup the application for debugging.
Public Sub SetupApplicationForDebug(Optional blnCheckInIDE As Boolean = True)
    ' Check if we are running in the IDE.
    If blnCheckInIDE Then
        If Not InIDE Then
            Exit Sub
        End If
    End If
    
    ' Open the testing database.
    frmMain.OpenDatabaseFile DBPATH
    
    ' Open the order CSV file.
    frmMain.txtOrderLocation.Text = ORDERPATH
    frmMain.ImportOrder
End Sub
