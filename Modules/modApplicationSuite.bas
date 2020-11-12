Attribute VB_Name = "modApplicationSuite"
''' modApplicationSuite
''' A module to help us handle the applications bundled with this one.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Opens the Order Importer application.
Public Sub OpenOrderImporter()
    Dim intResponse As Integer
    Dim strExePath As String
    
    ' Check if it exists.
    strExePath = OrderImporterPath
    If strExePath = vbNullString Then
        intResponse = MsgBox("Couldn't locate the Order Importer executable. Do you " & _
            "want to search for it?", vbYesNo + vbQuestion, "Couldn't Locate Executable")
        
        ' Search for it?
        If intResponse = vbYes Then
            dlgPathOptions.BrowseOrderImporter
            dlgPathOptions.ShowModal frmMain
            
            ' Let's go again.
            OpenOrderImporter
        End If
        
        ' Give up.
        Exit Sub
    End If
    
    ' Fire up the application.
    RunExecutable frmMain.hWnd, strExePath
End Sub
