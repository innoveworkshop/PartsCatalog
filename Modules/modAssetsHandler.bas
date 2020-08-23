Attribute VB_Name = "modAssetsHandler"
''' modAssetsHandler
''' Handles all the assets.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Win32 API imports.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' Constants.
Private Const IMAGE_FOLDER As String = "Images\"
Private Const DATASHEET_FOLDER As String = "Datasheets\"
Private Const IMAGE_EXT As String = ".bmp"
Private Const DATASHEET_EXT As String = ".pdf"

' Get path to the datasheets directory.
Public Function GetDatasheetsDirectory() As String
    GetDatasheetsDirectory = GetWorkspacePath() & DATASHEET_FOLDER
End Function

' Gets the path of a datasheet given a name.
Public Function GetComponentDatasheetPath(strName As String) As String
    Dim strPath As String
    
    ' Check if the datasheet exists.
    strPath = GetDatasheetsDirectory() & strName & DATASHEET_EXT
    If Dir(strPath) <> vbNullString Then
        GetComponentDatasheetPath = strPath
    Else
        GetComponentDatasheetPath = vbNullString
    End If
End Function

' Checks if a component has a datasheet.
Public Function ComponentHasDatasheet(strName As String) As Boolean
    If GetComponentDatasheetPath(strName) <> vbNullString Then
        ComponentHasDatasheet = True
    Else
        ComponentHasDatasheet = False
    End If
End Function

' Opens a component datasheet file.
Public Sub OpenComponentDatasheet(strName As String)
    If ComponentHasDatasheet(strName) Then
        ShellExecute Screen.ActiveForm.hWnd, "open", GetComponentDatasheetPath(strName), _
            vbNullString, vbNullString, 1
    End If
End Sub

' Gets the path to the images directory.
Public Function GetImagesDirectory() As String
    GetImagesDirectory = GetWorkspacePath() & IMAGE_FOLDER
End Function

' Gets the path of an image given a name.
Public Function GetComponentImagePath(strName As String, strPackage As String) As String
    Dim strPath As String
    
    ' Check for image by component name.
    strPath = GetImagesDirectory() & strName & IMAGE_EXT
    If Dir(strPath, vbNormal) <> vbNullString Then
        GetComponentImagePath = strPath
        Exit Function
    End If
    
    ' Check for image by component package.
    strPath = GetImagesDirectory() & strPackage & IMAGE_EXT
    If Dir(strPath, vbNormal) <> vbNullString Then
        GetComponentImagePath = strPath
        Exit Function
    End If
    
    GetComponentImagePath = vbNullString
End Function

