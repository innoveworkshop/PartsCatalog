Attribute VB_Name = "modAssetsHandler"
''' modAssetsHandler
''' Handles all the assets.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

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
    GetComponentDatasheetPath = GetDatasheetsDirectory() & strName & _
        DATASHEET_EXT
End Function

' Gets the path to the images directory.
Public Function GetImagesDirectory() As String
    GetImagesDirectory = GetWorkspacePath() & IMAGE_FOLDER
End Function

' Gets the path of an image given a name.
Public Function GetComponentImagePath(strName As String, strPackage As String) As String
    Dim strPath As String
    
    ' Check for image by component name.
    strPath = GetImagesDirectory() & strName & IMAGE_EXT
    If Not Dir(strPath, vbNormal) = vbNullString Then
        GetComponentImagePath = strPath
        Exit Function
    End If
    
    ' Check for image by component package.
    strPath = GetImagesDirectory() & strPackage & IMAGE_EXT
    If Not Dir(strPath, vbNormal) = vbNullString Then
        GetComponentImagePath = strPath
        Exit Function
    End If
    
    GetComponentImagePath = vbNullString
End Function

