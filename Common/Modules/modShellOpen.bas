Attribute VB_Name = "modShellOpen"
''' modShellOpen
''' Helper module to easily open files and execute programs.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Win32 API constants.
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

' Win32 API imports.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hWnd As Long, ByVal lpOperation As String, _
                    ByVal lpFile As String, ByVal lpParameters As String, _
                    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Open a file with the default application.
Public Sub OpenFile(hwndParent As Long, strPath As String)
    ShellExecute hwndParent, "open", strPath, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

' Executes an application.
Public Sub RunExecutable(hwndParent As Long, strExePath As String, _
        Optional strParameters As String = vbNullString, _
        Optional strWorkingDir As String = vbNullString, _
        Optional lngShowCmd As Long = SW_SHOWNORMAL)
    ShellExecute hwndParent, "open", strExePath, strParameters, strWorkingDir, lngShowCmd
End Sub

' Open a URL with the default browser.
Public Function OpenURL(strURL As String) As Long
    OpenURL = ShellExecute(0, "open", strURL, vbNullString, vbNullString, SW_SHOWNORMAL)
End Function
