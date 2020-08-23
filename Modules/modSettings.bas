Attribute VB_Name = "modSettings"
''' modSettings
''' Handles the applications settings.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private constants.
Private Const COMPANY_NAME As String = "Innove Workshop"
Private Const APP_NAME As String = "Parts Catalogger"

' Gets the last used database path. Returns vbNullString if there wasn't one or it is
' no longer valid.
Public Function LastUsedDatabasePath() As String
    Dim strPath As String
    
    ' Get the setting and check if it exists.
    strPath = ReadSetting("LastUsedDatabase", vbNullString)
    If strPath = vbNullString Then
        LastUsedDatabasePath = vbNullString
        Exit Function
    End If
    
    ' Check if the file actually exists.
    If Dir(strPath) = vbNullString Then
        LastUsedDatabasePath = vbNullString
        Exit Function
    End If
    
    ' Return the lasr used database.
    LastUsedDatabasePath = strPath
End Function

' Sets the last used database setting.
Public Sub SetLastUsedDatabasePath(strPath As String)
    WriteSetting "LastUsedDatabase", strPath
End Sub

' Writes a setting to the registry.
Private Sub WriteSetting(strKey As String, strValue As String)
    SaveSetting COMPANY_NAME, APP_NAME, strKey, strValue
End Sub

' Reads a setting entry from the registry.
Private Function ReadSetting(strKey As String, _
                             Optional strDefault As String = vbNullString) As String
    ReadSetting = GetSetting(COMPANY_NAME, APP_NAME, strKey, strDefault)
End Function
