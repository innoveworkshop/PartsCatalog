Attribute VB_Name = "modSettings"
''' modSettings
''' Handles the applications settings.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private constants.
Private Const COMPANY_NAME As String = "Innove Workshop"
Private Const APP_NAME As String = "Parts Catalogger"
Private Const ORDERIMPORTER_EXE As String = "OrderImporter.exe"

' Registry key names.
Private Const KEY_LASTUSEDDATABASE As String = "LastUsedDatabase"
Private Const KEY_ORDERIMPORTERPATH As String = "OrderImporterPath"

' Gets the last used database path. Returns vbNullString if there wasn't one or it is
' no longer valid.
Public Function LastUsedDatabasePath() As String
    Dim strPath As String
    
    ' Get the setting and check if it exists.
    strPath = ReadSetting(KEY_LASTUSEDDATABASE, vbNullString)
    If strPath = vbNullString Then
        LastUsedDatabasePath = vbNullString
        Exit Function
    End If
    
    ' Check if the file actually exists.
    If Dir(strPath) = vbNullString Then
        LastUsedDatabasePath = vbNullString
        Exit Function
    End If
    
    ' Return the last used database.
    LastUsedDatabasePath = strPath
End Function

' Sets the last used database setting.
Public Sub SetLastUsedDatabasePath(strPath As String)
    WriteSetting KEY_LASTUSEDDATABASE, strPath
End Sub

' Gets the Order Importer application path. Returns vbNullString if not available.
Public Function OrderImporterPath() As String
    Dim strPath As String
    
    ' Check if the application is in the same folder as the executable.
    strPath = App.Path & "\" & ORDERIMPORTER_EXE
    If Dir(strPath) <> vbNullString Then
        OrderImporterPath = strPath
        Exit Function
    End If
    
    ' Check if there's something in the registry.
    strPath = ReadSetting(KEY_ORDERIMPORTERPATH, vbNullString)
    If strPath <> vbNullString Then
        ' Check if the executable is still there.
        If Dir(strPath) <> vbNullString Then
            OrderImporterPath = strPath
            Exit Function
        End If
    End If
    
    ' Looks like we didn't find anything.
    OrderImporterPath = vbNullString
End Function

' Sets the Order Importer application path.
Public Sub SetOrderImporterPath(strPath As String)
    WriteSetting KEY_ORDERIMPORTERPATH, strPath
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
