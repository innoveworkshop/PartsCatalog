VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' SQLStatement
''' A class to build SQL statements safely and easily.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_strStatement As String

' Creates a new statement from scratch.
Public Sub Create(strStatement As String)
    m_strStatement = strStatement
End Sub

' Sanitizes a parameter input.
Private Function SanitizeInput(varInput As Variant) As Variant
    ' Do nothing if the input isn't a string.
    If VarType(varInput) <> vbString Then
        SanitizeInput = varInput
        Exit Function
    End If
    
    SanitizeInput = "'" & Replace(varInput, "'", "''") & "'"
End Function

' Statement string getter.
Public Property Get Statement() As String
    Statement = m_strStatement
End Property

' Statement parameter setter.
Public Property Let Parameter(strKey As String, varValue As Variant)
    If InStr(m_strStatement, strKey) = 0 Then
        MsgBox "No parameter named '" & strKey & "' in the prepared SQL statement.", _
            vbOKOnly + vbCritical, "SQL Statement Preparation Error"
        Exit Property
    End If

    ' Handle that pesky Null value.
    If IsNull(varValue) Then
        m_strStatement = Replace(m_strStatement, "[" & strKey & "]", "NULL")
        Exit Property
    End If

    m_strStatement = Replace(m_strStatement, "[" & strKey & "]", _
        SanitizeInput(varValue))
End Property
