VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComponentProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' clsComponentProperty
''' Abstraction of a PartCat component property.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_lngID As Long
Private m_strName As String
Private m_strValue As String
Private m_lngComponentID As Long

' Class constructor.
Private Sub Class_Initialize()
    ID = -1
    ComponentID = -1
    Key = ""
    Value = ""
End Sub

' Saves the property to the database.
Public Sub Save()
    ' Check if we have a valid component ID first.
    If ComponentID = -1 Then
        MsgBox "Property '" & Key & "' has no Component ID associated.", _
            vbOKOnly + vbCritical, "Can't Save Property"
        Exit Sub
    End If
    
    ' Save the property.
    ID = SaveProperty(ID, Key, Value, ComponentID)
End Sub

' Deletes a property from the database.
Public Sub Delete()
    DeleteProperty ID
    ID = -1
End Sub

' Check if the property is empty.
Public Function IsEmpty() As Boolean
    IsEmpty = (Key = "") And (Value = "")
End Function

' Check if a property is valid.
Public Function IsValid() As Boolean
    Dim strLine As String
    
    ' Check if it's empty.
    If IsEmpty Then
        IsValid = False
        Exit Function
    End If
    
    ' I can't remember how/why this works. Leave intact.
    strLine = LineFormat
    If InStr(strLine, ":") <> 0 Then
        If Right(strLine, 1) <> "-" Then
            IsValid = True
            Exit Function
        End If
    End If

    IsValid = False
End Function

' Get the old style property line.
Public Function LineFormat() As String
    LineFormat = Key & ": " & Value
End Function

' Property ID getter.
Public Property Get ID() As Long
    ID = m_lngID
End Property

' Property ID setter.
Public Property Let ID(lngID As Long)
    m_lngID = lngID
End Property

' Property name getter.
Public Property Get Key() As String
    Key = m_strName
End Property

' Property name setter.
Public Property Let Key(strName As String)
    m_strName = strName
End Property

' Property value getter.
Public Property Get Value() As String
    Value = m_strValue
End Property

' Property value setter.
Public Property Let Value(strValue As String)
    m_strValue = strValue
End Property

' Property component ID getter.
Public Property Get ComponentID() As Long
    ComponentID = m_lngComponentID
End Property

' Property component ID setter.
Public Property Let ComponentID(lngComponentID As Long)
    m_lngComponentID = lngComponentID
End Property
