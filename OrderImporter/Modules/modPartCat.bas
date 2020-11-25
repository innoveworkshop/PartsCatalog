Attribute VB_Name = "modPartCat"
''' modPartCat
''' A PartCat helper module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Components array.
Private m_arrComponents() As component
Private m_idxLastComponent As Long

' Initializes the components array.
Public Sub InitializeComponentsArray(lngSize As Long)
    ReDim m_arrComponents(lngSize)
    m_idxLastComponent = -1
End Sub

' Adds a component to the array.
Public Sub AddComponent(strName As String, strNotes As String, _
                        ByRef astrProperties() As String, _
                        lngQuantity As Long, strSearchCode As String)
    ' Increment the last component index and instantiate a new component.
    m_idxLastComponent = m_idxLastComponent + 1
    Set m_arrComponents(m_idxLastComponent) = New component
    
    ' Set the component attributes.
    With m_arrComponents(m_idxLastComponent)
        .Name = strName
        .Notes = strNotes
        .Quantity = lngQuantity
        .SearchCode = strSearchCode
        .Exported = False
    End With
    
    ' Set the component properties.
    Dim intIndex As Integer
    m_arrComponents(m_idxLastComponent).InitializeProperties UBound(astrProperties)
    If UBound(astrProperties) > 0 Then
        ' Set properties.
        For intIndex = 0 To UBound(astrProperties)
            m_arrComponents(m_idxLastComponent).Property(intIndex) = astrProperties(intIndex)
        Next intIndex
        
        ' Remove the invalid component properties.
        m_arrComponents(m_idxLastComponent).RemoveInvalidProperties
    End If
End Sub

' Gets a component from the components array.
Public Function GetComponent(lngIndex As Long) As component
    Set GetComponent = m_arrComponents(lngIndex)
End Function

' Gets the number of components in the array.
Public Function LastComponentIndex() As Long
    LastComponentIndex = m_idxLastComponent
End Function
