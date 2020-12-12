VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Component"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' clsComponent
''' Abstraction of a PartCat component.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Properties.
Private m_strName As String
Private m_lngQuantity As Long
Private m_strNotes As String
Private m_strDatasheet As String
Private m_aProperties() As ComponentProperty
Private m_intCategoryID As Integer
Private m_intSubCategoryID As Integer
Private m_intPackageID As Integer
Private m_lngPartID As Long
Private m_strSearchCode As String
Private m_blnExported As Boolean

' Class initializer.
Private Sub Class_Initialize()
    PartID = -1
    CategoryID = -1
    SubCategoryID = -1
    PackageID = -1
End Sub

' Initializes the properties array.
Public Sub InitializeProperties(intSize As Integer)
    Dim intIndex As Integer
    ReDim m_aProperties(intSize)
    
    ' Initialize the array with new properties.
    For intIndex = 0 To UBound(m_aProperties)
        Set m_aProperties(intIndex) = New ComponentProperty
    Next intIndex
End Sub

' Export the component to the database and download the datasheet.
Public Sub Export()
    ' Check if we have an opened database.
    If Not IsDatabaseAssociated Then
        MsgBox "There isn't a database currently opened. Open one before " & _
            "trying to import a component.", vbOKOnly + vbExclamation, _
            "No Database Associated"
        Exit Sub
    End If
    
    ' Save the component to the database.
    PartID = SaveComponent(PartID, Name, CStr(Quantity), Notes, CategoryID, _
        SubCategoryID, PackageID)
    
    ' Save component properties to the database.
    If Not m_aProperties(0).IsEmpty Then
        Dim intIndex As Integer
        For intIndex = 0 To UBound(m_aProperties)
            m_aProperties(intIndex).ComponentID = PartID
            m_aProperties(intIndex).Save
        Next intIndex
    End If
    
    ' Download datasheet.
    If Datasheet <> "" Then
        If Not DownloadComponentDatasheet(Name, Datasheet) Then
            MsgBox "Failed to download the component datasheet.", _
                vbOKOnly + vbExclamation, "Download Failed"
        
            Exported = False
            Exit Sub
        End If
    End If
    
    Exported = True
End Sub

' Adds a property to the array.
Public Sub AddProperty(strKey As String, strValue As String)
    If (UBound(m_aProperties) > 0) Or (Not m_aProperties(0).IsEmpty) Then
        ReDim Preserve m_aProperties(UBound(m_aProperties) + 1)
    End If
    
    Set m_aProperties(UBound(m_aProperties)) = New ComponentProperty
    m_aProperties(UBound(m_aProperties)).Key = strKey
    m_aProperties(UBound(m_aProperties)).Value = strValue
    m_aProperties(UBound(m_aProperties)).ComponentID = PartID
End Sub

' Edit a property from the array.
Public Sub EditProperty(strKey As String, strValue As String)
    Dim intIndex As Integer
    
    ' Go through the properties looking for the specified key.
    For intIndex = 0 To UBound(m_aProperties)
        If InStr(m_aProperties(intIndex).Key, strKey) Then
            m_aProperties(intIndex).Key = strKey
            m_aProperties(intIndex).Value = strValue
            m_aProperties(intIndex).ComponentID = PartID
            Exit Sub
        End If
    Next intIndex
    
    ' Key wasn't found.
    MsgBox "The property key you want to edit wasn't found.", _
        vbOKOnly + vbCritical, "Property Key Not Found"
End Sub

' Deletes a property from the array.
Public Sub DeleteProperty(strKey As String)
    Dim astrBuffer() As String
    Dim intIndex As Integer
    Dim intOffset As Integer
    
    ' Check if we are deleting the last available property.
    If UBound(m_aProperties) = 0 Then
        ReDim m_aProperties(0)
        Set m_aProperties(0) = New ComponentProperty
        Exit Sub
    End If
    
    ' Populate the buffer array.
    intOffset = 0
    ReDim apropBuffer(UBound(m_aProperties) - 1)
    For intIndex = 0 To UBound(m_aProperties)
        ' Check if we should delete the property.
        If InStr(m_aProperties(intIndex).Key, strKey) Then
            ' Add offset to the index.
            intOffset = -1
        Else
            If intIndex = 0 Then
                intOffset = 0
            End If
            
            ' Copy a non-deleted property
            Set apropBuffer(intIndex + intOffset) = m_aProperties(intIndex)
        End If
    Next intIndex
    
    ' Copy the buffer to the actual properties array.
    ReDim m_aProperties(UBound(apropBuffer))
    For intIndex = 0 To UBound(m_aProperties)
        Set m_aProperties(intIndex) = apropBuffer(intIndex)
    Next intIndex
End Sub

' Removes invalid properties from the properties array.
Public Sub RemoveInvalidProperties()
    Dim apropTemp() As ComponentProperty
    Dim intCountValid As Integer
    Dim intIndex As Integer
    
    ' Count valid properties.
    intCountValid = 0
    For intIndex = 0 To UBound(m_aProperties)
        If m_aProperties(intIndex).IsValid Then
            intCountValid = intCountValid + 1
        End If
    Next intIndex
    
    ' Go through adding valid items to the array.
    apropTemp = m_aProperties
    ReDim m_aProperties(intCountValid - 1)
    Dim intNewIndex As Integer
    intNewIndex = 0
    For intIndex = 0 To UBound(apropTemp)
        If apropTemp(intIndex).IsValid Then
            Set m_aProperties(intNewIndex) = apropTemp(intIndex)
            intNewIndex = intNewIndex + 1
        End If
    Next intIndex
End Sub

' Removes invalid file characters from a string, making it filename-safe.
Public Static Function RemoveInvalidFileCharacters(strName As String) As String
    Dim strBuffer As String
    strBuffer = strName
    
    ' Remove invalid characters.
    strBuffer = Replace(strBuffer, "<", "")
    strBuffer = Replace(strBuffer, ">", "")
    strBuffer = Replace(strBuffer, ":", "")
    strBuffer = Replace(strBuffer, """", "")
    strBuffer = Replace(strBuffer, "/", "")
    strBuffer = Replace(strBuffer, "\", "")
    strBuffer = Replace(strBuffer, "|", "")
    strBuffer = Replace(strBuffer, "?", "")
    strBuffer = Replace(strBuffer, "*", "")
    
    ' Trim whitespace and return.
    RemoveInvalidFileCharacters = Trim(strBuffer)
End Function

' Name getter.
Public Property Get Name() As String
    Name = m_strName
End Property

' Name setter.
Public Property Let Name(strName As String)
    m_strName = RemoveInvalidFileCharacters(strName)
End Property

' Quantity getter.
Public Property Get Quantity() As Long
    Quantity = m_lngQuantity
End Property

' Quantity setter.
Public Property Let Quantity(lngQuantity As Long)
    m_lngQuantity = lngQuantity
End Property

' Notes getter.
Public Property Get Notes() As String
    Notes = m_strNotes
End Property

' Notes setter.
Public Property Let Notes(strNotes As String)
    m_strNotes = strNotes
End Property

' Database ID getter.
Public Property Get PartID() As Long
    PartID = m_lngPartID
End Property

' Database ID setter.
Public Property Let PartID(lngPartID As Long)
    m_lngPartID = lngPartID
End Property

' Category ID getter.
Public Property Get CategoryID() As Integer
    CategoryID = m_intCategoryID
End Property

' Category ID setter.
Public Property Let CategoryID(intCategoryID As Integer)
    m_intCategoryID = intCategoryID
End Property

' Sub-Category ID getter.
Public Property Get SubCategoryID() As Integer
    SubCategoryID = m_intSubCategoryID
End Property

' Sub-Category ID setter.
Public Property Let SubCategoryID(intSubCategoryID As Integer)
    m_intSubCategoryID = intSubCategoryID
End Property

' Package ID getter.
Public Property Get PackageID() As Integer
    PackageID = m_intPackageID
End Property

' Package ID setter.
Public Property Let PackageID(intPackageID As Integer)
    m_intPackageID = intPackageID
End Property

' Datasheet getter.
Public Property Get Datasheet() As String
    Datasheet = m_strDatasheet
End Property

' Datasheet setter.
Public Property Let Datasheet(strDatasheet As String)
    m_strDatasheet = strDatasheet
End Property

' Search code getter.
Public Property Get SearchCode() As String
    SearchCode = m_strSearchCode
End Property

' Search code setter.
Public Property Let SearchCode(strSearchCode As String)
    m_strSearchCode = strSearchCode
End Property

' Exported getter.
Public Property Get Exported() As Boolean
    Exported = m_blnExported
End Property

' Exported setter.
Public Property Let Exported(blnExported As Boolean)
    m_blnExported = blnExported
End Property

' Property getter.
Public Property Get Property(intIndex As Integer) As String
    Property = m_aProperties(intIndex).LineFormat
End Property

' Property setter.
Public Property Let Property(intIndex As Integer, strProperty As String)
    Dim astrProperty() As String
    astrProperty = Split(strProperty, ": ")
    
    ' Looks like this will just be another invalid property...
    If UBound(astrProperty) < 1 Then
        Set m_aProperties(intIndex) = New ComponentProperty
        Exit Property
    End If
    
    ' Remove invalid characters from property key.
    astrProperty(0) = Replace(astrProperty(0), "  ", " ")
    astrProperty(0) = Replace(astrProperty(0), "'", "")
    astrProperty(0) = Replace(astrProperty(0), "(", "")
    astrProperty(0) = Replace(astrProperty(0), ")", "")
    astrProperty(0) = Replace(astrProperty(0), "@", "")
    astrProperty(0) = Replace(astrProperty(0), "%", "")
    astrProperty(0) = Replace(astrProperty(0), ".", "")
    
    Set m_aProperties(intIndex) = New ComponentProperty
    m_aProperties(intIndex).Key = astrProperty(0)
    m_aProperties(intIndex).Value = astrProperty(1)
End Property

' Properties getter.
Public Property Get Properties() As ComponentProperty()
    Properties = m_aProperties
End Property
