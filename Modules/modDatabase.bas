Attribute VB_Name = "modDatabase"
''' modDatabase
''' Database helper module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_strDatabasePath As String
Private m_strWorkspacePath As String
Private m_adoConnection As ADODB.Connection

' Checks if a database is associated.
Public Function IsDatabaseAssociated() As Boolean
    If m_strDatabasePath <> vbNullString Then
        IsDatabaseAssociated = True
    Else
        IsDatabaseAssociated = False
    End If
End Function

' Clears any database association.
Public Sub ClearDatabasePath()
    m_strDatabasePath = vbNullString
    m_strWorkspacePath = vbNullString
End Sub

' Sets the database path.
Public Sub SetDatabasePath(strPath As String)
    m_strDatabasePath = strPath
    m_strWorkspacePath = Left(strPath, InStrRev(strPath, "\"))
End Sub

' Gets the workspace path.
Public Function GetWorkspacePath() As String
    GetWorkspacePath = m_strWorkspacePath
End Function

' Converts a component properties grid into a tabbed properties string for the database.
Public Function ComponentTabbedGridProperties(grdProperties As MSFlexGrid) As String
    Dim strBuffer As String
    Dim intIndex As Integer
    
    ' Go through rows appending them to the string.
    For intIndex = 1 To grdProperties.Rows - 1
        ' Append the tab separator.
        If intIndex > 1 Then
            strBuffer = strBuffer & vbTab
        End If
        
        ' Append property.
        strBuffer = strBuffer & grdProperties.TextMatrix(intIndex, 0) & ": " & _
            grdProperties.TextMatrix(intIndex, 1)
    Next intIndex
    
    ComponentTabbedGridProperties = strBuffer
End Function

' Loads a component by its ID and populates a form.
Public Function LoadComponentDetail(lngID As Long, frmForm As frmComponent) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Open the database and query it.
    OpenConnection
    rs.Open "SELECT * FROM Components WHERE ID = " & lngID, _
        m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    If Not rs.EOF Then
        frmForm.PopulateFromRecordset rs
        LoadComponentDetail = True
    Else
        LoadComponentDetail = False
    End If
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    CloseConnection
End Function

' Load categories.
Public Sub LoadCategories(lstBox As Variant, Optional blnCloseExit As Boolean = True)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    rs.Open "SELECT ID, Name FROM Categories ORDER BY Name ASC", _
        m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Load sub-categories based on the parent category ID.
Public Sub LoadSubCategories(lngCatID As Long, lstBox As Variant, _
        Optional blnCloseExit As Boolean = True)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    rs.Open "SELECT ID, Name FROM SubCategories WHERE ParentID = " & lngCatID & _
        " ORDER BY Name ASC", m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Load components based on their category and sub-category IDs.
Public Sub LoadComponents(lngCatID As Long, lngSubCatID As Long, lstBox As Variant, _
        Optional blnCloseExit As Boolean = True)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    rs.Open "SELECT ID, Name FROM Components WHERE CategoryID = " & lngCatID & _
        " AND SubCategoryID = " & lngSubCatID & " ORDER BY Name ASC", _
        m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Load packages.
Public Sub LoadPackages(lstBox As Variant, Optional blnCloseExit As Boolean = True)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    rs.Open "SELECT ID, Name FROM Packages ORDER BY Name ASC", _
        m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    Do While Not rs.EOF
        lstBox.AddItem rs.Fields("Name")
        lstBox.ItemData(lstBox.NewIndex) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Opens a predefined database connection.
Private Sub OpenConnection()
    If Not m_adoConnection Is Nothing Then
        Exit Sub
    End If
    
    ' Setup connection.
    Set m_adoConnection = New ADODB.Connection
    m_adoConnection.Provider = "Microsoft.Jet.OLEDB.4.0"
    m_adoConnection.ConnectionString = "Data Source = " & m_strDatabasePath & ";"
    
    ' Open it.
    m_adoConnection.Open
End Sub

' Closes the default database connection.
Private Sub CloseConnection()
    m_adoConnection.Close
    Set m_adoConnection = Nothing
End Sub
