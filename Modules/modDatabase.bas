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
    For intIndex = 1 To grdProperties.Rows - 2
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

' Saves/creates a component to the database. For component creation lngID should be -1.
Public Function SaveComponent(lngID As Long, strName As String, strQuantity As String, _
        strNotes As String, lngCategoryID As Long, lngSubCategoryID As Long, _
        lngPackageID As Long, strProperties As String) As Long
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    If lngID = -1 Then
        ' Create the component.
        stmt.Create "INSERT INTO Components (Name, Quantity, Notes, CategoryID, " & _
            "SubCategoryID, PackageID, Properties) VALUES ('[Name]', [Quantity], " & _
            "'[Notes]', [CategoryID], [SubCategoryID], [PackageID], '[Properties]')"
    Else
        ' Update an existing component.
        stmt.Create "UPDATE Components SET Name = '[Name]', Quantity = [Quantity], " & _
            "Notes = '[Notes]', CategoryID = [CategoryID], SubCategoryID = [SubCategoryID], " & _
            "PackageID = [PackageID], Properties = '[Properties]' WHERE ID = [ID]"
        stmt.Parameter("ID") = lngID
    End If
    
    ' Add parameters.
    stmt.Parameter("Name") = strName
    stmt.Parameter("Quantity") = strQuantity
    stmt.Parameter("Notes") = strNotes
    stmt.Parameter("CategoryID") = lngCategoryID
    stmt.Parameter("SubCategoryID") = lngSubCategoryID
    stmt.Parameter("PackageID") = lngPackageID
    stmt.Parameter("Properties") = strProperties
    
    ' Execute the operation.
    m_adoConnection.Execute stmt.Statement
    
    ' Get the component ID.
    If lngID = -1 Then
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        ' Get the newly added component ID.
        Set rs = m_adoConnection.Execute("SELECT @@IDENTITY FROM Components")
        If Not rs.EOF Then
            SaveComponent = rs(0)
        Else
            SaveComponent = -1
        End If
        
        ' Clean up recordset.
        rs.Close
        Set rs = Nothing
    Else
        SaveComponent = lngID
    End If
    
    ' Close the connection.
    CloseConnection
End Function

' Deletes a component from the database.
Public Sub DeleteComponent(lngID As Long, Optional strName As String = vbNullString)
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    stmt.Create "DELETE * FROM Components Where ID = [ID]"
    stmt.Parameter("ID") = lngID
    
    ' Execute the operation close the connection.
    m_adoConnection.Execute stmt.Statement
    CloseConnection
    
    ' Delete the component datasheet and image as well.
    If strName <> vbNullString Then
        DeleteComponentDatasheet strName
        DeleteComponentImage strName
    End If
End Sub

' Saves/creates a package to the database. For package creation lngID should be -1.
Public Function SavePackage(lngID As Long, strName As String) As Long
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    If lngID = -1 Then
        ' Create the package.
        stmt.Create "INSERT INTO Packages (Name) VALUES ('[Name]')"
    Else
        ' Update an existing package.
        stmt.Create "UPDATE Packages SET Name = '[Name]' WHERE ID = [ID]"
        stmt.Parameter("ID") = lngID
    End If
    
    ' Add parameters and execute the operation.
    stmt.Parameter("Name") = strName
    m_adoConnection.Execute stmt.Statement
    
    ' Get the component ID.
    If lngID = -1 Then
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        ' Get the newly added package ID.
        Set rs = m_adoConnection.Execute("SELECT @@IDENTITY FROM Packages")
        If Not rs.EOF Then
            SavePackage = rs(0)
        Else
            SavePackage = -1
        End If
        
        ' Clean up recordset.
        rs.Close
        Set rs = Nothing
    Else
        SavePackage = lngID
    End If
    
    ' Close the connection.
    CloseConnection
End Function

' Deletes a package from the database.
Public Sub DeletePackage(lngID As Long)
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    stmt.Create "DELETE * FROM Packages Where ID = [ID]"
    stmt.Parameter("ID") = lngID
    
    ' Execute the operation close the connection.
    m_adoConnection.Execute stmt.Statement
    CloseConnection
End Sub

' Loads a component by its ID and populates a form.
Public Function LoadComponentDetail(lngID As Long, frmForm As frmComponent) As Boolean
    Dim rs As ADODB.Recordset
    Dim stmt As SQLStatement
    
    ' Initialize the objects.
    Set rs = New ADODB.Recordset
    Set stmt = New SQLStatement
    
    ' Open the database and query it.
    OpenConnection
    stmt.Create "SELECT * FROM Components WHERE ID = [ID]"
    stmt.Parameter("ID") = lngID
    rs.Open stmt.Statement, m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
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
    Dim stmt As SQLStatement
    
    ' Initialize the objects.
    Set rs = New ADODB.Recordset
    Set stmt = New SQLStatement
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    stmt.Create "SELECT ID, Name FROM SubCategories WHERE ParentID = [ParentID] " & _
        "ORDER BY Name ASC"
    stmt.Parameter("ParentID") = lngCatID
    rs.Open stmt.Statement, m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
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
    Dim stmt As SQLStatement
    
    ' Initialize the objects.
    Set rs = New ADODB.Recordset
    Set stmt = New SQLStatement
    
    ' Clear the list.
    lstBox.Clear
    
    ' Open the database and query it.
    OpenConnection
    stmt.Create "SELECT ID, Name FROM Components WHERE CategoryID = [CategoryID] " & _
        " AND SubCategoryID = [SubCategoryID] ORDER BY Name ASC"
    stmt.Parameter("CategoryID") = lngCatID
    stmt.Parameter("SubCategoryID") = lngSubCatID
    rs.Open stmt.Statement, m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
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
