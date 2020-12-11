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

' Gets the database path.
Public Function GetDatabasePath() As String
    GetDatabasePath = m_strDatabasePath
End Function

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
            "SubCategoryID, PackageID, Properties) VALUES ([Name], [Quantity], " & _
            "[Notes], [CategoryID], [SubCategoryID], [PackageID], [Properties])"
    Else
        ' Update an existing component.
        stmt.Create "UPDATE Components SET Name = [Name], Quantity = [Quantity], " & _
            "Notes = [Notes], CategoryID = [CategoryID], SubCategoryID = [SubCategoryID], " & _
            "PackageID = [PackageID], Properties = [Properties] WHERE ID = [ID]"
        stmt.Parameter("ID") = lngID
    End If
    
    ' Add parameters.
    stmt.Parameter("Name") = strName
    stmt.Parameter("Quantity") = CLng(strQuantity)
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

' Updates a component quantity.
Public Function UpdateComponentQuantity(lngID As Long, lngQuantity As Long) As Long
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    stmt.Create "UPDATE Components SET Quantity = [Quantity] WHERE ID = [ID]"
    stmt.Parameter("ID") = lngID
    stmt.Parameter("Quantity") = lngQuantity
    
    ' Execute the operation.
    m_adoConnection.Execute stmt.Statement
    
    ' Close the connection.
    CloseConnection
    UpdateComponentQuantity = lngID
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

' Saves/creates a property to the database. For property creation lngID should be -1.
Public Function SaveProperty(lngID As Long, strName As String, strValue As String, _
        lngComponentID As Long) As Long
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    If lngID = -1 Then
        ' Create the category.
        stmt.Create "INSERT INTO Properties ([Name], [Value], ComponentID) VALUES " & _
            "([_Name], [_Value], [ComponentID])"
    Else
        ' Update an existing category.
        stmt.Create "UPDATE Properties SET [Name] = [_Name], [Value] = [_Value], " & _
            "ComponentID = [ComponentID] WHERE [ID] = [_ID]"
        stmt.Parameter("_ID") = lngID
    End If
    
    ' Add parameters and execute the operation.
    stmt.Parameter("_Name") = strName
    stmt.Parameter("_Value") = strValue
    stmt.Parameter("ComponentID") = lngComponentID
    Debug.Print stmt.Statement
    m_adoConnection.Execute stmt.Statement
    
    ' Get the property ID.
    If lngID = -1 Then
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        ' Get the newly added property ID.
        Set rs = m_adoConnection.Execute("SELECT @@IDENTITY FROM Properties")
        If Not rs.EOF Then
            SaveProperty = rs(0)
        Else
            SaveProperty = -1
        End If
        
        ' Clean up recordset.
        rs.Close
        Set rs = Nothing
    Else
        SaveProperty = lngID
    End If
    
    ' Close the connection.
    CloseConnection
End Function

' Saves/creates a category to the database. For category creation lngID should be -1.
Public Function SaveCategory(lngID As Long, strName As String) As Long
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    If lngID = -1 Then
        ' Create the category.
        stmt.Create "INSERT INTO Categories (Name) VALUES ([Name])"
    Else
        ' Update an existing category.
        stmt.Create "UPDATE Categories SET Name = [Name] WHERE ID = [ID]"
        stmt.Parameter("ID") = lngID
    End If
    
    ' Add parameters and execute the operation.
    stmt.Parameter("Name") = strName
    m_adoConnection.Execute stmt.Statement
    
    ' Get the category ID.
    If lngID = -1 Then
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        ' Get the newly added category ID.
        Set rs = m_adoConnection.Execute("SELECT @@IDENTITY FROM Categories")
        If Not rs.EOF Then
            SaveCategory = rs(0)
        Else
            SaveCategory = -1
        End If
        
        ' Clean up recordset.
        rs.Close
        Set rs = Nothing
    Else
        SaveCategory = lngID
    End If
    
    ' Close the connection.
    CloseConnection
End Function

' Deletes a category from the database.
Public Sub DeleteCategory(lngID As Long)
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    stmt.Create "DELETE * FROM Categories Where ID = [ID]"
    stmt.Parameter("ID") = lngID
    
    ' Execute the operation close the connection.
    m_adoConnection.Execute stmt.Statement
    CloseConnection
End Sub

' Saves/creates a sub-category to the database. For sub-category creation lngID should be -1.
Public Function SaveSubCategory(lngID As Long, lngCategoryID, strName As String) As Long
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    If lngID = -1 Then
        ' Create the sub-category.
        stmt.Create "INSERT INTO SubCategories (Name, ParentID) VALUES ([Name], " & _
            "[ParentID])"
    Else
        ' Update an existing sub-category.
        stmt.Create "UPDATE SubCategories SET Name = [Name], ParentID = [ParentID] " & _
            "WHERE ID = [ID]"
        stmt.Parameter("ID") = lngID
    End If
    
    ' Add parameters and execute the operation.
    stmt.Parameter("Name") = strName
    stmt.Parameter("ParentID") = lngCategoryID
    m_adoConnection.Execute stmt.Statement
    
    ' Get the sub-category ID.
    If lngID = -1 Then
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        ' Get the newly added sub-category ID.
        Set rs = m_adoConnection.Execute("SELECT @@IDENTITY FROM SubCategories")
        If Not rs.EOF Then
            SaveSubCategory = rs(0)
        Else
            SaveSubCategory = -1
        End If
        
        ' Clean up recordset.
        rs.Close
        Set rs = Nothing
    Else
        SaveSubCategory = lngID
    End If
    
    ' Close the connection.
    CloseConnection
End Function

' Deletes a sub-category from the database.
Public Sub DeleteSubCategory(lngID As Long)
    Dim stmt As SQLStatement
    
    ' Open the database.
    OpenConnection
    
    ' Setup the statement.
    Set stmt = New SQLStatement
    stmt.Create "DELETE * FROM SubCategories Where ID = [ID]"
    stmt.Parameter("ID") = lngID
    
    ' Execute the operation close the connection.
    m_adoConnection.Execute stmt.Statement
    CloseConnection
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
        stmt.Create "INSERT INTO Packages (Name) VALUES ([Name])"
    Else
        ' Update an existing package.
        stmt.Create "UPDATE Packages SET Name = [Name] WHERE ID = [ID]"
        stmt.Parameter("ID") = lngID
    End If
    
    ' Add parameters and execute the operation.
    stmt.Parameter("Name") = strName
    m_adoConnection.Execute stmt.Statement
    
    ' Get the package ID.
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

' Searches for a componeny by name and return its ID if there is one. -1 otherwise.
Public Function FindExistingComponent(strName As String) As Long
    Dim rs As ADODB.Recordset
    Dim stmt As SQLStatement
    
    ' Initialize the objects.
    Set rs = New ADODB.Recordset
    Set stmt = New SQLStatement
    
    ' Open the database and query it.
    OpenConnection
    stmt.Create "SELECT ID FROM Components WHERE Name = [Name]"
    stmt.Parameter("Name") = strName
    rs.Open stmt.Statement, m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Return the component ID if found.
    If Not rs.EOF Then
        FindExistingComponent = rs.Fields("ID")
    Else
        FindExistingComponent = -1
    End If
    
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    CloseConnection
End Function

' Loads a component by its ID and populates a form.
Public Function LoadComponentDetail(lngID As Long, frmForm As Form) As Boolean
    Dim rs As ADODB.Recordset
    Dim stmt As SQLStatement
    
    ' Check if the specified Form is the right one.
    If (frmForm.Name <> "frmComponent") And _
            (frmForm.Name <> "frmDuplicateComponent") Then
        MsgBox "Specified form to have component loaded is not correct. Expected " & _
            "'frmComponent' got '" & frmForm.Name & "'.", vbOKOnly + vbCritical, _
            "Programming Error"
        
        LoadComponentDetail = False
        Exit Function
    End If
    
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

' Load properties into a MSFlexGrid.
Public Sub LoadProperties(lngComponentID As Long, grdGrid As MSFlexGrid, _
        Optional blnCloseExit As Boolean = True)
    Dim rs As ADODB.Recordset
    Dim stmt As SQLStatement
    
    ' Initialize the objects.
    Set rs = New ADODB.Recordset
    Set stmt = New SQLStatement
    
    ' Clear the grid.
    grdGrid.Rows = 2
    grdGrid.RowData(1) = -1
    grdGrid.TextMatrix(1, 0) = ""
    grdGrid.TextMatrix(1, 1) = ""
    
    ' Open the database and query it.
    OpenConnection
    stmt.Create "SELECT ID, Name, Value FROM Properties WHERE " & _
        "ComponentID = [ComponentID] ORDER BY Name ASC"
    stmt.Parameter("ComponentID") = lngComponentID
    rs.Open stmt.Statement, m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Populate list.
    grdGrid.Rows = 1
    Do While Not rs.EOF
        grdGrid.AddItem rs.Fields("Name") & vbTab & rs.Fields("Value")
        grdGrid.RowData(grdGrid.Rows - 1) = rs.Fields("ID")
        rs.MoveNext
    Loop
    
    ' Add empty row.
    grdGrid.AddItem "" & vbTab & ""
    grdGrid.RowData(grdGrid.Rows - 1) = -1

    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

' Loads a property into a pair of TextBoxes.
Public Sub LoadProperty(lngPropertyID As Long, txtName As TextBox, _
        txtValue As TextBox, Optional blnCloseExit As Boolean = True)
    Dim rs As ADODB.Recordset
    Dim stmt As SQLStatement
    
    ' Initialize the objects.
    Set rs = New ADODB.Recordset
    Set stmt = New SQLStatement
    
    ' Clear the text boxes.
    txtName.Text = ""
    txtValue.Text = ""
    
    ' Open the database and query it.
    OpenConnection
    stmt.Create "SELECT ID, Name, Value FROM Properties WHERE " & _
        "ID = [ID] ORDER BY Name ASC"
    stmt.Parameter("ID") = lngPropertyID
    rs.Open stmt.Statement, m_adoConnection, adOpenForwardOnly, adLockReadOnly
    
    ' Check if the property ID was valid.
    If rs.EOF Then
        MsgBox "There isn't a property with the ID of " & lngPropertyID, _
            vbOKOnly + vbCritical, "Invalid Property ID"
        GoTo Finish
    End If
    
    ' Populate the text boxes.
    txtName.Text = rs.Fields("Name")
    txtValue.Text = rs.Fields("Value")

Finish:
    ' Close recordset and connection.
    rs.Close
    Set rs = Nothing
    If blnCloseExit Then
        CloseConnection
    End If
End Sub

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
    ' Check if there's a database associated.
    If Not IsDatabaseAssociated Then
        MsgBox "Can't open a connection to the database because there isn't one associated.", _
            vbOKOnly + vbCritical, "Database Connection Error"
    End If
    
    ' Check if there's an opened conection already.
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
