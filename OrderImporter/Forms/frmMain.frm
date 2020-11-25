VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Order Importer"
   ClientHeight    =   9120
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDatabase 
      Caption         =   "Database"
      Height          =   975
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtDatabaseLocation 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   6135
      End
      Begin VB.CommandButton cmdBrowseDatabase 
         Caption         =   "..."
         Height          =   315
         Left            =   6360
         TabIndex        =   36
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "PartsCatalog Database Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSComctlLib.ImageList imlButtons 
      Left            =   4800
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6852
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D0B4
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   4080
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Component"
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   6855
      Begin VB.PictureBox picFindExisting 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         FillColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2905
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   510
         Width           =   255
      End
      Begin VB.PictureBox picRefreshPackages 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         FillColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6520
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   510
         Width           =   255
      End
      Begin VB.PictureBox picRefreshSubCategories 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         FillColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6520
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   1110
         Width           =   255
      End
      Begin VB.PictureBox picRefreshCategories 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         FillColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   1110
         Width           =   255
      End
      Begin VB.ComboBox cmbPackage 
         Height          =   315
         Left            =   4920
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbSubCategory 
         Height          =   315
         Left            =   3720
         TabIndex        =   29
         Top             =   1080
         Width           =   2775
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Import Component"
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CheckBox chkExported 
         Caption         =   "Imported"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   24
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Top             =   5820
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   315
         Left            =   2400
         TabIndex        =   22
         Top             =   5820
         Width           =   375
      End
      Begin VB.TextBox txtItemNumber 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Text            =   "0"
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   315
         Left            =   600
         TabIndex        =   19
         Top             =   5820
         Width           =   375
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   5820
         Width           =   375
      End
      Begin VB.CommandButton cmdLoadWebsite 
         Caption         =   "Distributor Website"
         Height          =   555
         Left            =   5040
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtDatasheetURL 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   4695
      End
      Begin MSFlexGridLib.MSFlexGrid grdProperties 
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483644
         GridColorFixed  =   -2147483644
         HighLight       =   2
         GridLinesFixed  =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
      End
      Begin VB.TextBox txtNotes 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox txtQuantity 
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Package:"
         Height          =   255
         Left            =   4920
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Sub-Category:"
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblNumberItems 
         Alignment       =   2  'Center
         Caption         =   "/ 000"
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   5860
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Datasheet URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "Distributor Order"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
      Begin VB.ComboBox cmbDistributor 
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":13916
         Left            =   5040
         List            =   "frmMain.frx":1391D
         TabIndex        =   5
         Text            =   "Farnell"
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Load Order"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   6615
      End
      Begin VB.CommandButton cmdBrowseOrder 
         Caption         =   "..."
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtOrderLocation 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Distributor:"
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Order File Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mniFileLoadOrder 
         Caption         =   "&Load Order..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mniFileOpenDatabase 
         Caption         =   "&Open Database..."
      End
      Begin VB.Menu mniFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuComponent 
      Caption         =   "&Component"
      Begin VB.Menu mniComponentPrevious 
         Caption         =   "P&revious"
      End
      Begin VB.Menu mniComponentNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu mniComponentSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentAddProperty 
         Caption         =   "&Add Property..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mniComponentDeleteProperty 
         Caption         =   "&Delete Propety"
         Shortcut        =   ^D
      End
      Begin VB.Menu mniComponentSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentLoadWebsite 
         Caption         =   "Load &Website..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mniComponentSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentExport 
         Caption         =   "&Import"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mniHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmMain
''' Application's main form.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Browse for PartsCatalog database.
Public Sub OpenDatabaseFile(Optional strPath As String = vbNullString)
    Dim strSetPath As String
    strSetPath = strPath
    
    ' Check if we should use the open dialog.
    If strPath = vbNullString Then
        ' Setup open dialog.
        dlgCommon.DialogTitle = "Open Database"
        dlgCommon.DefaultExt = "mdb"
        dlgCommon.Filter = "Microsoft Access Databases (*.mdb)|*.mdb|All Files (*.*)|*.*"
        dlgCommon.FileName = ""
        
        ' Open the dialog and set the path.
        dlgCommon.ShowOpen
        strSetPath = dlgCommon.FileName
    End If
    
    ' Set the database path.
    If strSetPath <> vbNullString Then
        SetDatabasePath strSetPath
        txtDatabaseLocation.Text = strSetPath
        
        ' Populate database-dependent components.
        PopulateComboBoxes
    Else
        txtDatabaseLocation.Text = ""
    End If
End Sub

' Imports an order into the system.
Public Sub ImportOrder()
    ' Check if we have an opened database.
    If Not IsDatabaseAssociated Then
        MsgBox "There isn't a database currently opened. Open one before " & _
            "loading a new order.", vbOKOnly + vbExclamation, "No Database Associated"
        Exit Sub
    End If
    
    ' Parse order and populate the components array.
    ParseFarnellOrder txtOrderLocation.Text
    
    ' Update the components record counter and show the first component.
    lblNumberItems.Caption = "of " & LastComponentIndex
    ShowComponent 0
End Sub

' Populates the ComboBoxes with data from the database.
Private Sub PopulateComboBoxes()
    ' Check if we have an opened database.
    If Not IsDatabaseAssociated Then
        MsgBox "There isn't a database currently opened. Open one before " & _
            "trying to refresh the ComboBoxes.", vbOKOnly + vbExclamation, _
            "No Database Associated"
        Exit Sub
    End If
    
    ' Load data into them.
    LoadPackages cmbPackage
    LoadCategories cmbCategory
End Sub

' Check for existing component in the database and lets the user decide what to do. Returns True if an abortion is needed.
Private Function CheckDuplicates( _
        Optional blnShowNoDuplicateDialog As Boolean = False) As Boolean
    Dim lngID As Long
    
    ' Check if this is already an imported component.
    If GetCurrentComponent.Exported Then
        Exit Function
    End If
    
    ' Search for existing component.
    lngID = FindExistingComponent(txtName.Text)
    
    ' Nothing was found.
    If lngID = -1 Then
        If blnShowNoDuplicateDialog Then
            MsgBox "No existing component with this name was found.", _
                vbOKOnly + vbInformation, "Nothing to see here"
        End If
        
        CheckDuplicates = False
        Exit Function
    End If
    
    ' We've got one!
    frmDuplicateComponent.PositionBySide Me, fraComponent
    LoadComponentDetail lngID, frmDuplicateComponent
    frmDuplicateComponent.Show vbModal, Me
    
    ' Check what sort of action we need to take.
    If frmDuplicateComponent.DoAction = actImportAnyway Then
        ' Import this thing anyway.
        ImportCurrentComponent
    ElseIf frmDuplicateComponent.DoAction = actUpdateQuantity Then
        ' Update the quantity, we are just restocking.
        UpdateComponentQuantity frmDuplicateComponent.ID, _
            CLng(txtQuantity.Text) + frmDuplicateComponent.Quantity
        
        ' Make sure it's marked as exported.
        GetCurrentComponent.Exported = True
        RefreshCurrentComponent
    Else
        ' Abort any subsequent import operations.
        CheckDuplicates = False
    End If
    
    CheckDuplicates = True
End Function

' Imports the current component into the database.
Private Sub ImportCurrentComponent()
    Dim component As component
    
    ' Check if we have an opened database.
    If Not IsDatabaseAssociated Then
        MsgBox "There isn't a database currently opened. Open one before " & _
            "trying to import a component.", vbOKOnly + vbExclamation, _
            "No Database Associated"
        Exit Sub
    End If
    
    ' Check if there's a component selected.
    If Not fraComponent.Enabled Then
        MsgBox "There isn't a component selected. We can't import this.", _
            vbOKOnly + vbCritical, "No Component Selected"
        Exit Sub
    End If
    
    ' Don't forget to save any changes and get the current component as well.
    SaveCurrentComponent
    Set component = GetCurrentComponent
    
    ' Set the component as exported.
    component.Export
    RefreshCurrentComponent
    
    ' Give the user some feedback.
    If component.Exported Then
        MsgBox component.Name & " imported successfully.", vbOKOnly + vbInformation, _
            "Component Exported"
    Else
        MsgBox component.Name & " import failed.", vbOKOnly + vbCritical, _
            "Failed to Export Component"
    End If
End Sub

' Shows a component by its index.
Public Sub ShowComponent(lngIndex As Long)
    ' Get the component.
    Dim component As component
    Set component = GetComponent(lngIndex)
    
    ' Set text fields.
    txtItemNumber.Text = CStr(lngIndex)
    txtName.Text = component.Name
    txtQuantity.Text = CStr(component.Quantity)
    txtNotes.Text = component.Notes
    txtDatasheetURL.Text = component.Datasheet
    
    ' Set comboboxes.
    SelectListItemByItemData cmbCategory, component.CategoryID
    SelectListItemByItemData cmbSubCategory, component.SubCategoryID
    SelectListItemByItemData cmbPackage, component.PackageID
    
    ' Set the exported checkbox.
    If component.Exported Then
        chkExported.Value = vbChecked
    Else
        chkExported.Value = vbUnchecked
    End If
    
    ' Preparate the grid for the properties.
    grdProperties.Rows = UBound(component.Properties) + 2
    
    ' Populate the properties.
    Dim intIndex As Integer
    Dim astrProperty() As String
    For intIndex = 0 To UBound(component.Properties)
        ' Check if the property is populated.
        If component.Property(intIndex) <> "" Then
            astrProperty = Split(component.Property(intIndex), ": ")
            grdProperties.TextMatrix(intIndex + 1, 0) = astrProperty(0)
            grdProperties.TextMatrix(intIndex + 1, 1) = astrProperty(1)
        Else
            grdProperties.TextMatrix(1, 0) = ""
            grdProperties.TextMatrix(1, 1) = ""
        End If
    Next intIndex
    
    ' Enable the component panel for editing and menu item.
    fraComponent.Enabled = True
    mnuComponent.Enabled = True
End Sub

' Gets the currently selected component.
Public Function GetCurrentComponent() As component
    Set GetCurrentComponent = GetComponent(CLng(txtItemNumber.Text))
End Function

' Refreshes the current component data.
Public Sub RefreshCurrentComponent()
    ShowComponent CLng(txtItemNumber.Text)
End Sub

' Saves the text fields to the current component.
Public Sub SaveCurrentComponent()
    Dim component As component
    Set component = GetCurrentComponent
    
    ' Save the component text fields.
    component.Name = txtName.Text
    component.Notes = txtNotes.Text
    component.Datasheet = txtDatasheetURL.Text
    component.Quantity = CLng(txtQuantity.Text)
    
    ' Save the category.
    If cmbCategory.ListIndex <> -1 Then
        component.CategoryID = cmbCategory.ItemData(cmbCategory.ListIndex)
    Else
        component.CategoryID = -1
    End If
    
    ' Save the sub-category.
    If cmbSubCategory.ListIndex <> -1 Then
        component.SubCategoryID = cmbSubCategory.ItemData(cmbSubCategory.ListIndex)
    Else
        component.SubCategoryID = -1
    End If
    
    ' Save the package.
    If cmbPackage.ListIndex <> -1 Then
        component.PackageID = cmbPackage.ItemData(cmbPackage.ListIndex)
    Else
        component.PackageID = -1
    End If
End Sub

' Selects a ComboBox item based on its ItemData ID.
Public Sub SelectListItemByItemData(cmbBox As ComboBox, intItemData As Integer, _
        Optional blnMinusOneDeselects As Boolean = True)
    Dim intIndex As Integer
    
    ' Check if we should treat the -1 ItemData as a deselection.
    If (intItemData = -1) And blnMinusOneDeselects Then
        cmbBox.ListIndex = -1
    End If
    
    ' Go through looking for a matching ItemData.
    For intIndex = 0 To cmbBox.ListCount - 1
        If cmbBox.ItemData(intIndex) = intItemData Then
            cmbBox.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
    
    ' Failed to find one. Perform a deselection just in case.
    cmbBox.ListIndex = -1
End Sub

' Deletes the currently selected property.
Public Sub DeleteSelectedProperty()
    Dim strKey As String
    Dim component As component

    ' Get component and selected key.
    Set component = GetCurrentComponent
    strKey = grdProperties.TextMatrix(grdProperties.Row, 0)
    
    ' Actually delete the property.
    component.DeleteProperty strKey
    
    ' Save component changes and reload the view.
    SaveCurrentComponent
    ShowComponent CLng(txtItemNumber.Text)
End Sub

' Opens the component distributor website with a search in place.
Private Sub LoadCurrentComponentWebsite()
    Dim component As component
    Set component = GetCurrentComponent
    
    OpenURL "https://pt.farnell.com/search?st=" & component.SearchCode
End Sub

' Category selection updated.
Private Sub cmbCategory_Click()
    If cmbCategory.ListIndex <> -1 Then
        LoadSubCategories cmbCategory.ItemData(cmbCategory.ListIndex), _
            cmbSubCategory
    Else
        cmbSubCategory.Clear
    End If
End Sub

' Database browse button clicked.
Private Sub cmdBrowseDatabase_Click()
    OpenDatabaseFile
End Sub

' Browse for the order file to load.
Private Sub cmdBrowseOrder_Click()
    ' Setup open dialog.
    dlgCommon.DialogTitle = "Import Distributor Order File"
    dlgCommon.DefaultExt = "csv"
    dlgCommon.Filter = "Comma Separated Files (*.csv)|*.csv|All Files (*.*)|*.*"
    dlgCommon.FileName = ""
    dlgCommon.ShowOpen
    
    ' Set the path.
    txtOrderLocation.Text = dlgCommon.FileName
End Sub

' Import current component into the database.
Private Sub cmdExport_Click()
    ' Check for duplicates.
    If CheckDuplicates Then
        Exit Sub
    End If
    
    ImportCurrentComponent
End Sub

' Go to the first component in the records.
Private Sub cmdFirst_Click()
    SaveCurrentComponent
    ShowComponent 0
End Sub

' Import the order file.
Private Sub cmdImport_Click()
    ' Check if there's an order file selected.
    If txtOrderLocation.Text = "" Then
        MsgBox "No order file selected. Please select one before importing.", _
            vbOKOnly + vbInformation, "No Order File Selected"
        Exit Sub
    End If
    
    ' Actually import the data.
    ImportOrder
End Sub

' Go to the last component in the records.
Private Sub cmdLast_Click()
    SaveCurrentComponent
    ShowComponent LastComponentIndex
End Sub

' Opens the component distributor website with a search in place.
Private Sub cmdLoadWebsite_Click()
    LoadCurrentComponentWebsite
End Sub

' Go to the next component in the records.
Private Sub cmdNext_Click()
    Dim lngCurrentIndex As Long
    
    lngCurrentIndex = CLng(txtItemNumber.Text)
    If lngCurrentIndex < LastComponentIndex Then
        SaveCurrentComponent
        ShowComponent lngCurrentIndex + 1
    End If
End Sub

' Go to the previous component in the records.
Private Sub cmdPrevious_Click()
    Dim lngCurrentIndex As Long
    
    lngCurrentIndex = CLng(txtItemNumber.Text)
    If lngCurrentIndex > 0 Then
        SaveCurrentComponent
        ShowComponent lngCurrentIndex - 1
    End If
End Sub

' Form just loaded.
Private Sub Form_Load()
    ' Set application icon.
    SetIcon Me.hWnd, "AAA_APPICON", True
    
    ' Setup the Flex Grid.
    grdProperties.TextMatrix(0, 0) = "Property"
    grdProperties.TextMatrix(0, 1) = "Value"
    grdProperties.ColWidth(0) = (grdProperties.Width / 2) - 5
    grdProperties.ColWidth(1) = (grdProperties.Width / 2) - 5
    
    ' Setup the image buttons.
    picFindExisting.Picture = imlButtons.ListImages("Find").ExtractIcon
    picRefreshCategories.Picture = imlButtons.ListImages("Refresh").ExtractIcon
    picRefreshSubCategories.Picture = imlButtons.ListImages("Refresh").ExtractIcon
    picRefreshPackages.Picture = imlButtons.ListImages("Refresh").ExtractIcon
    
    ' Disable the component panel and menu.
    fraComponent.Enabled = False
    mnuComponent.Enabled = False
End Sub

' User wants to edit a property.
Private Sub grdProperties_DblClick()
    Dim strKey As String
    Dim strValue As String
    Dim component As component
    
    ' Get properties and get user input.
    Set component = GetCurrentComponent
    strKey = grdProperties.TextMatrix(grdProperties.Row, 0)
    strValue = grdProperties.TextMatrix(grdProperties.Row, 1)
    strValue = InputBox(strKey & ":", "Edit Property", strValue)
    
    ' Change property if the user entered something.
    If strValue <> "" Then
        component.EditProperty strKey, strValue
        SaveCurrentComponent
        ShowComponent CLng(txtItemNumber.Text)
    End If
End Sub

' Check for keypresses on the properties grid.
Private Sub grdProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Delete a property for the user.
    If KeyCode = vbKeyDelete Then
        DeleteSelectedProperty
    End If
End Sub

' Component > Add Property clicked.
Private Sub mniComponentAddProperty_Click()
    If dlgProperty.ShowAdd(Me) Then
        SaveCurrentComponent
        ShowComponent CLng(txtItemNumber.Text)
    End If
End Sub

' Component > Delete Property clicked.
Private Sub mniComponentDeleteProperty_Click()
    DeleteSelectedProperty
End Sub

' Component > Import menu clicked.
Private Sub mniComponentExport_Click()
    ' If no database is associated, browse for one first.
    If Not IsDatabaseAssociated Then
        OpenDatabaseFile
    End If
    
    ' Import the current component if a database is associated.
    If IsDatabaseAssociated Then
        ImportCurrentComponent
    End If
End Sub

' Component > Load Website clicked.
Private Sub mniComponentLoadWebsite_Click()
    LoadCurrentComponentWebsite
End Sub

' Component > Next menu clicked.
Private Sub mniComponentNext_Click()
    cmdNext_Click
End Sub

' Component > Previous menu clicked.
Private Sub mniComponentPrevious_Click()
    cmdPrevious_Click
End Sub

' File > Exit menu clicked.
Private Sub mniFileExit_Click()
    Unload Me
End Sub

' File > Load Order menu clicked.
Private Sub mniFileLoadOrder_Click()
    cmdBrowseOrder_Click
    If txtOrderLocation.Text <> "" Then
        cmdImport_Click
    End If
End Sub

' File > Open Database menu clicked.
Private Sub mniFileOpenDatabase_Click()
    OpenDatabaseFile
End Sub

' Help > About menu clicked.
Private Sub mniHelpAbout_Click()
    frmAbout.Parent = Me
    frmAbout.Show
End Sub

' Find existing component name clicked.
Private Sub picFindExisting_Click()
    CheckDuplicates True
End Sub

' Refresh categories button clicked.
Private Sub picRefreshCategories_Click()
    LoadCategories cmbCategory
End Sub

' Refresh packages button clicked.
Private Sub picRefreshPackages_Click()
    LoadPackages cmbPackage
End Sub

' Refresh sub-categories button clicked.
Private Sub picRefreshSubCategories_Click()
    If cmbCategory.ListIndex <> -1 Then
        LoadSubCategories cmbCategory.ItemData(cmbCategory.ListIndex), _
            cmbSubCategory
    End If
End Sub
