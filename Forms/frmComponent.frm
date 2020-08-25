VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Component"
   ClientHeight    =   5865
   ClientLeft      =   6135
   ClientTop       =   3675
   ClientWidth     =   8505
   Icon            =   "frmComponent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   1
            Object.Width           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Duplicate"
            Object.ToolTipText     =   "Duplicate Component With New Name"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenDatasheet"
            Object.ToolTipText     =   "Open Datasheet"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteDatasheet"
            Object.ToolTipText     =   "Delete Datasheet"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DownloadDatasheet"
            Object.ToolTipText     =   "Download Datasheet"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spacer"
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "KeepOpen"
            Object.ToolTipText     =   "Maintain Window Opened"
            ImageIndex      =   5
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar stbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   5580
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   720
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":13916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":1A178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":209DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":2723C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":2DA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":34300
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   2025
      TabIndex        =   14
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdDatasheet 
      Caption         =   "Datasheet"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grdProperties 
      Height          =   3015
      Left            =   2280
      TabIndex        =   12
      Top             =   2520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5318
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483644
      GridColorFixed  =   -2147483644
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.TextBox txtNotes 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Text            =   "Notes"
      Top             =   1920
      Width           =   6135
   End
   Begin VB.ComboBox cmbPackage 
      Height          =   315
      Left            =   6840
      TabIndex        =   7
      Text            =   "Package"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox cmbSubCategory 
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Text            =   "Sub-Category"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Text            =   "Category"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Text            =   "Component Name"
      Top             =   720
      Width           =   5295
   End
   Begin VB.TextBox txtQuantity 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Text            =   "000000"
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Package:"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Sub-Category:"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Qnt:"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Menu mnuComponent 
      Caption         =   "&Component"
      Begin VB.Menu mniComponentRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mniComponentSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentDuplicate 
         Caption         =   "&Duplicate..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mniComponentSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mniComponentDelete 
         Caption         =   "D&elete"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mniComponentSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mniComponentKeepOpen 
         Caption         =   "&Keep Open"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      Begin VB.Menu mniImageBrowse 
         Caption         =   "&Browse..."
      End
      Begin VB.Menu mniImageSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniImageDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuDatasheet 
      Caption         =   "&Datasheet"
      Begin VB.Menu mniDatasheetOpen 
         Caption         =   "Open..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mniDatasheetDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mniDatasheetSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniDatasheetDownload 
         Caption         =   "Do&wnload..."
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmComponent
''' Component detail view form.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private methods.
Private m_lngComponentID As Long
Private m_strOriginalName As String
Private m_blnKeepOpen As Boolean
Private m_blnDirty As Boolean

' Shows the form keeping some distance from the parts chooser panel.
Public Sub ShowAligned()
    ' Move the window if necessary.
    If Me.Left < (frmPartChooser.Left + frmPartChooser.Width) Then
        Me.Left = Me.Left + frmPartChooser.Left + frmPartChooser.Width
    End If
    
    ' Set dirtiness.
    Dirty = False
    
    ' Show it.
    UpdateEnabledControls
    Show
End Sub

' Sets up the form for a new component.
Public Sub ShowNewComponent()
    ' Setup variables.
    ComponentID = -1
    StayOpen = True
    
    ' Populate form.
    txtQuantity.Text = 0
    txtName.Text = ""
    txtNotes.Text = ""
    Set picImage.Picture = Nothing
    LoadCategories cmbCategory, False
    cmbSubCategory.Text = ""
    LoadPackages cmbPackage, False
    
    ' Show the form and set dirtiness.
    ShowAligned
    Dirty = True
End Sub

' Duplicates the component.
Public Sub ShowNewDuplicate()
    Dim frmNewComponent As frmComponent
    Set frmNewComponent = New frmComponent
    
    ' Populate form.
    LoadComponentDetail ComponentID, frmNewComponent
    frmNewComponent.ComponentName = "Duplicate of " & txtName.Text
    
    ' Setup for new component.
    frmNewComponent.StayOpen = True
    frmNewComponent.ComponentID = -1
    
    ' Show the new component form and delete its local reference.
    frmNewComponent.ShowAligned
    frmNewComponent.Dirty = True
    Set frmNewComponent = Nothing
End Sub

' Refreshes the contents of the component form.
Public Sub ReloadContent()
    ' Check with the user for unsaved changes.
    If AbortUnsavedChanges Then
        Exit Sub
    End If
    
    ' Reload from the database.
    LoadComponentDetail ComponentID, Me
    
    ' Set dirtiness and update the status bar.
    Dirty = False
    SetStatusMessage "Component reloaded"
End Sub

' Deletes this component.
Public Sub DeleteMe()
    Dim intResponse As Integer
    
    ' Ask the user for confirmation.
    intResponse = MsgBox("Are you sure you want to delete this component?", _
        vbYesNo + vbQuestion, "Delete Component")
    If intResponse <> vbYes Then
        Exit Sub
    End If
    
    ' Perform the deletion.
    DeleteComponent ComponentID, m_strOriginalName
    Unload Me
    frmPartChooser.RefreshLists
End Sub

' Deletes the component datasheet.
Private Sub DeleteDatasheet()
    Dim intResponse As Integer
    
    ' Check if the user actually wants to do this.
    intResponse = MsgBox("Are you sure you want to delete this datasheet?", _
        vbYesNo + vbQuestion, "Delete " & ComponentName & " Datasheet")
    If intResponse = vbYes Then
        DeleteComponentDatasheet m_strOriginalName
        SetStatusMessage "Datasheet deleted"
    End If
    
    ' Update controls.
    UpdateEnabledControls
End Sub

' Deletes the component image.
Private Sub DeleteImage()
    Dim intResponse As Integer
    
    ' Check if the user actually wants to do this.
    intResponse = MsgBox("Are you sure you want to delete this component's image?" & _
        vbCrLf & vbCrLf & "Remember that this will only delete the component image" & _
        ", not the related package image.", vbYesNo + vbQuestion, _
        "Delete " & ComponentName & " Image")
    If intResponse = vbYes Then
        DeleteComponentImage m_strOriginalName
        SetStatusMessage "Component image deleted"
    End If
    
    ' Update controls.
    UpdateEnabledControls
End Sub

' Opens the component datasheet.
Public Sub OpenDatasheet()
    OpenComponentDatasheet m_strOriginalName
End Sub

' Downloads the component datasheet.
Private Sub DownloadDatasheet()
    Dim strURL As String
    Dim blnSuccess As Boolean
    
    ' Ask the user for the URL.
    strURL = InputBox("Please enter the URL for " & ComponentName & "'s datasheet:", _
        "Download " & ComponentName & " Datasheet")
    If strURL <> vbNullString Then
        ' Download the datasheet.
        blnSuccess = DownloadComponentDatasheet(m_strOriginalName, strURL)
        
        ' Check if we were successful.
        If blnSuccess Then
            SetStatusMessage "Datasheet downloaded successfully"
        Else
            SetStatusMessage "ERROR: Failed to download datasheet"
            MsgBox "Failed to download " & ComponentName & "'s datasheet", _
                vbOKOnly + vbExclamation, "Datasheet Download"
        End If
    End If
    
    ' Update controls.
    UpdateEnabledControls
End Sub

' Saves the associated component.
Public Sub Save()
    Dim lngCategoryID As Long
    Dim lngSubCategoryID As Long
    Dim lngPackageID As Long
    
    ' Setup before creating a new component.
    If IsNewComponent Then
        m_strOriginalName = txtName.Text
    End If
    
    ' Get category ID.
    If cmbCategory.ListIndex <> -1 Then
        lngCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex)
    Else
        lngCategoryID = -1
    End If
    
    ' Get sub-category ID.
    If cmbSubCategory.ListIndex <> -1 Then
        lngSubCategoryID = cmbSubCategory.ItemData(cmbSubCategory.ListIndex)
    Else
        lngSubCategoryID = -1
    End If
    
    ' Get package ID.
    If cmbPackage.ListIndex <> -1 Then
        lngPackageID = cmbPackage.ItemData(cmbPackage.ListIndex)
    Else
        lngPackageID = -1
    End If
    
    ' Save component and refresh the lists.
    ComponentID = SaveComponent(ComponentID, txtName.Text, txtQuantity.Text, _
        txtNotes.Text, lngCategoryID, lngSubCategoryID, lngPackageID, _
        ComponentTabbedGridProperties(grdProperties))
    frmPartChooser.RefreshLists
    
    ' Set dirtiness.
    Dirty = False
    
    ' Check if we are renaming and make sure to propagate this to the
    ' associated assets.
    If IsRename Then
        RenameComponentDatasheet m_strOriginalName, txtName.Text
        RenameComponentImage m_strOriginalName, txtName.Text
    End If
    m_strOriginalName = txtName.Text
    
    ' Update status bar and the tool bar.
    SetStatusMessage "Component saved"
    UpdateEnabledControls
End Sub

' Populate Form from Recordset.
Public Sub PopulateFromRecordset(rs As ADODB.Recordset)
    Dim intIndex As Integer
    
    ' Store the component ID.
    ComponentID = rs.Fields("ID")
    
    ' Set text fields.
    m_strOriginalName = rs.Fields("Name")
    txtName.Text = rs.Fields("Name")
    txtQuantity.Text = rs.Fields("Quantity")
    txtNotes.Text = rs.Fields("Notes")
    SetStatusMessage "Loaded text fields"
    
    ' Set the categories.
    cmbSubCategory.Clear
    LoadCategories cmbCategory, False
    If rs.Fields("CategoryID") >= 0 Then
        For intIndex = 0 To cmbCategory.ListCount
            If cmbCategory.ItemData(intIndex) = rs.Fields("CategoryID") Then
                cmbCategory.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    SetStatusMessage "Loaded categories"
    
    ' Load the sub-categories.
    LoadSubCategories rs.Fields("CategoryID"), cmbSubCategory, False
    If rs.Fields("SubCategoryID") >= 0 Then
        For intIndex = 0 To cmbSubCategory.ListCount
            If cmbSubCategory.ItemData(intIndex) = rs.Fields("SubCategoryID") Then
                cmbSubCategory.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    SetStatusMessage "Loaded sub-categories"
    
    ' Set the packages.
    LoadPackages cmbPackage, False
    If rs.Fields("PackageID") >= 0 Then
        For intIndex = 0 To cmbPackage.ListCount
            If cmbPackage.ItemData(intIndex) = rs.Fields("PackageID") Then
                cmbPackage.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    SetStatusMessage "Loaded component packages"
    
    ' Populate the properties grid.
    If Not IsNull(rs.Fields("Properties")) Then
        PopulatePropertiesGrid rs.Fields("Properties")
        SetStatusMessage "Loaded properties"
    Else
        grdProperties.Rows = 2
        grdProperties.TextMatrix(1, 0) = ""
        grdProperties.TextMatrix(1, 1) = ""
        SetStatusMessage "No component properties to load"
    End If
    
    ' Update controls.
    UpdateEnabledControls
End Sub

' Populates the properties grid.
Public Sub PopulatePropertiesGrid(strProperties As String)
    Dim astrProperties() As String
    Dim astrKeyValue() As String
    
    ' Split the properties and preparate the grid for the properties.
    astrProperties = Split(strProperties, vbTab)
    If UBound(astrProperties) = 0 Then
        grdProperties.Rows = UBound(astrProperties) + 2
    Else
        grdProperties.Rows = UBound(astrProperties) + 3
    End If
    
    ' Populate the properties.
    Dim intIndex As Integer
    For intIndex = 0 To UBound(astrProperties)
        ' Check if the property is populated.
        If astrProperties(intIndex) <> "" Then
            astrKeyValue = Split(astrProperties(intIndex), ": ")
            grdProperties.TextMatrix(intIndex + 1, 0) = astrKeyValue(0)
            grdProperties.TextMatrix(intIndex + 1, 1) = astrKeyValue(1)
        Else
            grdProperties.TextMatrix(1, 0) = ""
            grdProperties.TextMatrix(1, 1) = ""
        End If
    Next intIndex
End Sub

' Encodes the properties grid into a string to be stored in the database.
Public Function EncodePropertiesGrid() As String
    EncodePropertiesGrid = ""
End Function

' Aborts the current operation if the user selects Cancel to unsaved changes.
Public Function AbortUnsavedChanges() As Boolean
    Dim intResponse As Integer
    Dim strTitle As String
    
    ' Check for dirtiness.
    If Not Dirty Then
        AbortUnsavedChanges = False
        Exit Function
    End If
    
    ' Define the dialog title.
    If ComponentName <> vbNullString Then
        strTitle = "Save Changes to " & ComponentName & "?"
    Else
        strTitle = "Save New Component?"
    End If
    
    ' Ask the user.
    intResponse = MsgBox("You have changed this component." & vbCrLf & vbCrLf & _
        "Do you want to save the changes?", vbYesNoCancel + vbExclamation, strTitle)
    
    ' Decide what to do.
    If intResponse = vbCancel Then
        AbortUnsavedChanges = True
    Else
        ' Save before continuing.
        If intResponse = vbYes Then
            Save
        End If
        
        Dirty = False
        AbortUnsavedChanges = False
    End If
End Function

' Sets a status message in the statusbar.
Private Sub SetStatusMessage(strMessage As String)
    stbStatusBar.SimpleText = strMessage
End Sub

' Shows the component image.
Private Sub ShowImage()
    Dim strImage As String
    
    ' Get component image.
    strImage = GetComponentImagePath(m_strOriginalName, cmbPackage.Text)
    
    ' Set the component image.
    If strImage <> vbNullString Then
        Dim picBitmap As Picture
        On Error GoTo PictureError
        Set picBitmap = LoadPicture(strImage)
        
        picImage.AutoRedraw = True
        picImage.PaintPicture picBitmap, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight
        Set picImage.Picture = picImage.Image
    End If
    
    ' Handle image setting errors.
    Exit Sub
PictureError:
    Set picImage.Picture = Nothing
    MsgBox "An error occured while trying to load the image for this component.", _
        vbOKOnly + vbCritical, "Image Loading Error"
    SetStatusMessage "ERROR: Couldn't load the image for this component."
End Sub

' Updates which controls should be enabled/disabled.
Private Sub UpdateEnabledControls()
    Dim blnHasDatasheet As Boolean
    Dim blnHasImage As Boolean
    
    ' Check for datasheet and image.
    blnHasDatasheet = ComponentHasDatasheet(m_strOriginalName)
    blnHasImage = ComponentHasImage(m_strOriginalName)
    
    If IsNewComponent Then
        ' New component. Disable all the relevant buttons.
        tlbToolBar.Buttons("Refresh").Enabled = False
        tlbToolBar.Buttons("Duplicate").Enabled = False
        tlbToolBar.Buttons("Delete").Enabled = False
        tlbToolBar.Buttons("OpenDatasheet").Enabled = False
        tlbToolBar.Buttons("DeleteDatasheet").Enabled = False
        cmdDatasheet.Enabled = False
    Else
        ' Existing component. Enable all the options.
        tlbToolBar.Buttons("Refresh").Enabled = True
        tlbToolBar.Buttons("Duplicate").Enabled = True
        tlbToolBar.Buttons("Delete").Enabled = True
        
        ' Handle the absence of a datasheet.
        cmdDatasheet.Enabled = blnHasDatasheet
        tlbToolBar.Buttons("OpenDatasheet").Enabled = blnHasDatasheet
        tlbToolBar.Buttons("DeleteDatasheet").Enabled = blnHasDatasheet
        
        ' Handle the absence of an image.
        mniImageDelete.Enabled = blnHasImage
    End If
    
    ' Show component image.
    ShowImage
End Sub

' Setup the properties MSFlexGrid.
Private Sub SetupPropertiesGrid()
    ' Setup the columns size
    grdProperties.ColWidth(0) = (grdProperties.Width / 2) - 1
    grdProperties.ColWidth(1) = (grdProperties.Width / 2) - 1
    
    ' Setup header row.
    grdProperties.TextMatrix(0, 0) = "Property"
    grdProperties.ColAlignment(0) = flexAlignLeftCenter
    grdProperties.ColAlignment(1) = flexAlignLeftCenter
    grdProperties.FixedAlignment(0) = flexAlignCenterCenter
    grdProperties.FixedAlignment(1) = flexAlignCenterCenter
    grdProperties.TextMatrix(0, 1) = "Value"
End Sub

' Is the original name and the name in the TextBox different?
Private Function IsRename() As Boolean
    ' Check if we are creating a new component.
    If IsNewComponent Then
        IsRename = False
        Exit Function
    End If
    
    ' Check for actual component renaming.
    If m_strOriginalName <> txtName.Text Then
        IsRename = True
    Else
        IsRename = False
    End If
End Function

' Check if this is a new component.
Private Function IsNewComponent() As Boolean
    If ComponentID = -1 Then
        IsNewComponent = True
    Else
        IsNewComponent = False
    End If
End Function

' Category selection updated.
Private Sub cmbCategory_Click()
    ' Prevent the database connection from being closed when populating
    ' from a RecordSet.
    LoadSubCategories cmbCategory.ItemData(cmbCategory.ListIndex), _
        cmbSubCategory, IsNewComponent
    
    ' Set dirtiness.
    Dirty = True
End Sub

' Package selection updated.
Private Sub cmbPackage_Change()
    ' Set dirtiness.
    Dirty = True
End Sub

' Sub-category selection updated.
Private Sub cmbSubCategory_Change()
    ' Set dirtiness.
    Dirty = True
End Sub

' Open component datasheet.
Private Sub cmdDatasheet_Click()
    OpenDatasheet
End Sub

' Datasheet button was clicked in some way by the mouse.
Private Sub cmdDatasheet_MouseDown(Button As Integer, Shift As Integer, X As Single, _
        Y As Single)
    ' Detect right click.
    If Button = vbRightButton Then
        PopupMenu mnuDatasheet
    End If
End Sub

' Form just loaded up.
Private Sub Form_Load()
    ' Clear the opened status.
    StayOpen = False

    ' Setup some controls.
    UpdateEnabledControls
    SetupPropertiesGrid
    
    ' Set dirtiness.
    Dirty = False
End Sub

' Form is about to be unloaded.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If AbortUnsavedChanges Then
        Cancel = 1
    End If
End Sub

' Form just got unloaded.
Private Sub Form_Unload(Cancel As Integer)
    ' Make sure the parts chooser form doesn't use this form.
    ComponentID = -1
    StayOpen = True
End Sub

' Properties grid was double clicked.
Private Sub grdProperties_DblClick()
    Dim dlgProperty As dlgEditProperty
    
    ' Initialize the dialog.
    Set dlgProperty = New dlgEditProperty
    dlgProperty.CentralizeInForm frmMain, Me

    ' Determine if we are editing or adding a property.
    If grdProperties.Row = grdProperties.Rows - 1 Then
        ' Empty row clicked. Let's add a new entry.
        dlgProperty.ShowNew
    Else
        ' Edit an existing property.
        dlgProperty.ShowEditor grdProperties.TextMatrix(grdProperties.Row, 0), _
            grdProperties.TextMatrix(grdProperties.Row, 1)
    End If
    
    ' Should we save the property?
    If dlgProperty.Save Then
        ' Edit the current row.
        grdProperties.TextMatrix(grdProperties.Row, 0) = dlgProperty.Key
        grdProperties.TextMatrix(grdProperties.Row, 1) = dlgProperty.Value
        
        ' Add a new row in case we added a property.
        If grdProperties.Row = grdProperties.Rows - 1 Then
            grdProperties.Rows = grdProperties.Rows + 1
        End If
        
        ' Set status and dirtiness.
        SetStatusMessage dlgProperty.Key & " property edited"
        Dirty = True
    End If
    
    Set dlgProperty = Nothing
End Sub

' Component > Delete menu clicked.
Private Sub mniComponentDelete_Click()
    DeleteMe
End Sub

' Component > Duplicate menu clicked.
Private Sub mniComponentDuplicate_Click()
    ShowNewDuplicate
End Sub

' Component > Keep Opened menu clicked.
Private Sub mniComponentKeepOpen_Click()
    StayOpen = Not StayOpen
End Sub

' Component > Refresh menu clicked.
Private Sub mniComponentRefresh_Click()
    ReloadContent
End Sub

' Component > Save menu clicked.
Private Sub mniComponentSave_Click()
    Save
End Sub

' Datasheet > Delete menu clicked.
Private Sub mniDatasheetDelete_Click()
    DeleteDatasheet
End Sub

' Datasheet > Download menu clicked.
Private Sub mniDatasheetDownload_Click()
    DownloadDatasheet
End Sub

' Datasheet > Open menu clicked.
Private Sub mniDatasheetOpen_Click()
    OpenDatasheet
End Sub

' Image > Delete menu clicked.
Private Sub mniImageDelete_Click()
    DeleteImage
End Sub

' Image double clicked.
Private Sub picImage_DblClick()
    MsgBox "TODO: Browse for an image."
End Sub

' Image was clicked in some way by the mouse.
Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, _
        Y As Single)
    ' Detect right click.
    If Button = vbRightButton Then
        PopupMenu mnuImage
    End If
End Sub

' Handles the toolbar button clicks.
Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Refresh"
            ReloadContent
        Case "Duplicate"
            ShowNewDuplicate
        Case "Save"
            Save
        Case "Delete"
            DeleteMe
        Case "OpenDatasheet"
            OpenDatasheet
        Case "DeleteDatasheet"
            DeleteDatasheet
        Case "DownloadDatasheet"
            DownloadDatasheet
        Case "KeepOpen"
            StayOpen = (Button.Value = tbrPressed)
    End Select
End Sub

' Name text change event.
Private Sub txtName_Change()
    ' Change window title and component menu caption.
    Me.Caption = txtName.Text
    mnuComponent.Caption = txtName.Text
    
    ' Set dirtiness.
    Dirty = True
End Sub

' Notes change event.
Private Sub txtNotes_Change()
    ' Set dirtiness.
    Dirty = True
End Sub

' Validate the quantity input to only contain numbers.
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    ' Only accept numbers.
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case vbKeyBack, vbKeyClear, vbKeyDelete
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
        Case Else
            KeyAscii = 0
            Beep
    End Select
    
    ' Set dirtiness.
    Dirty = True
End Sub

' Getter for maintaining the window opened.
Public Property Get StayOpen() As Boolean
    StayOpen = m_blnKeepOpen
End Property

' Setter for maintaining the window opened.
Public Property Let StayOpen(blnKeepOpen As Boolean)
    m_blnKeepOpen = blnKeepOpen
    
    ' Set the toolbar button and menu item accordingly.
    mniComponentKeepOpen.Checked = blnKeepOpen
    If m_blnKeepOpen Then
        tlbToolBar.Buttons("KeepOpen").Value = tbrPressed
    Else
        tlbToolBar.Buttons("KeepOpen").Value = tbrUnpressed
    End If
End Property

' Getter for the dirty attribute.
Public Property Get Dirty() As Boolean
    Dirty = m_blnDirty
End Property

' Setter for the dirty attribute.
Public Property Let Dirty(blnDirty As Boolean)
    m_blnDirty = blnDirty
    
    ' Add or remove an asterisk to the window title.
    If blnDirty Then
        If Right(Me.Caption, 1) <> "*" Then
            Me.Caption = Me.Caption & "*"
        End If
    Else
        If Right(Me.Caption, 1) = "*" Then
            Me.Caption = Left$(Me.Caption, Len(Me.Caption) - 1)
        End If
    End If
End Property

' Getter for the component ID.
Public Property Get ComponentID() As Long
    ComponentID = m_lngComponentID
End Property

' Setter for the component ID.
Public Property Let ComponentID(lngID As Long)
    m_lngComponentID = lngID
    
    ' Make sure the original name is reset.
    If IsNewComponent Then
        m_strOriginalName = vbNullString
    End If
End Property

' Getter for the component name.
Public Property Get ComponentName() As String
    ComponentName = txtName.Text
End Property

' Setter for the component name.
Public Property Let ComponentName(strName As String)
    txtName.Text = strName
End Property
