VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComponent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Component"
   ClientHeight    =   5865
   ClientLeft      =   6135
   ClientTop       =   3315
   ClientWidth     =   8505
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
         NumButtons      =   7
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
            Key             =   "Spacer"
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComponent.frx":1A188
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

' Refreshes the contents of the component form.
Public Sub ReloadContent()
    LoadComponentDetail m_lngComponentID, Me
    SetStatusMessage "Component reloaded"
End Sub

' Saves the associated component.
Public Sub Save()
    ' Save component and refresh the lists.
    SaveComponent m_lngComponentID, txtName.Text, txtQuantity.Text, _
        txtNotes.Text, cmbCategory.ItemData(cmbCategory.ListIndex), _
        cmbSubCategory.ItemData(cmbSubCategory.ListIndex), _
        cmbPackage.ItemData(cmbPackage.ListIndex), _
        ComponentTabbedGridProperties(grdProperties)
    frmPartChooser.RefreshLists
    
    ' Check if we are renaming and make sure to propagate this to the
    ' associated assets.
    If IsRename Then
        RenameComponentDatasheet m_strOriginalName, txtName.Text
        RenameComponentImage m_strOriginalName, txtName.Text
    End If
    m_strOriginalName = txtName.Text
    
    ' Update status bar.
    SetStatusMessage "Component saved"
End Sub

' Populate Form from Recordset.
Public Sub PopulateFromRecordset(rs As ADODB.Recordset)
    Dim intIndex As Integer
    
    ' Store the component ID.
    m_lngComponentID = rs.Fields("ID")
    
    ' Set text fields.
    m_strOriginalName = rs.Fields("Name")
    txtName.Text = rs.Fields("Name")
    txtQuantity.Text = rs.Fields("Quantity")
    txtNotes.Text = rs.Fields("Notes")
    SetStatusMessage "Loaded text fields"
    
    ' Check for datasheet.
    cmdDatasheet.Enabled = ComponentHasDatasheet(rs.Fields("Name"))
    
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
    
    ' Set the component image.
    Dim strImage As String
    strImage = GetComponentImagePath(txtName.Text, cmbPackage.Text)
    If strImage <> vbNullString Then
        Dim picBitmap As Picture
        On Error GoTo PictureError
        Set picBitmap = LoadPicture(strImage)
        
        picImage.AutoRedraw = True
        picImage.PaintPicture picBitmap, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight
        Set picImage.Picture = picImage.Image
        SetStatusMessage "Component image loaded"
    End If
    
    Exit Sub
PictureError:
    Set picImage.Picture = Nothing
    MsgBox "An error occured while trying to load the image for this component.", _
        vbOKOnly + vbCritical, "Image Loading Error"
    SetStatusMessage "ERROR: Couldn't load the image for this component."
End Sub

' Populates the properties grid.
Public Sub PopulatePropertiesGrid(strProperties As String)
    Dim astrProperties() As String
    Dim astrKeyValue() As String
    
    ' Split the properties and preparate the grid for the properties.
    astrProperties = Split(strProperties, vbTab)
    grdProperties.Rows = UBound(astrProperties) + 2
    
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

' Sets a status message in the statusbar.
Private Sub SetStatusMessage(strMessage As String)
    stbStatusBar.SimpleText = strMessage
End Sub

' Make sure the form is kept opened.
Private Sub MaintainFormOpened(Optional blnState As Boolean = True)
    m_blnKeepOpen = blnState
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
    If m_strOriginalName = vbNullString Then
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

' Category selection updated.
Private Sub cmbCategory_Click()
    If cmbSubCategory.ListCount > 0 Then
        LoadSubCategories cmbCategory.ItemData(cmbCategory.ListIndex), _
            cmbSubCategory, True
    End If
End Sub

' Open component datasheet.
Private Sub cmdDatasheet_Click()
    OpenComponentDatasheet txtName.Text
End Sub

' Form just loaded up.
Private Sub Form_Load()
    ' Clear the opened status.
    m_blnKeepOpen = False

    ' Setup the Flex Grid.
    SetupPropertiesGrid
    
    ' Setup the ToolBar placeholder.
    tlbToolBar.Buttons("Spacer").Width = Me.ScaleWidth - _
        ((tlbToolBar.Buttons.Count - 1) * tlbToolBar.ButtonWidth) + _
        (tlbToolBar.ButtonWidth / 2)
End Sub

' Handles the toolbar button clicks.
Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Refresh"
            ReloadContent
        Case "Duplicate"
            MsgBox "Duplicate"
            frmPartChooser.RefreshLists
        Case "Save"
            Save
        Case "Delete"
            MsgBox "Delete"
            frmPartChooser.RefreshLists
        Case "KeepOpen"
            MaintainFormOpened (Button.Value = tbrPressed)
    End Select
End Sub

' Name text change event.
Private Sub txtName_Change()
    Me.Caption = txtName.Text
End Sub

' Getter for maintaining the window opened.
Public Property Get StayOpen() As Boolean
    StayOpen = m_blnKeepOpen
End Property

' Validate the quantity input to only contain numbers.
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case vbKeyBack, vbKeyClear, vbKeyDelete
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub
