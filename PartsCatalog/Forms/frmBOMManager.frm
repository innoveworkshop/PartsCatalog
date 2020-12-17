VERSION 5.00
Begin VB.Form frmBOMManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BOM Manager"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10935
   Icon            =   "frmBOMManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraComponents 
      Caption         =   "Components"
      Height          =   6375
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdComponentSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   5880
         Width           =   3255
      End
      Begin VB.CommandButton cmdComponentRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   5880
         Width           =   1695
      End
      Begin VB.CommandButton cmdComponentAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5880
         Width           =   1695
      End
      Begin VB.CommandButton cmdRefDesRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5400
         TabIndex        =   20
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefDesRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefDesAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtRefDes 
         Height          =   315
         Left            =   5400
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ListBox lstRefDes 
         Height          =   3960
         Left            =   3720
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cmbComponent 
         Height          =   315
         Left            =   3720
         TabIndex        =   12
         Top             =   480
         Width           =   3255
      End
      Begin VB.ListBox lstComponents 
         Height          =   5520
         ItemData        =   "frmBOMManager.frx":6852
         Left            =   120
         List            =   "frmBOMManager.frx":6859
         TabIndex        =   10
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblItemID 
         Alignment       =   1  'Right Justify
         Caption         =   "00000"
         Height          =   255
         Left            =   6240
         TabIndex        =   28
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Reference Designators:"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblDescription 
         Caption         =   "COMPONENT DESCRIPTION GOES IN HERE"
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   840
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblComponentID 
         Alignment       =   1  'Right Justify
         Caption         =   "00000"
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Component:"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraProject 
      Caption         =   "Project Information"
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   3495
      Begin VB.TextBox txtProjectDescription 
         Height          =   555
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtProjectRevision 
         Height          =   315
         Left            =   2880
         TabIndex        =   25
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtProjectName 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdProjectAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   3255
      End
      Begin VB.CommandButton cmdProjectSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdProjectRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Rev:"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblProjectID 
         Alignment       =   1  'Right Justify
         Caption         =   "000"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.ListBox lstProjects 
      Height          =   3180
      ItemData        =   "frmBOMManager.frx":686C
      Left            =   120
      List            =   "frmBOMManager.frx":6873
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Projects:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmBOMManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmBOMManager
''' A form to manage Bill of Materials.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

Dim m_lngProjectID As Long
Dim m_lngBOMItemID As Long

' Populates the project frame.
Public Sub PopulateProjectFromRecordset(rs As ADODB.Recordset)
    ' Populate Project frame.
    lblProjectID.Caption = rs.Fields("ID")
    txtProjectName.Text = rs.Fields("Name")
    txtProjectRevision.Text = rs.Fields("Revision")
    txtProjectDescription.Text = rs.Fields("Description")
    
    ' Populate Components frame.
    LoadProjectBOM lstComponents, False

    ' Update controls.
    UpdateEnabledControls
End Sub

' Populates the component pane.
Public Sub PopulateBOMItemFromRecordset(rs As ADODB.Recordset)
    Dim astrRefDes() As String
    Dim intIndex As Integer
    
    ' Populate component area.
    If IsNull(rs.Fields("ComponentID")) Then
        lblComponentID.Caption = ""
        ' TODO: Select the no selection option in the combobox.
    Else
        lblComponentID.Caption = rs.Fields("ComponentID")
        ' TODO: Select the component in the combobox.
    End If
    
    ' Populate reference designators.
    lblItemID.Caption = rs.Fields("ID")
    astrRefDes = Split(rs.Fields("RefDes"), ", ")
    lstRefDes.Clear
    For intIndex = 0 To UBound(astrRefDes)
        lstRefDes.AddItem astrRefDes(intIndex)
    Next intIndex

    ' Update controls.
    UpdateEnabledControls
End Sub

' Populates the form with a BOM.
Private Sub ShowProject(lngProjectID As Long)
    ' Check if we are clearing a project.
    If lngProjectID = -1 Then
        ' Clear everything in the Project frame.
        lblProjectID.Caption = ""
        txtProjectName.Text = ""
        txtProjectRevision.Text = ""
        txtProjectDescription.Text = ""
        
        ' Clear everything in the Components frame and update the controls.
        lstComponents.Clear
        UpdateEnabledControls
        Exit Sub
    End If
    
    ' Populate the form.
    LoadProjectDetail lngProjectID, Me
End Sub

' Populates the component panel with the selected item.
Private Sub ShowItem(lngItemID As Long)
    ' Check if we are clearing an item.
    If lngItemID = -1 Then
        ' Clear everything in the item pane.
        lblComponentID.Caption = ""
        lblItemID.Caption = ""
        txtRefDes.Text = ""
        lstRefDes.Clear
        
        ' Update controls.
        UpdateEnabledControls
        Exit Sub
    End If
    
    ' Populate the form.
    LoadProjectBOMItem lngItemID, Me
End Sub

' Updates the enabled/disabled controls.
Private Sub UpdateEnabledControls()
    ' Project-related controls.
    cmdProjectSave.Enabled = Not IsNewProject
    cmdProjectRemove.Enabled = Not IsNewProject
    
    ' Component-related controls.
    cmdComponentAdd.Enabled = Not IsNewProject
    cmdRefDesAdd.Enabled = Not IsNewBOMItem
    cmdRefDesRename.Enabled = (lstRefDes.ListIndex >= 0)
    cmdRefDesRemove.Enabled = (lstRefDes.ListIndex >= 0)
    cmdComponentRemove.Enabled = Not IsNewBOMItem
    cmdComponentSave.Enabled = Not IsNewBOMItem
    fraComponents.Enabled = Not IsNewProject
End Sub

' Checks if we are editing a new project.
Private Function IsNewProject() As Boolean
    IsNewProject = (ProjectID = -1)
End Function

' Checks if we are editing a new BOM item.
Private Function IsNewBOMItem() As Boolean
    IsNewBOMItem = (BOMItemID = -1)
End Function

' Add component button clicked.
Private Sub cmdComponentAdd_Click()
    Dim intIndex As Integer
    Dim lngItemID As Long
    Dim astrRefDes() As String
    
    ' Create an empty BOM item.
    lngItemID = SaveBOMItem(-1, ProjectID, astrRefDes, -1)
    
    ' Select the BOM item from the list.
    LoadProjectBOM lstComponents
    For intIndex = 0 To lstComponents.ListCount - 1
        If lstComponents.ItemData(intIndex) = lngItemID Then
            lstComponents.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Remove component button clicked.
Private Sub cmdComponentRemove_Click()
    ' Delete the item and reset the ID.
    DeleteBOMItem BOMItemID
    BOMItemID = -1
    
    ' Reload the BOM items and update controls.
    LoadProjectBOM lstComponents
    UpdateEnabledControls
End Sub

' Save component button clicked.
Private Sub cmdComponentSave_Click()
    Dim intIndex As Integer
    Dim lngItemID As Long
    Dim astrRefDes() As String
    
    ' Get the reference designator array.
    If lstRefDes.ListCount > 0 Then
        ReDim astrRefDes(lstRefDes.ListCount - 1)
    
        For intIndex = 0 To UBound(astrRefDes)
            astrRefDes(intIndex) = lstRefDes.List(intIndex)
        Next intIndex
    End If
    
    MsgBox "TODO: Get ID from combobox"
    
    ' Create an empty BOM item.
    lngItemID = SaveBOMItem(BOMItemID, ProjectID, astrRefDes, -1)
    
    ' Select the BOM item from the list.
    LoadProjectBOM lstComponents
    For intIndex = 0 To lstComponents.ListCount - 1
        If lstComponents.ItemData(intIndex) = lngItemID Then
            lstComponents.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Add project button clicked.
Private Sub cmdProjectAdd_Click()
    Dim intIndex As Integer
    
    ' Create the project, reset the ID.
    ProjectID = SaveProject(-1, txtProjectName.Text, txtProjectRevision.Text, _
        txtProjectDescription.Text)
    
    ' Populate project list and select the newly added item.
    LoadProjects lstProjects
    For intIndex = 0 To lstProjects.ListCount - 1
        If lstProjects.ItemData(intIndex) = ProjectID Then
            lstProjects.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Delete project button clicked.
Private Sub cmdProjectRemove_Click()
    ' Delete the project and clear the ID.
    DeleteProject ProjectID
    ProjectID = -1
    
    ' Reload the list and update controls.
    LoadProjects lstProjects
    UpdateEnabledControls
End Sub

' Rename project button clicked.
Private Sub cmdProjectSave_Click()
    Dim intIndex As Integer
    
    ' Update the project, reset the ID.
    ProjectID = SaveProject(ProjectID, txtProjectName.Text, txtProjectRevision.Text, _
        txtProjectDescription.Text)
    
    ' Populate project list and select the renamed item.
    LoadProjects lstProjects
    For intIndex = 0 To lstProjects.ListCount - 1
        If lstProjects.ItemData(intIndex) = ProjectID Then
            lstProjects.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Add reference designator button clicked.
Private Sub cmdRefDesAdd_Click()
    lstRefDes.AddItem txtRefDes.Text
    lstRefDes.ListIndex = lstRefDes.NewIndex
End Sub

' Remove reference designator button clicked.
Private Sub cmdRefDesRemove_Click()
    ' Check if there's anything selected.
    If lstRefDes.ListIndex < 0 Then
        Exit Sub
    End If
    
    ' Remove the selected item and update controls.
    lstRefDes.RemoveItem lstRefDes.ListIndex
    UpdateEnabledControls
End Sub

' Rename reference designator button clicked.
Private Sub cmdRefDesRename_Click()
    ' Check if there's anything selected.
    If lstRefDes.ListIndex < 0 Then
        Exit Sub
    End If
    
    ' Rename the selected item.
    lstRefDes.List(lstRefDes.ListIndex) = txtRefDes.Text
End Sub

' Form just loaded.
Private Sub Form_Load()
    ' Reset the project ID.
    ProjectID = -1
    
    ' Populate the projects and update controls.
    LoadProjects lstProjects
    UpdateEnabledControls
End Sub

' BOM item list item clicked.
Private Sub lstComponents_Click()
    ' Check if there's anything selected.
    If lstComponents.ListIndex < 0 Then
        Exit Sub
    End If
    
    BOMItemID = lstComponents.ItemData(lstComponents.ListIndex)
End Sub

' Project selection changed.
Private Sub lstProjects_Click()
    ' Check if there's anything selected.
    If lstProjects.ListIndex < 0 Then
        Exit Sub
    End If
    
    ' Update the project ID.
    ProjectID = lstProjects.ItemData(lstProjects.ListIndex)
End Sub

' Reference designator list clicked.
Private Sub lstRefDes_Click()
    ' Check if there's anything selected.
    If lstRefDes.ListIndex < 0 Then
        Exit Sub
    End If
    
    txtRefDes.Text = lstRefDes.List(lstRefDes.ListIndex)
    UpdateEnabledControls
End Sub

' Project ID getter.
Public Property Get ProjectID() As Long
    ProjectID = m_lngProjectID
End Property

' Project ID setter.
Public Property Let ProjectID(lngProjectID As Long)
    m_lngProjectID = lngProjectID
    
    ' Reset the Component ID as well if needed.
    If lngProjectID = -1 Then
        BOMItemID = -1
    End If
    
    ' Show the BOM.
    ShowProject lngProjectID
End Property

' BOM component ID getter.
Public Property Get BOMItemID() As Long
    BOMItemID = m_lngBOMItemID
End Property

' BOM component ID setter.
Public Property Let BOMItemID(lngItemID As Long)
    m_lngBOMItemID = lngItemID
    
    ' Show the component.
    ShowItem lngItemID
End Property
