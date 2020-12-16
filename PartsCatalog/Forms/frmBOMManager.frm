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
         Left            =   6480
         TabIndex        =   13
         Top             =   240
         Width           =   495
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
Dim m_lngBOMComponentID As Long

' Populates the form with a BOM.
Private Sub ShowProject(lngProjectID As Long)
    ' Check if we are clearing a project.
    If lngProjectID = -1 Then
        ' Clear everything in the Project frame.
        lblProjectID.Caption = ""
        txtProjectName.Text = ""
        txtProjectRevision.Text = ""
        txtProjectDescription.Text = ""
        
        ' Clear everything in the Components frame.
        lstComponents.Clear
        Exit Sub
    End If
    
    ' Populate the form.
    LoadProjectDetail lngProjectID, Me
End Sub

' Updates the enabled/disabled controls.
Private Sub UpdateEnabledControls()
    cmdProjectSave.Enabled = Not IsNewProject
    cmdProjectRemove.Enabled = Not IsNewProject
    fraComponents.Enabled = Not IsNewProject
End Sub

' Populates the project frame.
Public Sub PopulateFromRecordset(rs As ADODB.Recordset)
    ' Populate Project frame.
    lblProjectID.Caption = rs.Fields("ID")
    txtProjectName.Text = rs.Fields("Name")
    txtProjectRevision.Text = rs.Fields("Revision")
    txtProjectDescription.Text = rs.Fields("Description")

    ' Update controls.
    UpdateEnabledControls
End Sub

' Checks if we are editing a new project.
Private Function IsNewProject() As Boolean
    IsNewProject = (ProjectID = -1)
End Function

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

' Form just loaded.
Private Sub Form_Load()
    ' Reset the project ID.
    ProjectID = -1
    
    ' Populate the projects and update controls.
    LoadProjects lstProjects
    UpdateEnabledControls
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

' Project ID getter.
Public Property Get ProjectID() As Long
    ProjectID = m_lngProjectID
End Property

' Project ID setter.
Public Property Let ProjectID(lngProjectID As Long)
    m_lngProjectID = lngProjectID
    
    ' Reset the Component ID as well if needed.
    If lngProjectID = -1 Then
        BOMComponentID = -1
    End If
    
    ' Show the BOM.
    ShowProject lngProjectID
End Property

' BOM component ID getter.
Public Property Get BOMComponentID() As Long
    BOMComponentID = m_lngBOMComponentID
End Property

' BOM component ID setter.
Public Property Let BOMComponentID(lngBOMComponentID As Long)
    m_lngBOMComponentID = lngBOMComponentID
End Property
