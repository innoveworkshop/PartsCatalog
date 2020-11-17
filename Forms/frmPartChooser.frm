VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPartChooser 
   BorderStyle     =   0  'None
   Caption         =   "Component Selector"
   ClientHeight    =   8370
   ClientLeft      =   1425
   ClientTop       =   4215
   ClientWidth     =   4110
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picToggleVisibility 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.ImageList imlFunctions 
      Left            =   2640
      Top             =   2640
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
            Picture         =   "frmPartChooser.frx":0000
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPartChooser.frx":6862
            Key             =   "Right"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstComponents 
      Height          =   2400
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   3855
   End
   Begin VB.ListBox lstSubCategories 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   3855
   End
   Begin VB.ListBox lstCategories 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblComponents 
      Caption         =   "Components:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label lblSubCategories 
      Caption         =   "Sub-Categories:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label lblCategories 
      Caption         =   "Categories:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmPartChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmPartChooser
''' The part chooser form.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private constants.
Private Const CTRL_MARGIN As Integer = 120

' Private variables.
Private m_frmParent As MDIForm
Private m_frmLastOpened As frmComponent
Private m_sngPanelWidth As Single
Private m_blnVisibility As Boolean

' Refreshes the contents of the form.
Public Sub RefreshLists()
    Dim lngCategoryID As Long
    Dim lngSubCategoryID As Long
    Dim lngIndex As Long
    
    ' Reset IDs to invalid state.
    lngCategoryID = -1
    lngSubCategoryID = -1
    
    ' Check if there are any categories selected.
    If lstCategories.ListIndex >= 0 Then
        lngCategoryID = lstCategories.ItemData(lstCategories.ListIndex)
    End If
    
    ' Check if there are any sub-categories selected.
    If lstSubCategories.ListIndex >= 0 Then
        lngSubCategoryID = lstSubCategories.ItemData(lstSubCategories.ListIndex)
    End If

    ' Clear the lists
    ClearContents False
    
    ' Populate our categories list.
    If IsDatabaseAssociated Then
        LoadCategories lstCategories
    Else
        Exit Sub
    End If
    
    ' Select previously selected category.
    If lngCategoryID >= 0 Then
        For lngIndex = 0 To lstCategories.ListCount - 1
            If lstCategories.ItemData(lngIndex) = lngCategoryID Then
                lstCategories.ListIndex = lngIndex
                Exit For
            End If
        Next lngIndex
    End If
    
    ' Select previously selected sub-category.
    If lngSubCategoryID >= 0 Then
        For lngIndex = 0 To lstSubCategories.ListCount - 1
            If lstSubCategories.ItemData(lngIndex) = lngSubCategoryID Then
                lstSubCategories.ListIndex = lngIndex
                Exit For
            End If
        Next lngIndex
    End If
End Sub

' Clears the fields in the form.
Public Sub ClearContents(Optional blnClearLastOpenedForm As Boolean = True)
    ' Clear the last opened form.
    If blnClearLastOpenedForm Then
        Set m_frmLastOpened = Nothing
    End If
    
    ' Clear lists.
    lstCategories.Clear
    lstSubCategories.Clear
    lstComponents.Clear
End Sub

' Sets the parent form.
Public Sub SetParent(frmParent As MDIForm)
    Set m_frmParent = frmParent
End Sub

' Toggles the panel visibility (normal or minimized).
Public Sub ToggleVisibility()
    Visibility = Not Visibility
End Sub

' Resizes the form to fit its parent.
Public Sub ResizeToFitParent()
    DockInParent
End Sub

' Docks this form in its parent.
Private Sub DockInParent()
    Dim sngListHeight As Single
    Dim sngListWidth As Single
    
    ' Position on the top-left corner.
    Left = 0
    Top = 0
    
    ' Set the height and calculate the height of each control group.
    Height = m_frmParent.ScaleHeight
    sngListHeight = (Height - (CTRL_MARGIN * 4) - (lblCategories.Height * 3)) / 3
    
    ' Set the width and calculate the width of each control group.
    Width = PanelWidth
    sngListWidth = Width - (CTRL_MARGIN * 2)
    
    ' Position the visibility toggler and its caption image.
    picToggleVisibility.Top = 0
    picToggleVisibility.Left = Width - picToggleVisibility.Width
    picToggleVisibility.Picture = imlFunctions.ListImages("Left").ExtractIcon
    
    ' Position and resize the categories group.
    lblCategories.Top = CTRL_MARGIN / 2
    lstCategories.Top = lblCategories.Top + lblCategories.Height
    lstCategories.Height = sngListHeight
    lstCategories.Width = sngListWidth
    
    ' Position and resize the sub-categories group.
    lblSubCategories.Top = lstCategories.Top + lstCategories.Height + CTRL_MARGIN
    lstSubCategories.Top = lblSubCategories.Top + lblSubCategories.Height
    lstSubCategories.Height = sngListHeight
    lstSubCategories.Width = sngListWidth
    
    ' Position and resize the components group.
    lblComponents.Top = lstSubCategories.Top + lstSubCategories.Height + CTRL_MARGIN
    lstComponents.Top = lblComponents.Top + lblComponents.Height
    lstComponents.Height = Me.Height - lstComponents.Top
    lstComponents.Width = sngListWidth
End Sub

' Hides the panel in a minimized state.
Private Sub HidePanel()
    ' Position the visibility toggler.
    picToggleVisibility.Top = 0
    picToggleVisibility.Left = 0
    
    ' Resize the panel.
    Height = picToggleVisibility.Height
    Width = picToggleVisibility.Width
    
    ' Change the toggler image.
    picToggleVisibility.Picture = imlFunctions.ListImages("Right").ExtractIcon
End Sub

' Opens up a new component view.
Private Sub ShowComponent()
    Dim lngComponentID As Long
    Dim frmForm As frmComponent

    ' Check if we have a last opened form.
    If m_frmLastOpened Is Nothing Then
        Set frmForm = New frmComponent
        Set m_frmLastOpened = frmForm
    End If
    
    ' Check if the last opened form wants to remain opened.
    If m_frmLastOpened.StayOpen Then
        Set frmForm = New frmComponent
        Set m_frmLastOpened = frmForm
    Else
        Set frmForm = m_frmLastOpened
    End If
    
    ' Get component ID and show the component dialog.
    lngComponentID = lstComponents.ItemData(lstComponents.ListIndex)
    If LoadComponentDetail(lngComponentID, frmForm) Then
        frmForm.ShowAligned
    Else
        Unload frmForm
    End If

    Set frmForm = Nothing
End Sub

' Event fired when the form loads up.
Private Sub Form_Load()
    ' Clear the contents.
    ClearContents
    
    ' Set the preffered panel width and show the form in dock.
    PanelWidth = Width
    Visibility = True
    
    ' Populate our categories list.
    If IsDatabaseAssociated Then
        LoadCategories lstCategories
    End If
End Sub

' Handles the categories list click event.
Private Sub lstCategories_Click()
    ' Check if there's anything selected.
    If lstCategories.ListIndex < 0 Then
        Exit Sub
    End If

    lstComponents.Clear
    LoadSubCategories lstCategories.ItemData(lstCategories.ListIndex), _
        lstSubCategories
End Sub

' Handles the sub-categories list click event.
Private Sub lstSubCategories_Click()
    ' Check if there's anything selected.
    If lstSubCategories.ListIndex < 0 Then
        Exit Sub
    End If

    LoadComponents lstCategories.ItemData(lstCategories.ListIndex), _
        lstSubCategories.ItemData(lstSubCategories.ListIndex), lstComponents
End Sub

' Handles the components list click event.
Private Sub lstComponents_Click()
    ' Check if there's anything selected.
    If lstComponents.ListIndex < 0 Then
        Exit Sub
    End If

    ShowComponent
End Sub

' Visibility flag getter.
Public Property Get Visibility() As Boolean
    Visibility = m_blnVisibility
End Property

' Visibility flag setter.
Public Property Let Visibility(blnVisibility As Boolean)
    m_blnVisibility = blnVisibility
    
    If blnVisibility Then
        ' Show panel.
        DockInParent
    Else
        ' Hide panel.
        HidePanel
    End If
End Property

' Panel width getter.
Public Property Get PanelWidth() As Single
    PanelWidth = m_sngPanelWidth
End Property

' Panel width setter.
Public Property Let PanelWidth(sngPanelWidth As Single)
    m_sngPanelWidth = sngPanelWidth
    Width = sngPanelWidth
End Property

' Toggle visibility event fired.
Private Sub picToggleVisibility_Click()
    ToggleVisibility
End Sub
