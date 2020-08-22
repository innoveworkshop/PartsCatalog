VERSION 5.00
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

' Sets the parent form.
Public Sub SetParent(frmParent As MDIForm)
    Set m_frmParent = frmParent
End Sub

' Resizes the form to fit its parent.
Public Sub ResizeToFitParent()
    DockInParent
End Sub

' Docks this form in its parent.
Private Sub DockInParent()
    Dim intListHeight As Integer
    
    ' Position on the top-left corner.
    Left = 0
    Top = 0
    
    ' Set the height and calculate the height of each control group.
    Me.Height = m_frmParent.ScaleHeight
    intListHeight = (Me.Height - (CTRL_MARGIN * 4) - (lblCategories.Height * 3)) / 3
    
    ' Position and resize the categories group.
    lblCategories.Top = CTRL_MARGIN / 2
    lstCategories.Top = lblCategories.Top + lblCategories.Height
    lstCategories.Height = intListHeight
    
    ' Position and resize the sub-categories group.
    lblSubCategories.Top = lstCategories.Top + lstCategories.Height + CTRL_MARGIN
    lstSubCategories.Top = lblSubCategories.Top + lblSubCategories.Height
    lstSubCategories.Height = intListHeight
    
    ' Position and resize the components group.
    lblComponents.Top = lstSubCategories.Top + lstSubCategories.Height + CTRL_MARGIN
    lstComponents.Top = lblComponents.Top + lblComponents.Height
    lstComponents.Height = Me.Height - lstComponents.Top
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
    
    ' Position the new form.
    If frmForm.Left < (Me.Left + Me.Width) Then
        frmForm.Left = frmForm.Left + Me.Left + Me.Width
    End If
    
    ' Get component ID.
    lngComponentID = lstComponents.ItemData(lstComponents.ListIndex)
    
    ' TODO: Have this form component as a private variable and check if we
    '       should open a new one based on it's movement previously.
    If LoadComponentDetail(lngComponentID, frmForm) Then
        frmForm.Show
    Else
        Unload frmForm
    End If

    Set frmForm = Nothing
End Sub

' Event fired when the form loads up.
Private Sub Form_Load()
    ' Clear the last opened form.
    Set m_frmLastOpened = Nothing

    ' Dock the form.
    DockInParent
    
    ' Populate our categories list.
    LoadCategories lstCategories
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
