VERSION 5.00
Begin VB.Form frmPartChooser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Component Selector"
   ClientHeight    =   8190
   ClientLeft      =   1470
   ClientTop       =   4590
   ClientWidth     =   3855
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   3855
   Begin VB.ListBox lstComponents 
      Height          =   2400
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   3855
   End
   Begin VB.ListBox lstSubCategories 
      Height          =   2400
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   3855
   End
   Begin VB.ListBox lstCategories 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Components:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Sub-Categories:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Categories:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
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

' Opens up a new component view.
Private Sub ShowComponent()
    Dim lngComponentID As Long
    Dim frmForm As frmComponent
    Set frmForm = frmComponent 'New frmComponent
    
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
    ' Populate our categories list.
    LoadCategories lstCategories
End Sub

' Handles the categories list click event.
Private Sub lstCategories_Click()
    lstComponents.Clear
    LoadSubCategories lstCategories.ItemData(lstCategories.ListIndex), _
        lstSubCategories
End Sub

' Handles the sub-categories list click event.
Private Sub lstSubCategories_Click()
    LoadComponents lstCategories.ItemData(lstCategories.ListIndex), _
        lstSubCategories.ItemData(lstSubCategories.ListIndex), lstComponents
End Sub

' Handles the components list click event.
Private Sub lstComponents_Click()
    ShowComponent
End Sub
