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

' Updates the sub-categories list according to the categories selection.
Private Sub UpdateSubCategories()
    lstComponents.Clear
    LoadSubCategories lstCategories.ItemData(lstCategories.ListIndex), _
        lstSubCategories
End Sub

' Updates the component list according to the sub-categories selection.
Private Sub UpdateComponents()
    LoadComponents lstCategories.ItemData(lstCategories.ListIndex), _
        lstSubCategories.ItemData(lstSubCategories.ListIndex), lstComponents
End Sub

' Event fired when the form loads up.
Private Sub Form_Load()
    ' Populate our categories list.
    LoadCategories lstCategories
End Sub

' Handles the categories list click event.
Private Sub lstCategories_Click()
    UpdateSubCategories
End Sub

' Handles the categories list key press event.
Private Sub lstCategories_KeyPress(KeyAscii As Integer)
    UpdateSubCategories
End Sub

' Handles the sub-categories list click event.
Private Sub lstSubCategories_Click()
    UpdateComponents
End Sub

' Handles the sub-categories list key press event.
Private Sub lstSubCategories_KeyPress(KeyAscii As Integer)
    UpdateComponents
End Sub
