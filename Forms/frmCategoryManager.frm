VERSION 5.00
Begin VB.Form frmCategoryManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Category Manager"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11040
   Icon            =   "frmCategoryManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5766.268
   ScaleMode       =   0  'User
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSubCategory 
      Caption         =   "Sub-Category"
      Height          =   2055
      Left            =   7320
      TabIndex        =   10
      Top             =   3360
      Width           =   3495
      Begin VB.TextBox txtSubCategoryName 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdSubCategoryAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmdSubCategoryRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdSubCategoryRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblSubCategoryID 
         Alignment       =   1  'Right Justify
         Caption         =   "00000"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraCategory 
      Caption         =   "Category"
      Height          =   2055
      Left            =   7320
      TabIndex        =   4
      Top             =   240
      Width           =   3495
      Begin VB.CommandButton cmdCategoryRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdCategoryRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdCategoryAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtCategoryName 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblCategoryID 
         Alignment       =   1  'Right Justify
         Caption         =   "000"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ListBox lstCategories 
      Height          =   5130
      ItemData        =   "frmCategoryManager.frx":6852
      Left            =   120
      List            =   "frmCategoryManager.frx":6859
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.ListBox lstSubCategories 
      Height          =   5130
      ItemData        =   "frmCategoryManager.frx":686C
      Left            =   3720
      List            =   "frmCategoryManager.frx":6873
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Sub-Categories:"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Categories:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmCategoryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmCategoryManager
''' A form to manage categories and sub-categories.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_lngCategoryID As Long
Private m_lngSubCategoryID As Long

' Check if we are dealing with a new category or editing an existing one.
Public Function IsNewCategory() As Boolean
    If CategoryID = -1 Then
        IsNewCategory = True
    Else
        IsNewCategory = False
    End If
End Function

' Check if we are dealing with a new sub-category or editing an existing one.
Public Function IsNewSubCategory() As Boolean
    If SubCategoryID = -1 Then
        IsNewSubCategory = True
    Else
        IsNewSubCategory = False
    End If
End Function

' Updates the enabled/disabled controls.
Private Sub UpdateEnabledControls()
    cmdCategoryRename.Enabled = Not IsNewCategory
    cmdCategoryRemove.Enabled = Not IsNewCategory
    txtSubCategoryName.Enabled = Not IsNewCategory
    cmdSubCategoryAdd.Enabled = Not IsNewCategory
    cmdSubCategoryRename.Enabled = Not IsNewSubCategory
    cmdSubCategoryRemove.Enabled = Not IsNewSubCategory
End Sub

' Category add button clicked.
Private Sub cmdCategoryAdd_Click()
    Dim intIndex As Integer
    
    ' Add the category.
    CategoryID = SaveCategory(-1, txtCategoryName.Text)
    
    ' Populate category list and select the newly added item.
    LoadCategories lstCategories
    For intIndex = 0 To lstCategories.ListCount - 1
        If lstCategories.ItemData(intIndex) = CategoryID Then
            lstCategories.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Category remove button clicked.
Private Sub cmdCategoryRemove_Click()
    ' Check if we can do this.
    If IsNewCategory Then
        Exit Sub
    End If
    
    ' Delete category.
    DeleteCategory CategoryID
    
    ' Reset the ID, name, and repopulate the categories list.
    CategoryID = -1
    txtCategoryName.Text = ""
    LoadCategories lstCategories
End Sub

' Category rename button clicked.
Private Sub cmdCategoryRename_Click()
    Dim intIndex As Integer

    ' Check if we can do this.
    If IsNewCategory Then
        Exit Sub
    End If
    
    ' Update the category, reset the ID.
    CategoryID = SaveCategory(CategoryID, txtCategoryName.Text)
    
    ' Populate category list and select the newly added item.
    LoadCategories lstCategories
    For intIndex = 0 To lstCategories.ListCount - 1
        If lstCategories.ItemData(intIndex) = CategoryID Then
            lstCategories.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Sub-category add button clicked.
Private Sub cmdSubCategoryAdd_Click()
    Dim intIndex As Integer
    
    ' Add the sub-category.
    SubCategoryID = SaveSubCategory(-1, CategoryID, txtSubCategoryName.Text)
    
    ' Populate sub-category list and select the newly added item.
    LoadSubCategories CategoryID, lstSubCategories
    For intIndex = 0 To lstSubCategories.ListCount - 1
        If lstSubCategories.ItemData(intIndex) = SubCategoryID Then
            lstSubCategories.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Sub-category remove button clicked.
Private Sub cmdSubCategoryRemove_Click()
    ' Check if we can do this.
    If IsNewSubCategory Then
        Exit Sub
    End If
    
    ' Delete sub-category.
    DeleteSubCategory SubCategoryID
    
    ' Reset the ID, name, and repopulate the sub-categories list.
    SubCategoryID = -1
    txtSubCategoryName.Text = ""
    LoadSubCategories CategoryID, lstSubCategories
End Sub

' Sub-category rename button clicked.
Private Sub cmdSubCategoryRename_Click()
    Dim intIndex As Integer

    ' Check if we can do this.
    If IsNewSubCategory Then
        Exit Sub
    End If
    
    ' Update the sub-category, reset the ID.
    SubCategoryID = SaveSubCategory(SubCategoryID, CategoryID, txtSubCategoryName.Text)
    
    ' Populate sub-category list and select the newly added item.
    LoadSubCategories CategoryID, lstSubCategories
    For intIndex = 0 To lstSubCategories.ListCount - 1
        If lstSubCategories.ItemData(intIndex) = SubCategoryID Then
            lstSubCategories.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Form just loaded.
Private Sub Form_Load()
    ' Reset the category ID.
    CategoryID = -1
    
    ' Populate the categories list and update controls.
    LoadCategories lstCategories
    UpdateEnabledControls
End Sub

' Category selection changed.
Private Sub lstCategories_Click()
    ' Update the text field and ID.
    txtCategoryName.Text = lstCategories.Text
    CategoryID = lstCategories.ItemData(lstCategories.ListIndex)
    
    ' Populate sub-categories list.
    LoadSubCategories CategoryID, lstSubCategories
End Sub

' Sub-category selection changed.
Private Sub lstSubCategories_Click()
    ' Update the text field and ID.
    txtSubCategoryName.Text = lstSubCategories.Text
    SubCategoryID = lstSubCategories.ItemData(lstSubCategories.ListIndex)
End Sub

' Gets the category ID.
Public Property Get CategoryID() As Long
    CategoryID = m_lngCategoryID
End Property

' Sets the category ID.
Public Property Let CategoryID(lngID As Long)
    m_lngCategoryID = lngID
    
    ' Sets the ID label.
    If lngID <> -1 Then
        lblCategoryID.Caption = lngID
    Else
        lblCategoryID.Caption = ""
        txtCategoryName.Text = ""
    End If
    
    ' Reset the sub-category ID.
    SubCategoryID = -1
End Property

' Gets the sub-category ID.
Public Property Get SubCategoryID() As Long
    SubCategoryID = m_lngSubCategoryID
End Property

' Sets the sub-category ID.
Public Property Let SubCategoryID(lngID As Long)
    m_lngSubCategoryID = lngID
    
    ' Sets the ID label.
    If lngID <> -1 Then
        lblSubCategoryID.Caption = lngID
    Else
        lblSubCategoryID.Caption = ""
        txtSubCategoryName.Text = ""
        lstSubCategories.Clear
    End If
    
    ' Update the controls.
    UpdateEnabledControls
End Property
