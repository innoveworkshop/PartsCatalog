VERSION 5.00
Begin VB.Form frmDuplicateComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existing Component"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6855
   Icon            =   "frmDuplicateComponent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImportAnyway 
      Caption         =   "Import As New"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdateQuantity 
      Caption         =   "Update Quantity"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtQuantity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Text            =   "000000"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Text            =   "Component Name"
      Top             =   840
      Width           =   3855
   End
   Begin VB.ComboBox cmbCategory 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Text            =   "Category"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox cmbSubCategory 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Text            =   "Sub-Category"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox cmbPackage 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      TabIndex        =   2
      Text            =   "Package"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtNotes 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   "Notes"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   1785
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Do you want to import a new component with the same name as this one or just update the quantity of this existing component?"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   6615
   End
   Begin VB.Label Label7 
      Caption         =   "We've found this existing component that matches the name of the component you're trying to import into the database:"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Qnt:"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Sub-Category:"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Package:"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "frmDuplicateComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmDuplicateComponent
''' Duplicate component dialog box.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private enumerations.
Public Enum Action
    actNothing
    actUpdateQuantity
    actImportAnyway
End Enum

' Private variables.
Private m_actAction As Action
Private m_lngID As Long
Private m_lngQuantity As Long

' Positions this dialog by the side of an anchor frame in the parent window.
Public Sub PositionBySide(frmParent As Form, fraAnchor As Frame)
    Top = frmParent.Top + fraAnchor.Top + 800
    Left = frmParent.Left + frmParent.Width + 150
    
    ' Check if we should move to the right side of the parent form.
    If Screen.Width < Left Then
        Left = frmParent.Left - Width - 150
    End If
End Sub

' Populate Form from Recordset.
Public Sub PopulateFromRecordset(rs As ADODB.Recordset)
    Dim intIndex As Integer
    
    ' Store the component ID and quantity.
    ID = rs.Fields("ID")
    Quantity = rs.Fields("Quantity")
    
    ' Set text fields.
    txtName.Text = rs.Fields("Name")
    txtQuantity.Text = rs.Fields("Quantity")
    txtNotes.Text = rs.Fields("Notes")
    
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
    
    ' Show component image.
    ShowImage rs.Fields("Name")
End Sub

' Shows the component image.
Private Sub ShowImage(strName As String)
    Dim strImage As String
    
    ' Get component image.
    strImage = GetComponentImagePath(strName, cmbPackage.Text)
    
    ' Set the component image.
    If strImage <> vbNullString Then
        Dim picBitmap As Picture
        On Error GoTo PictureError
        Set picBitmap = LoadPicture(strImage)
        
        picImage.AutoRedraw = True
        picImage.PaintPicture picBitmap, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight
        Set picImage.Picture = picImage.Image
    Else
        Set picImage.Picture = Nothing
    End If
    
    ' Handle image setting errors.
    Exit Sub
PictureError:
    Set picImage.Picture = Nothing
    MsgBox "An error occured while trying to load the image for this component.", _
        vbOKOnly + vbCritical, "Image Loading Error"
End Sub

' Cancel the whole thing.
Private Sub cmdCancel_Click()
    DoAction = actNothing
    Unload Me
End Sub

' Import this thing anyway.
Private Sub cmdImportAnyway_Click()
    DoAction = actImportAnyway
    Unload Me
End Sub

' Just update the quantity.
Private Sub cmdUpdateQuantity_Click()
    DoAction = actUpdateQuantity
    Unload Me
End Sub

' Form just loaded up.
Private Sub Form_Load()
    ' Reset variables.
    ID = -1
    DoAction = actNothing
    Quantity = 0
End Sub

' Gets the action to be taken.
Public Property Get DoAction() As Action
    DoAction = m_actAction
End Property

' Sets the action to be taken.
Public Property Let DoAction(actAction As Action)
    m_actAction = actAction
End Property

' Gets the duplicate component ID.
Public Property Get ID() As Long
    ID = m_lngID
End Property

' Sets the duplicate component ID.
Public Property Let ID(lngID As Long)
    m_lngID = lngID
End Property

' Gets the quantity of the duplicate component.
Public Property Get Quantity() As Long
    Quantity = m_lngQuantity
End Property

' Sets the quantity of the duplicate component.
Public Property Let Quantity(lngQuantity As Long)
    m_lngQuantity = lngQuantity
End Property
