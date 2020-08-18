VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Component"
   ClientHeight    =   5085
   ClientLeft      =   6135
   ClientTop       =   3375
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picImage 
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   15
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdDatasheet 
      Caption         =   "Datasheet"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grdProperties 
      Height          =   2775
      Left            =   2160
      TabIndex        =   12
      Top             =   2280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.TextBox txtNotes 
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Text            =   "Notes"
      Top             =   1440
      Width           =   6135
   End
   Begin VB.ComboBox cmbPackage 
      Height          =   315
      Left            =   6720
      TabIndex        =   7
      Text            =   "Package"
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox cmbSubCategory 
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      Text            =   "Sub-Category"
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Text            =   "Category"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Text            =   "Component Name"
      Top             =   240
      Width           =   5295
   End
   Begin VB.TextBox txtQuantity 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "000000"
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Properties:"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Package:"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Sub-Category:"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Qnt:"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   0
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

' Populate Form from Recordset.
Public Sub PopulateFromRecordset(rs As ADODB.Recordset)
    Dim intIndex As Integer
    
    ' Store the component ID.
    m_lngComponentID = rs.Fields("ID")
    
    ' Set text fields.
    txtQuantity.Text = rs.Fields("Quantity")
    txtName.Text = rs.Fields("Name")
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
    End If
    
    Exit Sub
PictureError:
    Set picImage.Picture = Nothing
    MsgBox "An error occured while trying to load the image for this component.", _
        vbOKOnly + vbCritical, "Image Loading Error"
End Sub

' Category selection updated.
Private Sub cmbCategory_Click()
    If cmbSubCategory.ListCount > 0 Then
        LoadSubCategories cmbCategory.ItemData(cmbCategory.ListIndex), _
            cmbSubCategory, True
    End If
End Sub

' Form just loaded up.
Private Sub Form_Load()
    ' Setup the Flex Grid.
    grdProperties.TextMatrix(0, 0) = "Property"
    grdProperties.TextMatrix(0, 1) = "Value"
    grdProperties.ColWidth(0) = (grdProperties.Width / 2) - 45
    grdProperties.ColWidth(1) = (grdProperties.Width / 2) - 45
End Sub
