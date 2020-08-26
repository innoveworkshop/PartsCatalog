VERSION 5.00
Begin VB.Form frmPackageManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Package Manager"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frmPackageManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   3600
      ScaleHeight     =   2865
      ScaleWidth      =   2985
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.ListBox lstPackages 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      Caption         =   "000"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblName 
      Caption         =   "Package Name:"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPackageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmPackageManager
''' A simple way to manage your component packages.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_lngID As Long

' Check if we are dealing with a new package or editing an existing one.
Public Function IsNewPackage() As Boolean
    If ID = -1 Then
        IsNewPackage = True
    Else
        IsNewPackage = False
    End If
End Function

' Updates the enabled/disabled controls.
Private Sub UpdateEnabledControls()
    cmdRename.Enabled = Not IsNewPackage
    cmdRemove.Enabled = Not IsNewPackage
End Sub

' Shows the selected package image.
Private Sub ShowImage()
    Dim strImage As String
    
    ' Get package image.
    strImage = GetComponentImagePath(vbNullString, txtName.Text)
    
    ' Set image.
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
    MsgBox "An error occured while trying to load the image for this package.", _
        vbOKOnly + vbCritical, "Image Loading Error"
End Sub

' Add a new package.
Private Sub cmdAdd_Click()
    Dim intIndex As Integer

    ' Add package.
    ID = SavePackage(-1, txtName.Text)
    
    ' Populate packages list and select the newly added item.
    LoadPackages lstPackages
    For intIndex = 0 To lstPackages.ListCount - 1
        If lstPackages.ItemData(intIndex) = ID Then
            lstPackages.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Remove the package.
Private Sub cmdRemove_Click()
    ' Check if we can do this.
    If IsNewPackage Then
        Exit Sub
    End If
    
    ' Delete package.
    DeletePackage ID
    
    ' Reset the ID, name, and repopulate the packages list.
    ID = -1
    txtName.Text = ""
    LoadPackages lstPackages
End Sub

' Rename the package.
Private Sub cmdRename_Click()
    Dim intIndex As Integer

    ' Check if we can do this.
    If IsNewPackage Then
        Exit Sub
    End If
    
    ' Update the package, reset the ID.
    ID = SavePackage(ID, txtName.Text)
    
    ' Populate packages list and select the newly added item.
    LoadPackages lstPackages
    For intIndex = 0 To lstPackages.ListCount - 1
        If lstPackages.ItemData(intIndex) = ID Then
            lstPackages.ListIndex = intIndex
            Exit Sub
        End If
    Next intIndex
End Sub

' Form just loaded up.
Private Sub Form_Load()
    ' Reset variables.
    ID = -1
    
    ' Populate packages list and update controls.
    LoadPackages lstPackages
    UpdateEnabledControls
End Sub

' Package selection changed.
Private Sub lstPackages_Click()
    ' Update the text field and ID.
    txtName.Text = lstPackages.Text
    ID = lstPackages.ItemData(lstPackages.ListIndex)
    
    ' Show package image and update controls.
    ShowImage
    UpdateEnabledControls
End Sub

' Gets the package ID.
Public Property Get ID() As Long
    ID = m_lngID
End Property

' Sets the package ID.
Public Property Let ID(lngID As Long)
    m_lngID = lngID
    
    ' Sets the ID label.
    If lngID <> -1 Then
        lblID.Caption = lngID
    Else
        lblID.Caption = ""
    End If
    
    ' Update the controls.
    UpdateEnabledControls
End Property
