VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmComponent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Component"
   ClientHeight    =   5865
   ClientLeft      =   6135
   ClientTop       =   3315
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   741
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolBar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Description     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Description     =   "Delete"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Duplicate"
            Description     =   "Duplicate"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "KeepOpen"
            Description     =   "Keep Window Opened"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.StatusBar stbStatusBar 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   5535
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picImage 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   15
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdDatasheet 
      Caption         =   "Datasheet"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grdProperties 
      Height          =   2775
      Left            =   2280
      TabIndex        =   12
      Top             =   2760
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.TextBox txtNotes 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Text            =   "Notes"
      Top             =   1920
      Width           =   6135
   End
   Begin VB.ComboBox cmbPackage 
      Height          =   315
      Left            =   6840
      TabIndex        =   7
      Text            =   "Package"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox cmbSubCategory 
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Text            =   "Sub-Category"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Text            =   "Category"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Text            =   "Component Name"
      Top             =   720
      Width           =   5295
   End
   Begin VB.TextBox txtQuantity 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Text            =   "000000"
      Top             =   720
      Width           =   735
   End
   Begin ComctlLib.ImageList imlToolBar 
      Left            =   960
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmComponent.frx":0000
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmComponent.frx":0112
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmComponent.frx":0224
            Key             =   "Duplicate"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmComponent.frx":0336
            Key             =   "KeepOpen"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Properties:"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Package:"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Sub-Category:"
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Qnt:"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   480
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
Private m_blnKeepOpen As Boolean

' Populate Form from Recordset.
Public Sub PopulateFromRecordset(rs As ADODB.Recordset)
    Dim intIndex As Integer
    
    ' Store the component ID.
    m_lngComponentID = rs.Fields("ID")
    
    ' Set text fields.
    txtQuantity.Text = rs.Fields("Quantity")
    txtName.Text = rs.Fields("Name")
    txtNotes.Text = rs.Fields("Notes")
    SetStatusMessage "Loaded text fields"
    
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
    SetStatusMessage "Loaded categories"
    
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
    SetStatusMessage "Loaded sub-categories"
    
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
    SetStatusMessage "Loaded component packages"
    
    ' Populate the properties grid.
    If Not IsNull(rs.Fields("Properties")) Then
        PopulatePropertiesGrid rs.Fields("Properties")
        SetStatusMessage "Loaded properties"
    Else
        grdProperties.Rows = 2
        grdProperties.TextMatrix(1, 0) = ""
        grdProperties.TextMatrix(1, 1) = ""
        SetStatusMessage "No component properties to load"
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
        SetStatusMessage "Component image loaded"
    End If
    
    Exit Sub
PictureError:
    Set picImage.Picture = Nothing
    MsgBox "An error occured while trying to load the image for this component.", _
        vbOKOnly + vbCritical, "Image Loading Error"
    SetStatusMessage "ERROR: Couldn't load the image for this component."
End Sub

' Populates the properties grid.
Public Sub PopulatePropertiesGrid(strProperties As String)
    Dim astrProperties() As String
    Dim astrKeyValue() As String
    
    ' Split the properties and preparate the grid for the properties.
    astrProperties = Split(strProperties, vbTab)
    grdProperties.Rows = UBound(astrProperties) + 2
    
    ' Populate the properties.
    Dim intIndex As Integer
    For intIndex = 0 To UBound(astrProperties)
        ' Check if the property is populated.
        If astrProperties(intIndex) <> "" Then
            astrKeyValue = Split(astrProperties(intIndex), ": ")
            grdProperties.TextMatrix(intIndex + 1, 0) = astrKeyValue(0)
            grdProperties.TextMatrix(intIndex + 1, 1) = astrKeyValue(1)
        Else
            grdProperties.TextMatrix(1, 0) = ""
            grdProperties.TextMatrix(1, 1) = ""
        End If
    Next intIndex
End Sub

' Encodes the properties grid into a string to be stored in the database.
Public Function EncodePropertiesGrid() As String
    EncodePropertiesGrid = ""
End Function

' Sets a status message in the statusbar.
Private Sub SetStatusMessage(strMessage As String)
    stbStatusBar.SimpleText = strMessage
End Sub

' Make sure the form is kept opened.
Private Sub MaintainFormOpened(Optional blnState As Boolean = True)
    m_blnKeepOpen = blnState
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
    ' Clear the opened status.
    m_blnKeepOpen = False

    ' Setup the Flex Grid.
    grdProperties.TextMatrix(0, 0) = "Property"
    grdProperties.TextMatrix(0, 1) = "Value"
    grdProperties.ColWidth(0) = (grdProperties.Width / 2) - 45
    grdProperties.ColWidth(1) = (grdProperties.Width / 2) - 45
End Sub

' Handles the toolbar button clicks.
Private Sub tlbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Save"
            MsgBox "Save"
        Case "Delete"
            MsgBox "Delete"
        Case "Duplicate"
            MsgBox "Duplicate"
        Case "KeepOpen"
            MaintainFormOpened (Button.Value = tbrPressed)
    End Select
End Sub

' Name text change event.
Private Sub txtName_Change()
    Me.Caption = txtName.Text
End Sub

' Getter for maintaining the window opened.
Public Property Get StayOpen() As Boolean
    StayOpen = m_blnKeepOpen
End Property
