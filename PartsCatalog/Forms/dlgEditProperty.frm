VERSION 5.00
Begin VB.Form dlgEditProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Property"
   ClientHeight    =   2055
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3495
   Icon            =   "dlgEditProperty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Caption         =   "Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "dlgEditProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' dlgEditProperty
''' A simple dialog to edit a component property.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_lngID As Long
Private m_strName As String
Private m_strValue As String
Private m_lngComponentID As Long
Private m_blnChanged As Boolean

' Shows the dialog as a new property.
Public Sub ShowNew(lngComponentID As Long)
    ' Setup values and controls.
    ID = -1
    ComponentID = lngComponentID
    Key = ""
    Value = ""
    Me.Caption = "New Property"
    OKButton.Caption = "Add"
    
    ' Show the modal dialog.
    Show vbModal
End Sub

' Shows the dialog as a property editor.
Public Sub ShowEditor(lngComponentID As Long, lngID As Long)
    ' Store IDs.
    ID = lngID
    ComponentID = lngComponentID
    
    ' Setup controls.
    Me.Caption = "Edit Property"
    OKButton.Caption = "Save"
    
    ' Populate the dialog and display as modal.
    LoadProperty ID, txtName, txtValue
    Show vbModal
End Sub

' Property ID getter.
Public Property Get ID() As Long
    ID = m_lngID
End Property

' Property ID setter.
Public Property Let ID(lngID As Long)
    m_lngID = lngID
End Property

' Property name getter.
Public Property Get Key() As String
    Key = m_strName
End Property

' Property name setter.
Public Property Let Key(strName As String)
    m_strName = strName
    txtName.Text = strName
End Property

' Property value getter.
Public Property Get Value() As String
    Value = m_strValue
End Property

' Property value setter.
Public Property Let Value(strValue As String)
    m_strValue = strValue
    txtValue.Text = strValue
End Property

' Component ID getter.
Public Property Get ComponentID() As Long
    ComponentID = m_lngComponentID
End Property

' Component ID setter.
Public Property Let ComponentID(lngComponentID As Long)
    m_lngComponentID = lngComponentID
End Property

' Should we save the property?
Public Property Get Changed() As Boolean
    Changed = m_blnChanged
End Property

' Set the save property flag.
Private Property Let Changed(blnChanged As Boolean)
    m_blnChanged = blnChanged
End Property

' Cancel button just got clicked.
Private Sub CancelButton_Click()
    Changed = False
    Unload Me
End Sub

' Dialog just loaded.
Private Sub Form_Load()
    Changed = False
End Sub

' Changed button got clicked.
Private Sub OKButton_Click()
    ' Set the name and value.
    Key = txtName.Text
    Value = txtValue.Text
    
    ' Save the edited property to the database.
    ID = SaveProperty(ID, Key, Value, ComponentID)
    
    ' Set the changed flag and exit.
    Changed = True
    Unload Me
End Sub
