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
Private m_strName As String
Private m_strValue As String
Private m_blnSave As Boolean

' Shows the dialog as a new property.
Public Sub ShowNew()
    ' Setup values and controls.
    Key = ""
    Value = ""
    Me.Caption = "New Property"
    OKButton.Caption = "Add"
    
    ' Show the modal dialog.
    Show vbModal
End Sub

' Shows the dialog as a property editor.
Public Sub ShowEditor(strName As String, strValue As String)
    ' Setup values and controls.
    Key = strName
    Value = strValue
    Me.Caption = "Edit Property"
    OKButton.Caption = "Save"
    
    ' Show the modal dialog.
    Show vbModal
End Sub

' Centralizes the dialog in the middle of the called form.
Public Sub CentralizeInForm(frmMother As MDIForm, frmParent As Form)
    Me.Top = frmMother.Top + frmParent.Top + (frmParent.Height / 2)
    Me.Left = frmMother.Left + frmParent.Left + (frmParent.Width / 2) - _
        (Me.Width / 2)
End Sub

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

' Should we save the property?
Public Property Get Save() As Boolean
    Save = m_blnSave
End Property

' Set the save property flag.
Private Property Let Save(blnSave As Boolean)
    m_blnSave = blnSave
End Property

' Cancel button just got clicked.
Private Sub CancelButton_Click()
    Save = False
    Unload Me
End Sub

' Dialog just loaded.
Private Sub Form_Load()
    Save = False
End Sub

' Save button got clicked.
Private Sub OKButton_Click()
    Key = txtName.Text
    Value = txtValue.Text
    
    Save = True
    Unload Me
End Sub
