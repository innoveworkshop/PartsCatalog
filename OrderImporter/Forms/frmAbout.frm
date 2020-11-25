VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3090
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2132.773
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   240
      Picture         =   "frmAbout.frx":6852
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   240
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2145
      Width           =   1260
   End
   Begin VB.CommandButton cmdWebsite 
      Caption         =   "&Website"
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   2595
      Width           =   1245
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright whatever @ some year"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1356.278
      Y2              =   1356.278
   End
   Begin VB.Label lblDescription 
      Caption         =   "A description."
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   1366.631
      Y2              =   1366.631
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblDeveloper 
      Caption         =   "Some text about the developer."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   240
      TabIndex        =   4
      Top             =   2145
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmAbout
''' A pretty standard VB6 application about form.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Win32 API imports.
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' Private variables.
Private m_frmParent As Form

' Centralizes the dialog in the middle of the called form.
Private Sub CentralizeForm()
    Top = Parent.Top + (Parent.Height / 2) - (Height / 2)
    Left = Parent.Left + (Parent.Width / 2) - (Width / 2)
End Sub

' User clicked the OK button.
Private Sub cmdOK_Click()
    Unload Me
End Sub

' User clicked the Website button.
Private Sub cmdWebsite_Click()
    ShellExecute Parent.hwnd, vbNullString, "http://innoveworkshop.com/", _
        vbNullString, vbNullString, vbNormalFocus
End Sub

' Dialog just loaded.
Private Sub Form_Load()
    ' Set all the textual stuff.
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = App.Comments
    lblDeveloper.Caption = "This application was developed by Nathan Campos for " & _
        App.CompanyName & "."
    lblCopyright.Caption = App.LegalCopyright
    
    ' Move the form to the right place.
    CentralizeForm
End Sub

' Gets the parent form.
Public Property Get Parent() As Form
    Set Parent = m_frmParent
End Property

' Sets the parent form.
Public Property Let Parent(frmParent As Form)
    Set m_frmParent = frmParent
End Property
