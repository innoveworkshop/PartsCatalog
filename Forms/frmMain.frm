VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Parts Catalogger"
   ClientHeight    =   8955
   ClientLeft      =   7365
   ClientTop       =   3675
   ClientWidth     =   14520
   Begin MSComctlLib.StatusBar stbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   8640
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mniFileOpenDatabase 
         Caption         =   "&Open Database..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mniFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mniTestForm 
      Caption         =   "TestForm"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmMain
''' Main application form.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Event fired when the form loads up.
Private Sub MDIForm_Load()
    SetDatabasePath "C:\Documents and Settings\Administrator\My Documents\PartCat Experiment\PartCat.mdb"
    frmPartChooser.SetParent Me
    frmPartChooser.Show
End Sub

' Form resized event.
Private Sub MDIForm_Resize()
    frmPartChooser.ResizeToFitParent
End Sub

' Event fired when the form is about to be closed.
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frmForm As Form
    
    For Each frmForm In Forms
        Unload frmForm
    Next
End Sub

' Exits the application.
Private Sub mniFileExit_Click()
    Unload Me
End Sub
