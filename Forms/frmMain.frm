VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Parts Catalogger"
   ClientHeight    =   8955
   ClientLeft      =   7365
   ClientTop       =   3675
   ClientWidth     =   14520
   Begin ComctlLib.StatusBar stbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   8640
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
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
