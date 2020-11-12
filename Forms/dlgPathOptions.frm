VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dlgPathOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Path Options"
   ClientHeight    =   2895
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6105
   Icon            =   "dlgPathOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAppSuite 
      Caption         =   "Application Suite"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5895
      Begin VB.TextBox txtOrderImporter 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   5175
      End
      Begin VB.CommandButton cmdBrowseOrderImporter 
         Caption         =   "..."
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Order Importer:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Database"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton cmdBrowseLastDatabase 
         Caption         =   "..."
         Height          =   315
         Left            =   5400
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtLastDatabase 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Default Database:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "dlgPathOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' dlgPathOptions
''' A little settings dialog with options to set different path related stuff.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Private variables.
Private m_strLastDatabasePath As String

' Show this dialog as a modal in the middle of the parent form.
Public Sub ShowModal(frmParent As Form)
    CentralizeFormInForm Me, frmParent
    Show vbModal, frmParent
End Sub

' Save all the settings.
Private Sub SaveSettings()
    SetLastUsedDatabasePath txtLastDatabase.Text
    SetOrderImporterPath txtOrderImporter.Text
End Sub

' Browse for the default database.
Private Sub BrowseLastDatabase()
    ' Setup the dialog.
    dlgOpen.DefaultExt = "mdb"
    dlgOpen.DialogTitle = "Select Default Database"
    dlgOpen.Filter = "Microsoft Access Databases (*.mdb)|*.mdb|All Files (*.*)|*.*"

    ' Open the dialog and set the selected path.
    dlgOpen.ShowOpen
    If dlgOpen.FileName <> vbNullString Then
        txtLastDatabase.Text = dlgOpen.FileName
    End If
End Sub

' Browse for the default database.
Public Sub BrowseOrderImporter()
    ' Setup the dialog.
    dlgOpen.DefaultExt = "exe"
    dlgOpen.DialogTitle = "Find the Order Importer Executable"
    dlgOpen.Filter = "Executable File (*.exe)|*.exe|All Files (*.*)|*.*"

    ' Open the dialog and set the selected path.
    dlgOpen.ShowOpen
    If dlgOpen.FileName <> vbNullString Then
        txtOrderImporter.Text = dlgOpen.FileName
    End If
End Sub

' Cancel button clicked.
Private Sub CancelButton_Click()
    Unload Me
End Sub

' Browse for the last database used.
Private Sub cmdBrowseLastDatabase_Click()
    BrowseLastDatabase
End Sub

' Browse for the Order Importer executable.
Private Sub cmdBrowseOrderImporter_Click()
    BrowseOrderImporter
End Sub

' Form just loaded.
Private Sub Form_Load()
    ' Load settings.
    txtLastDatabase.Text = LastUsedDatabasePath
    txtOrderImporter.Text = OrderImporterPath
    
    ' Set private variables.
    m_strLastDatabasePath = txtLastDatabase.Text
End Sub

' OK button clicked.
Private Sub OKButton_Click()
    Dim intResponse As Integer
    
    ' Save settings and close the dialog.
    SaveSettings
    Unload Me
    
    ' Check if the user changed the default database and ask for a reload.
    If m_strLastDatabasePath <> LastUsedDatabasePath Then
        intResponse = MsgBox("You've changed the default database path. Do you wish " & _
            "to reload the database and start using the new one?", vbYesNo + vbQuestion, _
            "Default Database Changed")
        
        ' Reload the databse if the user wants to.
        If intResponse = vbYes Then
            frmMain.ReloadDatabase LastUsedDatabasePath
        End If
    End If
End Sub
