VERSION 5.00
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
         Caption         =   "Last Used Database:"
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

' Cancel button clicked.
Private Sub CancelButton_Click()
    Unload Me
End Sub

' Form just loaded.
Private Sub Form_Load()
    ' Load settings.
    txtLastDatabase.Text = LastUsedDatabasePath
    txtOrderImporter.Text = OrderImporterPath
End Sub

' OK button clicked.
Private Sub OKButton_Click()
    SaveSettings
    Unload Me
End Sub
