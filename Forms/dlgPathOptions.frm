VERSION 5.00
Begin VB.Form dlgPathOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Path Options"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDatabase 
      Caption         =   "Database"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
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
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
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

' Form just loaded.
Private Sub Form_Load()
    
End Sub
