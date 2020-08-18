VERSION 5.00
Begin VB.Form frmPartChooser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Component Selector"
   ClientHeight    =   8190
   ClientLeft      =   4500
   ClientTop       =   3165
   ClientWidth     =   3855
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   3855
   Begin VB.ListBox lstComponents 
      Height          =   2400
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   3855
   End
   Begin VB.ListBox lstSubCategories 
      Height          =   2400
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   3855
   End
   Begin VB.ListBox lstCategories 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Components:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Sub-Categories:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Categories:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmPartChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
