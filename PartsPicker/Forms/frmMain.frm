VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Parts Picker"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraComponents 
      Caption         =   "Components"
      Height          =   5535
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdComponentSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   5040
         Width           =   3255
      End
      Begin VB.CommandButton cmdComponentRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton cmdComponentAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton cmdRefDesRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5400
         TabIndex        =   20
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefDesRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefDesAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtRefDes 
         Height          =   315
         Left            =   5400
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ListBox lstRefDes 
         Height          =   3180
         Left            =   3720
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cmbComponent 
         Height          =   315
         Left            =   3720
         TabIndex        =   12
         Top             =   480
         Width           =   3255
      End
      Begin VB.ListBox lstComponents 
         Height          =   4740
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0007
         TabIndex        =   10
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Reference Designators:"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblDescription 
         Caption         =   "COMPONENT DESCRIPTION GOES IN HERE"
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   840
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblComponentID 
         Alignment       =   1  'Right Justify
         Caption         =   "00000"
         Height          =   255
         Left            =   6480
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Component:"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraProject 
      Caption         =   "Project Information"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   3495
      Begin VB.TextBox txtCategoryName 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdCategoryAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmdCategoryRename 
         Caption         =   "Rename"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdCategoryRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCategoryID 
         Alignment       =   1  'Right Justify
         Caption         =   "000"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.ListBox lstProjects 
      Height          =   3180
      ItemData        =   "frmMain.frx":001A
      Left            =   120
      List            =   "frmMain.frx":0021
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Projects:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' frmMain
''' Main form of the Parts Picker application.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit
