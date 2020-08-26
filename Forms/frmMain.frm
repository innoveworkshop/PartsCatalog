VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Parts Catalogger"
   ClientHeight    =   8955
   ClientLeft      =   7365
   ClientTop       =   3675
   ClientWidth     =   14520
   Icon            =   "frmMain.frx":0000
   Begin MSComDlg.CommonDialog dlgOpenDatabase 
      Left            =   12840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mdb"
      DialogTitle     =   "Open Database"
      Filter          =   "Microsoft Access Databases (*.mdb)|*.mdb|All Files (*.*)|*.*"
   End
   Begin MSComctlLib.ImageList imlToolBar 
      Left            =   13440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":209DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2723C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34300
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenDatabase"
            Object.ToolTipText     =   "Open Database"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CloseDatabase"
            Object.ToolTipText     =   "Close Database"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReloadDatabase"
            Object.ToolTipText     =   "Reload Database"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Categories"
            Object.ToolTipText     =   "Manage (Sub-)Categories"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Packages"
            Object.ToolTipText     =   "Manage Packages"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Components"
            Object.ToolTipText     =   "Manage Components"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddComponent"
            Object.ToolTipText     =   "Add Component"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
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
      Begin VB.Menu mniFileReloadDatabase 
         Caption         =   "&Reload Database"
         Shortcut        =   ^R
      End
      Begin VB.Menu mniFileCloseDatabase 
         Caption         =   "&Close Database"
         Shortcut        =   ^W
      End
      Begin VB.Menu mniFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuManage 
      Caption         =   "&Manage"
      Begin VB.Menu mniManageCategories 
         Caption         =   "C&ategories..."
      End
      Begin VB.Menu mniManagePackages 
         Caption         =   "&Packages..."
      End
      Begin VB.Menu mniManageComponents 
         Caption         =   "&Components..."
      End
      Begin VB.Menu mniManageSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mniManageAddComponent 
         Caption         =   "A&dd Component..."
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mniHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
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

' Shows the package manager.
Private Sub ManagePackages()
    frmPackageManager.Left = frmPackageManager.Left + frmPartChooser.Width
    frmPackageManager.Show
End Sub

' Shows a component dialog for creating a new component.
Private Sub NewComponent()
    Dim frmNewComponent As frmComponent
    Set frmNewComponent = New frmComponent
    
    ' Show the new component form and remove its reference.
    frmNewComponent.ShowNewComponent
    Set frmNewComponent = Nothing
End Sub

' Open a new database.
Private Sub OpenDatabaseFile(Optional strPath As String = vbNullString)
    Dim strSetPath As String
    strSetPath = strPath
    
    ' Check if we should use the open dialog.
    If strPath = vbNullString Then
        dlgOpenDatabase.ShowOpen
        strSetPath = dlgOpenDatabase.FileName
    End If
    
    ' Set the database path.
    If strSetPath <> vbNullString Then
        SetDatabasePath strSetPath
        SetLastUsedDatabasePath strSetPath
    End If
End Sub

' Clears the whole thing.
Private Sub CloseDatabase()
    ClearDatabasePath
    SetLastUsedDatabasePath vbNullString
    CloseAllChilds
    frmPartChooser.ClearContents
End Sub

' Reloads the database and updates everything in the application.
Private Sub ReloadDatabase()
    Dim frmForm As Form

    ' Reload component forms.
    For Each frmForm In Forms
        If frmForm.Name = "frmComponent" Then
            frmForm.Refresh
        End If
    Next frmForm
    
    ' Reload part chooser form.
    frmPartChooser.RefreshLists
End Sub

' Closes all the child windows that aren't panels.
Private Sub CloseAllChilds()
    Dim frmForm As Form
    
    For Each frmForm In Forms
        If frmForm.Name = "frmComponent" Then
            Unload frmForm
        End If
    Next frmForm
End Sub

' Event fired when the form loads up.
Private Sub MDIForm_Load()
    ' Open the last used database.
    OpenDatabaseFile LastUsedDatabasePath
    
    ' Setup the parts chooser panel.
    frmPartChooser.SetParent Me
    frmPartChooser.Show
End Sub

' Form is about to be unloaded.
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim frmForm As Form
    Dim frmComp As frmComponent
    
    ' Go through component forms checking if they have unsaved changes.
    For Each frmForm In Forms
        If frmForm.Name = "frmComponent" Then
            Set frmComp = frmForm
            If frmComp.AbortUnsavedChanges Then
                Cancel = 1
                Set frmComp = Nothing
                
                Exit Sub
            End If
            
            Set frmComp = Nothing
        End If
    Next
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

' File > Close Database menu clicked.
Private Sub mniFileCloseDatabase_Click()
    CloseDatabase
End Sub

' File > Exit menu clicked.
Private Sub mniFileExit_Click()
    Unload Me
End Sub

' File > Open Database menu clicked.
Private Sub mniFileOpenDatabase_Click()
    OpenDatabaseFile
End Sub

' File > Reload Database menu clicked.
Private Sub mniFileReloadDatabase_Click()
    ReloadDatabase
End Sub

' Help > About menu clicked.
Private Sub mniHelpAbout_Click()
    frmAbout.Parent = Me
    frmAbout.Show vbModal, Me
End Sub

' Manage > Add Component menu clicked.
Private Sub mniManageAddComponent_Click()
    NewComponent
End Sub

' Manage > Packages menu clicked.
Private Sub mniManagePackages_Click()
    ManagePackages
End Sub

' Toolbar button clicked event.
Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "OpenDatabase"
            OpenDatabaseFile
        Case "ReloadDatabase"
            ReloadDatabase
        Case "CloseDatabase"
            CloseDatabase
        Case "Categories"
            MsgBox "Categories"
        Case "Packages"
            ManagePackages
        Case "Components"
            MsgBox "Components"
        Case "AddComponent"
            NewComponent
    End Select
End Sub
