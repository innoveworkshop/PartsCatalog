Attribute VB_Name = "modAppIcon"
''' modAppIcon
''' A module to properly set the application icon.
'''
''' Source: Providing a proper VB Application Icon <http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp>
''' Author: Nathan Campos <nathan@innnoveworkshop.com>

Option Explicit

' Win32 API constants.
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Private Const GW_OWNER = 4

' Win32 API imports.
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" _
    (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, _
    ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, _
    ByVal wCmd As Long) As Long

' Sets the icon via a resource.
Public Sub SetIcon(ByVal hWnd As Long, ByVal sIconResName As String, _
        Optional ByVal bSetAsAppIcon As Boolean = True)
    Dim lhWndTop As Long
    Dim lhWnd As Long
    Dim cx As Long
    Dim cy As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long

    If bSetAsAppIcon Then
        ' Find VB's hidden parent window.
        lhWnd = hWnd
        lhWndTop = lhWnd
        Do While Not (lhWnd = 0)
            lhWnd = GetWindow(lhWnd, GW_OWNER)
            If Not (lhWnd = 0) Then
                lhWndTop = lhWnd
            End If
        Loop
    End If

    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, _
        cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    End If
    SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
    
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, _
          cx, cy, LR_SHARED)
    If bSetAsAppIcon Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    End If
    SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
End Sub
