Attribute VB_Name = "modFormHelper"
''' modFormHelper
''' A collection of routines to help handling forms.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Win32 API stuff.
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

' Retrieves the owner of a child form.
Public Function ParentForm(frmChild As Form) As Form
    Dim frmParent As Form
    Dim hwndParent As Long
    
    ' Get the parent window handle.
    hwndParent = GetWindow(frmChild.hwnd, GW_OWNER)
    
    ' Try to find a matching form in our application.
    For Each frmParent In Forms
        If frmParent.hwnd = hwndParent Then
            Set ParentForm = frmParent
    Next frmParent
    
    ' No owner form found.
    Set ParentForm = Nothing
End Function

' Centralizes the dialog in the middle of the called form.
Public Sub CentralizeFormInForm(frmChild As Form, frmParent As Form)
    frmChild.Top = frmParent.Top + (frmParent.Height / 2) - (frmChild.Height / 2)
    frmChild.Left = frmParent.Left + (frmParent.Width / 2) - (frmChild.Width / 2)
End Sub

' Centralizes the dialog in the middle of an MDI child form.
Public Sub CentralizeFormInMDIChild(frmChild As Form, frmMDIParent As MDIForm, _
        frmParent As Form)
    frmChild.Top = frmMDIParent.Top + frmParent.Top + (frmParent.Height / 2)
    frmChild.Left = frmMDIParent.Left + frmParent.Left + (frmParent.Width / 2) - _
        (frmChild.Width / 2)
End Sub

