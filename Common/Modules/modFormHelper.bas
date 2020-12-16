Attribute VB_Name = "modFormHelper"
''' modFormHelper
''' A collection of routines to help handling forms.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Win32 API stuff.
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, _
    ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

' Retrieves the owner of a child form.
Public Function ParentForm(frmChild As Form) As Form
    Dim frmParent As Form
    Dim hwndParent As Long
    
    ' Get the parent window handle.
    hwndParent = GetWindow(frmChild.hWnd, GW_OWNER)
    
    ' Try to find a matching form in our application.
    For Each frmParent In Forms
        If frmParent.hWnd = hwndParent Then
            Set ParentForm = frmParent
        End If
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

' If an MDI form is opened behind the side panel nudge it to the right.
Public Sub FormDodgeSidePanel(frmChild As Form, frmSidePanel As Form, _
        frmMDIParent As MDIForm)
    Dim sngLeft As Single
    Dim sngPastParentEnd As Single
    
    ' Check if we need to do something in the first place.
    If frmChild.Left > (frmSidePanel.Left + frmSidePanel.Width) Then
        Exit Sub
    End If
    
    ' Calculate positions.
    sngLeft = frmChild.Left + frmSidePanel.Width
    sngPastParentEnd = (sngLeft + frmChild.Width) - frmMDIParent.Width
    
    ' Check if we need to nudge it back a bit because we've gone too far right.
    If sngPastParentEnd > 0 Then
        sngLeft = sngLeft - sngPastParentEnd - 200
    End If
    
    ' Set the childs Left position.
    frmChild.Left = sngLeft
End Sub
