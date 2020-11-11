Attribute VB_Name = "modFormHelper"
''' modFormHelper
''' A collection of routines to help handling forms.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

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

