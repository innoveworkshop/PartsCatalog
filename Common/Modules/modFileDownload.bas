Attribute VB_Name = "modFileDownload"
''' modFileDownload
''' File download helper module.
'''
''' Author: Nathan Campos <nathan@innoveworkshop.com>
Option Explicit

' Import the API download function.
Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
    Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

' Constants that the API function returns.
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

' Downloads a file from a source to a destination.
Public Function DownloadFile(strURL As String, strDestination As String) As Boolean
    DeleteUrlCacheEntry strURL
    DownloadFile = URLDownloadToFile(0&, strURL, strDestination, _
        BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
End Function
