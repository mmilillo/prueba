Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function DownloadFile(sURLFile As String, sLocalFilename As String) As Boolean
    Dim lRetVal As Long
      
    lRetVal = URLDownloadToFile(0, sURLFile, sLocalFilename, 0, 0)
    If lRetVal = 0 Then DownloadFile = True
End Function


