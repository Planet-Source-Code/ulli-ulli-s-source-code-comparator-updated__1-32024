Attribute VB_Name = "mAPI"
Option Explicit

DefLng A-Z 'we're 32 bit!
 
'Splash duration
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds)

'Get temp file name for saving the current source
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

'About box
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'Send Mail
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL          As Long = 1
Public Const SE_NO_ERROR            As Long = 33 'Values below 33 are error returns

'VB tab width and font
Public Const VBSettings             As String = "Software\Microsoft\VBA\Microsoft Visual Basic"
Public Const Fontface               As String = "Fontface"
Public Const Fontheight             As String = "Fontheight"
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey, ByVal lpSubKey As String, ByVal ulOptions, ByVal samDesired, phkResult) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey, ByVal lpValueName As String, ByVal lpReserved, lpType, lpData As Any, lpcbData) As Long
Public Declare Sub RegCloseKey Lib "advapi32.dll" (ByVal hKey)
Public Const HKEY_CURRENT_USER      As Long = &H80000001
'Public Const HKEY_LOCAL_MACHINE     As Long = &H80000002
Public Const KEY_QUERY_VALUE        As Long = 1
Public Const REG_OPTION_RESERVED    As Long = 0
Public Const ERROR_NONE             As Long = 0

Public Function AppDetails() As String

    With App
        AppDetails = .ProductName & " Version " & .Major & "." & .Minor & "." & .Revision
    End With 'APP

End Function

Public Sub SendMeMail(FromhWnd, Subject As String)

    If ShellExecute(FromhWnd, vbNullString, "mailto:UMGEDV@AOL.COM?subject=" & Subject & " &body=Hi Ulli,", vbNullString, App.Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
        Beep
        MsgBox "Cannot send Mail from this System.", vbCritical, "Mail disabled/not installed"
    End If

End Sub

':) Ulli's VB Code Formatter V2.10.7 (24.02.2002 22:14:09) 30 + 19 = 49 Lines
