Attribute VB_Name = "modData"
'****************************************************************************************************************************
'These are the Client Properties!!! Add anything if you need to.
'****************************************************************************************************************************
Type Client
    Name As String
    Login As Boolean
    Action As String
    LastCommand As String
    NextCommand As String
    Attempts As Integer
    iCount As Long
    rREC As String
End Type

Public Clients() As Client
'****************************************************************************************************************************

'****************************************************************************************************************************
' Various declarations
'****************************************************************************************************************************

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long


Public programVer As String
Public prgVersion As String
Type MYVERSION
    lMajorVersion As Long
    lMinorVersion As Long
    lExtraInfo As Long
End Type

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Global Const VER_PLATFORM_WIN32s = 0
Global Const VER_PLATFORM_WIN32_WINDOWS = 1
Global Const VER_PLATFORM_WIN32_NT = 2

Public BANSERVER As Boolean





'**************************************************************'
'Windows version :-p
'**************************************************************'
Function WindowsVersion() As MYVERSION
Dim myOS As OSVERSIONINFO, WinVer As MYVERSION
Dim lResult As Long

    myOS.dwOSVersionInfoSize = Len(myOS)    'should be 148
    
    lResult = GetVersionEx(myOS)
        
    'Fill user type with pertinent info
    WinVer.lMajorVersion = myOS.dwMajorVersion
    WinVer.lMinorVersion = myOS.dwMinorVersion
    WinVer.lExtraInfo = myOS.dwPlatformId
    
    WindowsVersion = WinVer

End Function

'**************************************************************'
'This function hides the program from the CTRL + ALT + DEL list
'Only works on Win98/95 and NOT in NT.
'**************************************************************'
Public Sub MakeMeService()
On Error Resume Next
Dim pid As Long
Dim reserv As Long

pid = GetCurrentProcessId()
regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub



'**************************************************************
'This function shows again the program to CTRL+ALT+DEL list.
'**************************************************************'
Public Sub UnMakeMeService()
On Error Resume Next
Dim pid As Long
Dim reserv As Long

pid = GetCurrentProcessId()
regserv = RegisterServiceProcess(pid, _
RSP_UNREGISTER_SERVICE)
End Sub


'**************************************************************
'Centers form to the screen.
'**************************************************************'
Sub CenterForm(Frm As Form)
Frm.Left = (Screen.Width / 2) - (Frm.Width / 2)
Frm.Top = (Screen.Height / 2) - (Frm.Height / 2)
End Sub


