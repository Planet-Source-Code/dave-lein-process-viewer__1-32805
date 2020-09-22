Attribute VB_Name = "modMain"
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Public Const MAX_PATH& = 260
Public Const dashes = " ----- "

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

Public Sub FormDrag(frmFormToDrag As Form)
    ReleaseCapture
    Call SendMessage(frmFormToDrag.hwnd, &HA1, 2, 0&)
End Sub

Public Sub Pause(PauseTime As Single)
    Dim StartTime As Single
    StartTime = Timer
    
    Do While Timer < PauseTime + StartTime
        DoEvents        'Dont hog processor
    Loop
End Sub

Public Sub TerminateTask(app_name As String)
    Target = app_name
    EnumWindows AddressOf EnumCallback, 0
End Sub

Public Sub WindowHandle(hWindow, mCase As Long)
Select Case mCase
    Case 0
        X = SendMessage(hWindow, WM_CLOSE, 0, 0)
    Case 1
        X = ShowWindow(hWindow, SW_SHOW)
    Case 2
        X = ShowWindow(hWindow, SW_HIDE)
    Case 3
        X = ShowWindow(hWindow, SW_Maximize)
    Case 4
        X = ShowWindow(hWindow, SW_Minimize)
    Case 5
        X = ShowWindow(hWindow, SW_Normal)
End Select
End Sub

Public Function EnumCallback(ByVal app_hWnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim title As String
Dim length As Long

    length = GetWindowText(app_hWnd, buf, Len(buf))
    title = Left$(buf, length)

    If InStr(title, Target) <> 0 Then
        SendMessage app_hWnd, WM_CLOSE, 0, 0&
    End If
    
    EnumCallback = 1
End Function

Public Function FileExists(FileName As String) As Boolean
'vbAlias = 64 (&H40) , vbArchive = 32 (&H20) , vbDirectory = 16 (&H16)
'vbVolume = 8 , vbSystem , vbHidden = 2 , vbReadOnly = 1 , vbNormal = 0

 Dim TempAttr As Integer
 
 On Error GoTo ErrorFileExist

 TempAttr = GetAttr(FileName)    'If can't get attribute it will error
 FileExists = ((TempAttr And vbDirectory) = 0)  'Return True essentially
 GoTo ExitFileExist
   
ErrorFileExist:                  'If no attributes come here
 FileExists = False              'Return False
 Resume ExitFileExist
   
ExitFileExist:
 On Error GoTo 0

End Function

Public Function GetComputerName() As String
    GetComputerName = GetStringValue("HKEY_LOCAL_MACHINE\SYSTEM\CURRENTCONTROLSET\CONTROL\COMPUTERNAME\COMPUTERNAME", "COMPUTERNAME")
End Function

Public Function GetProcessorID() As String
    GetProcessorID = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "Identifier")
End Function

Public Function GetProcessorName(ProcessorType As String)
    If ProcessorType = "AMD" Then
        GetProcessorName = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString")  'AMD
    Else
        GetProcessorName = GetStringValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "MMXIdentifier")        'Intel
    End If
End Function

Public Function GetWindowTitle(ByVal hwnd As Long) As String
On Error Resume Next
Dim S As String

L = GetWindowTextLength(hwnd)
S = Space(L + 1)

GetWindowText hwnd, S, L + 1
GetWindowTitle = Left$(S, L)
End Function

