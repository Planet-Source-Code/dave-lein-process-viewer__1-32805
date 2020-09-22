VERSION 5.00
Begin VB.Form frmProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "This is whats running on your computer"
   ClientHeight    =   6990
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5940
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh All Lists"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   1425
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   4680
      Width           =   5295
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FFFF&
      Height          =   1425
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox txtProcessTitle 
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FF00&
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label lblVisible 
      AutoSize        =   -1  'True
      Caption         =   "Visible Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   1725
   End
   Begin VB.Label lblProcesses 
      AutoSize        =   -1  'True
      Caption         =   "Processes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lblHidden 
      AutoSize        =   -1  'True
      Caption         =   "Hidden Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   4
      Top             =   4320
      Width           =   1755
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sorry this isn't commented. It was part of another program
'I wrote back in the day and I never got around to commenting it.
'Feel free to use this as u wish. It will be implemented with
'My remote client havoc soon :)

Public Function KillApp(myName As String) As Boolean
On Error GoTo errorhandler

GoSub Start

errorhandler:
    Exit Function

Start:

Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim I As Integer

Const PROCESS_ALL_ACCESS = 0
Const TH32CS_SNAPPROCESS As Long = 2&

appCount = 0

uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)

List1.Clear

Do While rProcessFound
    I = InStr(1, uProcess.szexeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szexeFile, I - 1))
    
    List1.AddItem (szExename)

    If Right$(szExename, Len(myName)) = LCase$(myName) Then
        KillApp = True
        appCount = appCount + 1
        myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        AppKill = TerminateProcess(myProcess, exitCode)
        Call CloseHandle(myProcess)
    End If

    rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop

End Function

Private Sub cmdRefresh_Click()
    List1.Clear: List2.Clear: List3.Clear
    KillApp ("")
    Call RefreshList
End Sub

Private Sub Form_Load()
    KillApp ("")
    RefreshList
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        FormDrag Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmProcess
End Sub

Private Sub List1_Click()
    txtProcessTitle = List1.Text
End Sub

Private Sub RefreshList()
For I = 1 To 10000
    A$ = GetWindowTitle(I)
    z = FindWindow(vbNullString, A$)
    hW = frmProcess.hwnd
    If z <> 0 Then
        If A$ <> vbNullString And LCase(A$) <> LCase(APPCap) And LCase(A$) <> "This is whats running on your computer right now" And I <> hW Then
            If IsWindowEnabled(z) = 0 Then
                If IsWindowVisible(z) = 0 Then
                    List3.AddItem "[Froze] " + A$
                ElseIf IsWindowVisible(z) = 1 Then
                    List2.AddItem "[Froze] " + A$
                End If
            ElseIf IsWindowEnabled(z) = 1 Then
                If IsWindowVisible(z) = 0 Then
                    List3.AddItem A$
                ElseIf IsWindowVisible(z) = 1 Then
                    List2.AddItem A$
                End If
            End If
        End If
    End If
Next I
End Sub

Private Sub List1_DblClick()
    MsgBox List1.List(List1.ListIndex)
End Sub

Private Sub List2_DblClick()
    MsgBox List2.List(List2.ListIndex)
End Sub

Private Sub List3_DblClick()
    MsgBox List3.List(List3.ListIndex)
End Sub

Private Sub mnuExit_Click()
    Unload frmProcess
End Sub
