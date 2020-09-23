Attribute VB_Name = "modApi32"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : modApi32
'    Project    : CSScrollBox
'    Created By : Project Administrator
'    Description: Api definitions
'
'    Modified   : 22/4/2004 18:48:58
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

' Used to force window on top.
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
' Set TopMost Constants
Private Const SWP_NOOWNERZORDER = &H200              ' Don"t do owner Z ordering
Private Const SWP_FRAMECHANGED = &H20                ' The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

'Sound and media functions
Private Declare Function mciGetErrorString Lib "winmm" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function PlaySound Lib "winmm" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySoundAsBytes Lib "winmm.dll" Alias "sndPlaySoundA" (ab As Any, ByVal dwFlags As Long) As Boolean

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' PlaySound enumeration
Public Enum PlaySoundFlags
    SND_ASYNC = &H1
    SND_LOOP = &H8
    SND_NODEFAULT = &H2
    SND_NOSTOP = &H10
    SND_NOWAIT = &H2000
    SND_PURGE = &H40
    SND_SYNC = &H0
    SND_ALIAS = &H10000
    SND_FILENAME = &H20000
    SND_MEMORY = &H4
End Enum

Public bytChunk() As Byte ' was public before marclei???

Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Public Function LoadPictureBytes(b() As Byte) As IPicture
    Dim LowerBound As Long
    Dim ByteCount  As Long
    Dim hMem  As Long
    Dim lpMem  As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.IUnknown

    On Error GoTo Err_Init
    If UBound(b, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                    Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), LoadPictureBytes)
                End If
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:
    If Err.Number = 9 Then
        'Uninitialized array
        MsgBox "You must pass a non-empty byte array to this function!"
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If

End Function

Public Sub DoEvents2()
    'Sleep 1 is to prevent 100% CPU Usage
    Sleep 1
    DoEvents
End Sub

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    Do Until GetTickCount > EndTime
        DoEvents2
    Loop
End Function

Public Sub StayOnTop(hWndA As Long, Optional Enable As Boolean)
    Dim Flags As Long
    Flags = IIf(Enable, HWND_TOPMOST, HWND_NOTOPMOST)
    ' set the windows top most
    SetWindowPos hWndA, Flags, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

' Play sounds and music functions
Public Function PlayShortWav(ByVal sndName As String) As Long
    TerminateAllSound
    PlayShortWav = PlaySound(sndName, 0, PlaySoundFlags.SND_NODEFAULT Or PlaySoundFlags.SND_SYNC)
End Function

Public Function PlayWav(ByVal StrName As String) As Long
    Dim RetStr As String
    Dim CallBack As Long
    Dim ShortName As String
    
    TerminateAllSound
    
    RetStr = Space$(128)
    ShortName = GetShortFileName(StrName)
    PlayWav = mciSendString("open waveaudio!" & ShortName & " alias wav", RetStr, 128, CallBack)
    PlayWav = mciSendString("play wav", RetStr, 128, CallBack)
End Function

Public Function PlayMidi(ByVal StrName As String) As Long
    TerminateAllSound

    Dim RetStr As String, CallBack As Long, ShortName
    
    RetStr = Space$(128)
    ShortName = GetShortFileName(StrName)
    PlayMidi = mciSendString("open sequencer!" & ShortName & " alias midi", RetStr, 128, CallBack)
    PlayMidi = mciSendString("play midi", RetStr, 128, 1)
End Function

'Stop sounds and music functions
Public Function StopWav() As Long
    Dim RetStr As String, CallBack As Long
    
    RetStr = Space$(128)
    StopWav = mciSendString("stop wav", RetStr, 128, CallBack)
    StopWav = mciSendString("close wav", RetStr, 128, CallBack)
End Function

Public Function StopMidi() As Long
    Dim RetStr As String, CallBack As Long
    
    RetStr = Space$(128)
    StopMidi = mciSendString("stop midi", RetStr, 128, CallBack)
    StopMidi = mciSendString("close midi", RetStr, 128, CallBack)

End Function

Public Function TerminateAllSound() As Long
    StopWav
    StopMidi
End Function

' Function to get the short hand file name for use with mci strings
Private Function GetShortFileName(ByVal FileName As String) As String
    Dim rc As Long
    Dim ShortPath As String
    
    Const PATH_LEN& = 164
    
    ShortPath = String$(PATH_LEN + 1, 0)
    rc = GetShortPathName(FileName, ShortPath, PATH_LEN)
    GetShortFileName = Left$(ShortPath, rc)
End Function

'This I had help on from the Programming Author
'Denis Wiegand the maker of the code "Waves"
'This is basicly just a little more advanced version of it
'I haven't tested it yet but it should work
Sub StopWaveBytes()
    Dim X As Integer
    
    X% = sndPlaySound("", SND_ASYNC + SND_NODEFAULT)
    Erase bytChunk
End Sub

'-----------------------------------------------------------------
' WARNING:  If you want to play sound files asynchronously in
'           Win32, then you MUST change bytSound() from a local
'           variable to a module-level or static variable. Doing
'           this prevents your array from being destroyed before
'           PlaySoundFlags is complete. If you fail to do this, you
'           will pass an invalid memory pointer, which will cause
'           a GPF in the Multimedia Control Interface (MCI).
'-----------------------------------------------------------------
Public Sub PlayWaveBytes(ab() As Byte, Optional vntFlags As PlaySoundFlags)
    bytChunk = ab
    If IsMissing(vntFlags) Or vntFlags = 0 Then
        vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
    End If
    If (vntFlags And SND_MEMORY) = 0 Then
        vntFlags = vntFlags Or SND_MEMORY
    End If
    Dim s As String
    
    sndPlaySoundAsBytes bytChunk(0), CLng(vntFlags)
End Sub
'--end code
