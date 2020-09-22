Attribute VB_Name = "modCustomMsgBox"
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long



Public strFileToPlay As String
Public bPlaying As Boolean
'=========================================================================================
'  modCustomMsgBox
'  routines and functions to allow customization to a msgbox
'=========================================================================================
'  Created By:Amer
'  Published Date: 02/23/2001
'  Legal Copyright: Amer © 02/23/2001
'=========================================================================================
Option Explicit
Global MsgBoxIcon As Integer
Private Const WH_CBT As Long = &H5
Private Const HCBT_ACTIVATE As Long = &H5
Private Const STM_SETICON As Long = &H170
Private Const MODAL_WINDOW_CLASSNAME As String = "#32770"
Private Const SS_ICON As Long = &H3
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const STM_SETIMAGE As Long = &H172
Private Const IMAGE_CURSOR As Long = &H2

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long

Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Type ANICURSOR
   m_hCursor As Long
   m_hWnd As Long
End Type

Private pHook As Long
Private phIcon As Long
Private pAniIcon As String

Dim MsgBoxTitle As String
Dim MsgBoxText As String
Dim MsgBoxX As Long
Dim MsgBoxY As Long
Dim MsgBoxTimeOut As Integer
Dim MsgBoxHWND As Long
Dim MsgBoxButtons As Long
Dim MsgBoxCheckBoxText As String
Dim MsgBoxCheckBoxHWND As Long
Dim MsgBoxCheckBoxValue As Boolean
Dim MsgBoxButtonCount As Integer
Dim MsgBoxButtonText() As String
Dim MsgBoxOriginalButtonText() As String

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const MB_CLOSEMSGBOX = &H5000&
Private Const MB_MOVEMSGBOX = &H5001&
Private Const MB_CHECK_CHECKBOXVALUE = &H5003&
Private Const MB_CHANGEBUTTONTEXT = &H5002&
Private Const MB_ADDCHECKBOX = &H5004&


Private Const WS_TABSTOP = &H10000
Private Const WM_SETFONT = &H30
Private Const BS_AUTOCHECKBOX = &H3&
Private Const HWND_TOPMOST = -1
Private Const BM_GETSTATE = &HF2
Private Const WM_GETFONT = &H31

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal bRepaint&) As Long
Private Declare Function GetFocus Lib "user32" () As Long
'=========================================================================================
Public Function CreateSpecialMsgbox(Text As String, Buttons As Long, Title As String, Optional ByVal TimeOut As Integer = -1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, Optional CheckBoxText As String = "", Optional CheckBoxValue As Boolean, Optional ByVal ButtonCount As Integer, Optional ButtonText As String, Optional OriginalButtonText As String, Optional ByVal hIcon As Long = 0&, Optional ByVal AniIcon As String = "", Optional SoundFile As String) As VbMsgBoxResult
' routine to create the special msgbox based on passed parameters
  On Error Resume Next
  MsgBoxHWND = Screen.ActiveForm.hwnd
  MsgBoxText = Text
  MsgBoxTitle = Title
  MsgBoxButtons = Buttons
  MsgBoxTimeOut = TimeOut * 1000
  MsgBoxX = x
  MsgBoxY = y
  MsgBoxCheckBoxText = CheckBoxText
  MsgBoxButtonCount = ButtonCount
  MsgBoxButtonText = Split(ButtonText, ",")
  MsgBoxOriginalButtonText = Split(OriginalButtonText, ",")
 If SoundFile <> "" Then
 If Right$(SoundFile, 4) = ".mp3" Then
 
playMp SoundFile
Else
 PlaySound SoundFile, ByVal 0&, SND_FILENAME Or SND_ASYNC
End If

End If
'------Additional Code For CustomIcon And Animation--------------------
 '_____________________________________________________________________
   pHook = SetWindowsHookEx(WH_CBT, _
          AddressOf MsgBoxHookProc, _
                     App.hInstance, _
                 GetCurrentThreadId())
                 
   ' Save other arguments for use in MsgBoxHookProc
   phIcon = hIcon
   pAniIcon = AniIcon
   
   ' If a custom icon (animated or otherwise) is required
   ' make sure the msgbox makes room for it by setting the
   ' style to vbInformation; ensure other icon styles are set
   ' off, as if more than one are set no icon can be displayed.
   If Len(AniIcon) <> 0 Or phIcon <> 0 Then
      MsgBoxIcon = MsgBoxIcon And Not (vbCritical)
      MsgBoxIcon = MsgBoxIcon And Not (vbExclamation)
      MsgBoxIcon = MsgBoxIcon And Not (vbQuestion)
      MsgBoxIcon = MsgBoxIcon Or vbInformation
   End If
'_____________________________________________________________________


  'set the timer to fire for various effects
  If MsgBoxX <> -1 And MsgBoxY <> -1 Then
    SetTimer MsgBoxHWND, MB_MOVEMSGBOX, 0&, AddressOf NewTimerProc
  End If
  If MsgBoxTimeOut > 0 Then
    SetTimer MsgBoxHWND, MB_CLOSEMSGBOX, MsgBoxTimeOut, AddressOf NewTimerProc
  End If
  If ButtonText <> "" And OriginalButtonText <> "" Then
    SetTimer MsgBoxHWND, MB_CHANGEBUTTONTEXT, 0&, AddressOf NewTimerProc
  End If
  If MsgBoxCheckBoxText <> "" Then
    SetTimer MsgBoxHWND, MB_CHECK_CHECKBOXVALUE, 0&, AddressOf NewTimerProc
    SetTimer MsgBoxHWND, MB_ADDCHECKBOX, 0&, AddressOf NewTimerProc
  End If

  ' creat the special Msgbox
  CreateSpecialMsgbox = MessageBox(MsgBoxHWND, MsgBoxText, MsgBoxTitle, MsgBoxButtons)

  If MsgBoxCheckBoxText <> "" Then
    'cancel the checkbox fire timer
    MsgBoxCheckBoxHWND = 0
    KillTimer MsgBoxHWND, MB_CHECK_CHECKBOXVALUE
    'return its state
    CheckBoxValue = MsgBoxCheckBoxValue
  End If
  
End Function 'CreateSpecialMsgbox(Text As String, Buttons As Long, Title As String, Optional ByVal TimeOut As Integer = -1, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1, Optional ByVal CheckBoxText As String = "", Optional ByRef CheckBoxValue As Boolean, Optional ByVal ButtonCount As Integer, Optional ButtonText As String, Optional OriginalButtonText As String) As VbMsgBoxResult
'=========================================================================================
Public Function NewTimerProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' routine to fire for timer events which determine what should be done for the custom msgbox
  Dim TempForm As Form
  On Error Resume Next
  If wParam = MB_CHECK_CHECKBOXVALUE Then
    If MsgBoxCheckBoxHWND > 0 Then
      'Returns the value of the checkbox on extended MsgBox
      MsgBoxCheckBoxValue = (SendMessage(MsgBoxCheckBoxHWND, BM_GETSTATE, 0, 0&) <> 0)
    Else
      KillTimer hwnd, wParam
    End If
  Else
    'Cancel timer
    KillTimer hwnd, wParam
    MsgBoxHWND = FindWindow("#32770", MsgBoxTitle)
    If MsgBoxHWND <> 0 Then
      Select Case wParam
        Case MB_CLOSEMSGBOX
          SetForegroundWindow MsgBoxHWND
          If MsgBoxCheckBoxText <> "" Then
            If MsgBoxCheckBoxHWND = GetFocus&() Then
              If MsgBoxButtonCount = 1 Then SendKeys "{TAB}"
              If MsgBoxButtonCount > 1 Then SetForegroundWindow MsgBoxHWND
            End If
          End If
          SendKeys "{enter}"
        Case MB_MOVEMSGBOX
          Dim W As Single
          Dim H As Single
          Dim rBox As RECT
          W = Screen.Width / Screen.TwipsPerPixelX
          H = Screen.Height / Screen.TwipsPerPixelY
          GetWindowRect MsgBoxHWND, rBox
          With rBox
            If MsgBoxX > (W - (.Right - .Left) - 1) Then MsgBoxX = (W - (.Right - .Left) - 1)
            If MsgBoxY > (H - (.Bottom - .Top) - 1) Then MsgBoxY = (H - (.Bottom - .Top) - 1)
          End With
          If MsgBoxX < 1 Then MsgBoxX = 1: If MsgBoxY < 1 Then MsgBoxY = 1
          SetWindowPos MsgBoxHWND, HWND_TOPMOST, MsgBoxX, MsgBoxY, 0, 0, SWP_NOSIZE
        Case MB_ADDCHECKBOX
          Dim hFont As Long
          Dim lCaptionHwnd As Long
          Dim R As RECT
          Dim nHeight As Integer
          'Find the window
          lCaptionHwnd = FindWindowEx(MsgBoxHWND, 0, "Static", MsgBoxText)
          GetWindowRect MsgBoxHWND, R
          Set TempForm = Screen.ActiveForm
          nHeight = TempForm.TextHeight(MsgBoxCheckBoxText) / Screen.TwipsPerPixelY
          'Create the checkbox control
          MsgBoxCheckBoxHWND = CreateWindowEx(0, "Button", MsgBoxCheckBoxText, WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_AUTOCHECKBOX, 3, (R.Bottom - R.Top) - nHeight - 15, (TempForm.TextWidth(MsgBoxCheckBoxText) / Screen.TwipsPerPixelX) + 22, nHeight, MsgBoxHWND, 0, App.hInstance, ByVal 0&)
          ' set the font of the checkbox to the same as the messagebox
          hFont = SendMessage(lCaptionHwnd, WM_GETFONT, 0, 0&)
          SendMessage MsgBoxCheckBoxHWND, WM_SETFONT, hFont, 0&
          'move the new checkbox to the correct position
          MoveWindow MsgBoxHWND, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top + nHeight, 1&
        Case MB_CHANGEBUTTONTEXT
          Dim hButton As Long
          Dim Counter As Integer
          'replace the captions where required
          For Counter = LBound(MsgBoxOriginalButtonText) To UBound(MsgBoxOriginalButtonText)
            If Len(MsgBoxButtonText(Counter)) > 0 Then
              hButton = FindWindowEx(MsgBoxHWND, 0&, "Button", MsgBoxOriginalButtonText(Counter))
              If hButton <> 0 Then
                SetWindowText hButton, MsgBoxButtonText(Counter)
              End If
            End If
          Next Counter
      End Select
    End If
  End If
  'clear objects
  Set TempForm = Nothing
End Function 'NewTimerProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'=========================================================================================

Private Function MsgBoxHookProc(ByVal CodeNo As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) As Long
   Dim ClassNameSize As Long
   Dim sClassName As String
   Dim hIconWnd As Long
   Dim M As ANICURSOR

   ' Call the next hook; this is standard stuff.
   MsgBoxHookProc = CallNextHookEx(pHook, CodeNo, wParam, lParam)
   ' Only interfere if the msgbox activate message is being dealt with:
   If CodeNo = HCBT_ACTIVATE Then
      ' Check the classname; exit if not a standard msgbox.
      sClassName = Space$(32)
      ClassNameSize = GetClassName(wParam, sClassName, 32)
      If Left$(sClassName, ClassNameSize) <> MODAL_WINDOW_CLASSNAME Then Exit Function
   
      ' If displaying custom icon (animated or not), get icon window handle.
      If phIcon <> 0 Or Len(pAniIcon) <> 0 Then _
         hIconWnd = FindWindowEx(wParam, 0&, "Static", vbNullString)
      
      ' If custom (non- animated) icon, set here:
      If phIcon <> 0 Then SendMessage hIconWnd, STM_SETICON, phIcon, ByVal 0&
      
      ' If custom (animated) icon, set here: (animated takes precidence)
      If Len(pAniIcon) Then AniCreate M, pAniIcon, hIconWnd, 0, 0
      
      'unhook.
      UnhookWindowsHookEx pHook
   End If
End Function

Public Sub AniCreate(ByRef m_AniStuff As ANICURSOR, sAniName As String, hwndParent As Long, x As Long, y As Long)
   ' Creates an animated cursor on hwndParent at x,y
   
   ' First destroy previous ani if m_AniStuff refers to one.
   AniDestroy m_AniStuff
   With m_AniStuff
      ' Get cursor.
      .m_hCursor = LoadCursorFromFile(sAniName)
      If .m_hCursor Then
         ' Create cursor window.
         .m_hWnd = CreateWindowEx(0, "Static", "", WS_CHILD Or WS_VISIBLE Or SS_ICON, ByVal 20, ByVal 20, 0, 0, hwndParent, 0, App.hInstance, ByVal 0)
         If .m_hWnd Then
            ' Place cursor in window & position
            SendMessage .m_hWnd, STM_SETIMAGE, IMAGE_CURSOR, ByVal .m_hCursor
            SetWindowPos .m_hWnd, 0, x, y, 0, 0, SWP_NOZORDER Or SWP_NOSIZE
         Else
            ' Clean up.
            DestroyCursor .m_hCursor
         End If
      End If
   End With
End Sub

Public Sub AniDestroy(ByRef m_AniStuff As ANICURSOR)
   ' Destroy animated cursor referenced by m_AniStuff
   With m_AniStuff
      If .m_hCursor Then _
         If DestroyCursor(.m_hCursor) Then .m_hCursor = 0
      If IsWindow(.m_hWnd) Then _
         If DestroyWindow(.m_hWnd) Then .m_hWnd = 0
   End With
End Sub

Sub playMp(File As String)
CloseAll

strFileToPlay = File
'"C:\My Documents\Custom MessageBox\Custom MessageBox\Sound\m_128kbps.mp3"
 'comDlg.FileName
        strFileToPlay = """" & strFileToPlay & """"
        Call OpenMovie
        Call PlayMovie
End Sub


Public Sub PlayMovie()
    If strFileToPlay <> "" Then
        mciSendString "play " & strFileToPlay, 0, 0, 0
        bPlaying = True
        'frmMain.lblCaption.Caption = "[ Playing ]"
        
    End If
End Sub

Public Sub StopMovie()
    If bPlaying Then
        mciSendString "stop " & strFileToPlay, 0, 0, 0
        bPlaying = False
      '  frmMain.lblCaption.Caption = "[ Stoped ]"
    End If
End Sub

Public Sub CloseMovie()
    If bPlaying Then
        mciSendString "close " & strFileToPlay, 0, 0, 0
        bPlaying = False
        'frmMain.lblCaption.Caption = "[ No Media ]"
       ' UpdateScreen
    End If
End Sub

Public Sub CloseAll()
    mciSendString "close all", 0, 0, 0
End Sub

Public Sub OpenMovie()
    If strFileToPlay <> "" Then
        mciSendString "open " & strFileToPlay & " type MPEGVideo", 0, 0, 0
    End If
End Sub
