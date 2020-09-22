VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFileToPlay As String
Private Sub Command1_Click()
strFileToPlay = App.Path & "\Sound\m_128kbps.mp3"
Debug.Print strFileToPlay
        strFileToPlay = """" & strFileToPlay & """"

        
        Call OpenMovie
        Call PlayMovie

  'strFileToPlay = """" & strFileToPlay & """"
        ' App.Path & "\Sound\m_128kbps.mp3"
End Sub

Sub playmp3(strFileToPlay As String)
        mciSendString "play " & strFileToPlay, 0, 0, 0

End Sub

Public Sub OpenMP3(SoundFile As String)
    If SoundFile <> "" Then
        mciSendString "open " & SoundFile & " type MPEGVideo", 0, 0, 0
    End If
End Sub
Public Sub OpenMovie()
Debug.Print strFileToPlay
    If strFileToPlay <> "" Then
        mciSendString "open" & strFileToPlay & " type MPEGVideo", 0, 0, 0
    End If
End Sub
Public Sub PlayMovie()
Debug.Print strFileToPlay
    If strFileToPlay <> "" Then
        mciSendString "play " & strFileToPlay, 0, 0, 0
        'bPlaying = True
       ' frmMain.lblCaption.Caption = "[ Playing ]"
        
    End If
End Sub
