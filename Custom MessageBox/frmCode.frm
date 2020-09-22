VERSION 5.00
Begin VB.Form frmCode 
   Caption         =   "Genrated MessageBox Syntex Code"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2640
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtModule 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmCode.frx":000C
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim CheckBoxValue As Boolean
CreateSpecialMsgbox "Message Box Body Text", 546, "Message Box Title", -1, -1, -1, "CheckBox Text", CheckBoxValue, 3, "&Abort,&Retry,&Ignore", "&Abort,&Retry,&Ignore"
End Sub

'=========================================================================================
Private Sub Form_Load()
' routine to resize textbox and set form on top
  On Error Resume Next
  txtCode.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
End Sub 'Form_Load()
'=========================================================================================
Private Sub Form_Resize()
' routine to resize textbox
  On Error Resume Next
  txtCode.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
End Sub 'Form_Resize()
'=========================================================================================
