VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Add WAV sound to AVI video Example"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   ScaleHeight     =   1905
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   $"frmMain.frx":0000
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FileCopy App.Path & "\Source.avi", App.Path & "\Dest.avi"
Call AddAudioStream(App.Path & "\Dest.avi", App.Path & "\Source.wav") 'AddAudioStream(Destination AVI FilePath, Source WAV FilePath)
MsgBox "Play ""Dest.avi"" file, please."
'To use this function, need 2 files
'1. Existing avi video file(Without sound) -> ex) App.Path & "\Dest.avi"
'2. Existing wav audio file -> ex) App.Path & "\Source.wav"
End Sub
