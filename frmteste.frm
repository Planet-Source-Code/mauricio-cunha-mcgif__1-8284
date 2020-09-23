VERSION 5.00
Object = "{65E3E3C3-30B6-11D4-AABA-0004ACBF1E11}#1.0#0"; "mcgif.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTeste 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test of MC Gif Control"
   ClientHeight    =   2805
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin prjmcgif.mcgif mcgif1 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4683
      BorderStyle     =   1
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDial 
      Left            =   2880
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuBrowse 
         Caption         =   "&Browse..."
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "FrmTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdPlay_Click()
mcgif1.Play
CmdStop.Enabled = True
CmdPlay.Enabled = False

End Sub

Private Sub CmdStop_Click()
mcgif1.StopPlay
CmdStop.Enabled = False
CmdPlay.Enabled = True

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mcgif1_Click()
If CmdPlay.Enabled = False Then
 CmdStop_Click
 Exit Sub
Else
 CmdPlay_Click
 Exit Sub
End If

End Sub

Private Sub MnuAbout_Click()
mcgif1.ShowAboutBox
End Sub

Private Sub MnuBrowse_Click()
On Error Resume Next

CDial.Filter = "Animated Gif (*.gif)|*.gif"
CDial.FileName = ""
CDial.ShowOpen

If CDial.FileName <> "" Then
 mcgif1.FileName = CDial.FileName
 mcgif1.OpenWithOutDlg
End If

End Sub

Private Sub MnuExit_Click()
Unload Me
End Sub

