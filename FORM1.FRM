VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form1 
   Caption         =   "Internet Explorer Text reader ve3ded@hotmail.com"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   6180
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6405
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   9390
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3360
      Top             =   600
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5640
      Top             =   720
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu stop 
      Caption         =   "S&top"
   End
   Begin VB.Menu read 
      Caption         =   "R&ead"
   End
   Begin VB.Menu exit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu save 
      Caption         =   "S&ave"
   End
   Begin VB.Menu clip 
      Caption         =   "C&lipboard read"
   End
   Begin VB.Menu font 
      Caption         =   "Fo&nt Size"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim merlin As IAgentCtlCharacterEx
Const DATAPATH = "merlin.acs"
Private Sub clip_Click()
Text1.Text = Clipboard.GetText
merlin.Show
merlin.Speak Text1.Text + " "
merlin.Speak "I have read the contents of the clipboard"
merlin.Hide
Text1.Text = " "
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub font_Click()
Dim f
merlin.Show
bb:
merlin.Speak "type a number from 8 to 72"
f = InputBox("Type a number from 8 to 72", "Type a number from 8 to 72")
If f < 8 Or f > 72 Then GoTo bb
Text1.FontSize = f
merlin.Hide
End Sub

Private Sub Form_Load()
Agent1.Characters.Load "merlin.acs", DATAPATH
    Set merlin = Agent1.Characters("merlin.acs")
    merlin.LanguageID = &H409
       merlin.MoveTo 200, 355

End Sub

Private Sub read_Click()
Dim i$
merlin.Show
merlin.Speak "Hit enter to  continue"
i$ = InputBox("Hit enter to continue", "Hit enter to coninue")
merlin.Show
merlin.Speak Text1.Text + " "
merlin.Speak "I have read the entire page"
merlin.Hide
End Sub

Private Sub save_Click()
Dim s$
merlin.Show
merlin.Speak "What is the name of the text file you wish to save"
s$ = InputBox("what is the name of the text file you wish to save", "what is the name of the text file you wish to save")
Open s$ For Output As #1
Print #1, Text1.Text
Close #1
merlin.Speak "File is saved"
merlin.Hide
End Sub

Private Sub stop_Click()
merlin.stop
End Sub
Private Sub Timer1_Timer()
GetExplorerText
End Sub
