VERSION 5.00
Begin VB.Form frmRfl
   BackColor       =   &H0000FF00&
   Caption         =   "Red Fire Light"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   FillColor       =   &H0000FF00&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1
      Interval        =   1000
      Left            =   4200
      Top             =   2760
   End
   Begin VB.CommandButton cmdCommand
      Caption         =   "Command Prompt"
      BeginProperty Font
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdRun
      BackColor       =   &H000080FF&
      Caption         =   "Run"
      BeginProperty Font
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaskColor       =   &H000080FF&
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblDate
      BackColor       =   &H0000FF00&
      Caption         =   "Date"
      BeginProperty Font
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblTime
      BackColor       =   &H0000FF00&
      Caption         =   "Time"
      BeginProperty Font
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Image recycle
      Height          =   720
      Left            =   1680
      Picture         =   "rfl.frx":0000
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image mydocs
      Height          =   720
      Left            =   840
      Picture         =   "rfl.frx":068A
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image mycomp
      Height          =   720
      Left            =   120
      Picture         =   "rfl.frx":1554
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label1
      BackColor       =   &H0000FF00&
      Caption         =   "Welcome to Red Fire Light!"
      BeginProperty Font
         Name            =   "Old English Text MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu file
      Caption         =   "&File"
      Begin VB.Menu exit
         Caption         =   "Exit"
      End
      Begin VB.Menu popup
         Caption         =   "PopUp"
         Visible         =   0   'False
         Begin VB.Menu apps
            Caption         =   "Applications"
            Begin VB.Menu paint
               Caption         =   "Paint"
            End
            Begin VB.Menu agent
               Caption         =   "Agent stuff"
            End
            Begin VB.Menu xearth
               Caption         =   "Xearth"
            End
            Begin VB.Menu acc
               Caption         =   "Accessories"
            End
            Begin VB.Menu office
               Caption         =   "Office Stuff"
            End
            Begin VB.Menu gamesfactory
               Caption         =   "The Games Factory"
            End
            Begin VB.Menu iconforge
               Caption         =   "IconForge"
            End
         End
         Begin VB.Menu games
            Caption         =   "Games"
            Begin VB.Menu dragon
               Caption         =   "Dragon game"
            End
         End
         Begin VB.Menu rflstuff
            Caption         =   "Red Fire Light Stuff"
         End
         Begin VB.Menu startmenu
            Caption         =   "Your Start Menu"
         End
         Begin VB.Menu help
            Caption         =   "Help"
         End
         Begin VB.Menu intranet
            Caption         =   "Intranet"
         End
         Begin VB.Menu restart
            Caption         =   "Restart"
         End
         Begin VB.Menu shutdown
            Caption         =   "Shut down"
         End
      End
   End
End
Attribute VB_Name = "frmRfl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acc_Click()
Shell "Explorer " & Environ("UserProfile") & "\start menu\programs\accessories", vbMaximizedFocus
End Sub

Private Sub agent_Click()
Shell "Explorer " & Environ("winbootdir") & "\msagent\chars", vbMaximizedFocus
End Sub

Private Sub cmdCommand_Click()
Shell Environ("comspec"), vbNormalFocus
End Sub

Private Sub cmdRun_Click()
dlgRun.Show
End Sub

Private Sub dragon_Click()
frmDragon.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu popup
End If
End Sub

Private Sub gamesfactory_Click()
Shell "c:\gfactory\gfact32.exe", vbMaximizedFocus

End Sub

Private Sub help_Click()
frmAbout.Show

End Sub

Private Sub iconforge_Click()
Shell "c:\program files\iconforge\iconforge.exe", vbMaximizedFocus

End Sub

Private Sub intranet_Click()
Shell "C:\Program Files\Internet Explorer\iexplore.exe http://bozo/", vbMaximizedFocus

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Me.PopupMenu popup
End If

End Sub

Private Sub mycomp_DblClick()
Shell "Explorer my computer", vbMaximizedFocus
End Sub

Private Sub mydocs_DblClick()
Shell "explorer 'my documents'", vbMaximizedFocus

End Sub

Private Sub office_Click()
Shell "Explorer " & Environ("UserProfile") & "\Start Menu\Programs\Microsoft Office", vbNormalFocus

End Sub

Private Sub paint_Click()
Shell "" & Environ("winbootdir") & "\system32\Mspaint.exe", vbMaximizedFocus

End Sub

Private Sub recycle_DblClick()
Shell "Explorer c:\recycled", vbMaximizedFocus

End Sub

Private Sub restart_Click()
Shell "" & Environ("winbootdir") & "\RUNDLL.EXE user.exe,exitwindowsexec"
End Sub

Private Sub rflstuff_Click()
Shell "Explorer .", vbMaximizedFocus

End Sub

Private Sub shutdown_Click()
Shell "" & Environ("winbootdir") & "\RUNDLL32.EXE User,ExitWindows"
End Sub

Private Sub startmenu_Click()
Shell "explorer " & Environ("UserProfile") & "\start menu", vbMaximizedFocus

End Sub

Private Sub Timer1_Timer()
lblTime = Time
lblDate = Date
End Sub

Private Sub xearth_Click()
Shell "xearth\xearth.exe"
End Sub
