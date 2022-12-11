VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dlgRun
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run File"
   ClientHeight    =   1245
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1
      Left            =   120
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".exe .com .bat"
      DialogTitle     =   "Browse for File"
      Filter          =   "*.exe *.com *.bat"
   End
   Begin VB.CommandButton cmdBrowse
      Caption         =   "Browse..."
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "dlgRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmdBrowse_Click()
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
Select Case UCase$(Right$(CommonDialog1.FileName, 4))
Case ".EXE"
Text1.Text = CommonDialog1.FileName
Case ".BAT"
Text1.Text = CommonDialog1.FileName
Case ".COM"
Text1.Text = CommonDialog1.FileName
Case Else
CommonDialog1.FileName = "Start " + CommonDialog1.FileName
Text1.Text = CommonDialog1.FileName
End Select
End Sub

Private Sub OKButton_Click()
Shell Text1.Text, vbNormalFocus
Unload Me
End Sub
