VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Type message and hit [ENTER]."
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   165
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3090
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "Sample Text...."
      Top             =   75
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sent As Boolean

Private Sub Form_Load()
    Form1.Show
    Winsock1.Connect "127.0.0.1", "21"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Winsock1.SendData Text1.Text: Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
    If Winsock1.State = 8 Then Me.Caption = "Disconnected": Text1.Enabled = False: Timer1.Enabled = False: Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    Me.Caption = "Connected"
    Text1.Enabled = True
    Timer1.Enabled = True
End Sub
