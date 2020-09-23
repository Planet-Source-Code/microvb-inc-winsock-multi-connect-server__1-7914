VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   3075
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Initializing Ports 0 - 9999"
      Top             =   1260
      Width           =   2850
   End
   Begin MSWinsockLib.Winsock host 
      Index           =   0
      Left            =   45
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   21
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   1950
      TabIndex        =   4
      Top             =   1545
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   9999
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1350
      Top             =   45
   End
   Begin VB.ListBox List2 
      Height          =   840
      ItemData        =   "Form1.frx":0000
      Left            =   5745
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   165
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00004000&
      Height          =   2955
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   6510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4335
      TabIndex        =   2
      Top             =   3135
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   3015
      Left            =   60
      Top             =   60
      Width           =   6570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim databuffer As String
Dim connection As Long
Dim maxconnections As Long

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Form1.Show
    'Keep track of up to 1000 clients simultaneously connected.
    'Increase this number to add support for more connections.
    maxconnections = 1000
    
    'Port the server is listening/hosting on.
    host(0).LocalPort = 21
    
    'Enumerate unconnected clients according to 'maxconnections'
    Text1.Text = "Enumerating Sockets 1 to " & maxconnections
    ProgressBar1.Max = maxconnections
    For x = 0 To maxconnections: DoEvents
        List2.AddItem Format(x, "0000")
        ProgressBar1.Value = x:
    Next x
    ProgressBar1.Visible = False: Text1.Visible = False: List1.Visible = True
    host(0).Listen
    Label1.Caption = "Hosting on port: " & host(0).LocalPort
End Sub

Private Sub host_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    List2.Selected(0) = True
    connection = List2.Text
    connection = connection + 1
    If host(0).State = 2 Then
        Load host(connection)
        host(connection).Accept requestID
        host(connection).SendData "Connected!!"
        List2.RemoveItem (0)
        List1.AddItem "Connected: " & Format(connection, "0000")
        If Timer1.Enabled = False Then Timer1.Enabled = True
    End If

End Sub

Private Sub host_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strdata As String
    host(Index).GetData strdata
    'If the client sends a command in the format COMMAND]]> strip the command out
    z = InStr(strdata, "]]>")
    If z > 0 Then cmd = Mid(strdata, 1, z - 1)
    
    'If the command is DISCONNECT]]> re-add winsock host to available list thus
    'freeing up some memory, and some internet ports. I have not figured out how
    'to setup a disconnect method if the user does not send one without taking up
    'considerable resources.
    Select Case cmd
        Case "DISCONNECT"
            host(Index).Close
            If Index > 0 Then List2.AddItem Format(Index - 1, "0000"): Unload host(Index)
            List1.AddItem "Disconnected: " & Format(Index, "0000")
            Exit Sub
    End Select
    
    'Show recieved data in command window as well as the client that sent it
    List1.AddItem Format(Index, "0000") & " DATA: " & strdata
End Sub

Private Sub Timer1_Timer()
    'Ensure that all sockets that are not connected, are no longer listening
    'and add them back to the socket pool. What is wierd about this, is the
    'error trapping only works most of the time... other times it will ignore
    'it and still produce Error 340. Control or Array does not exist when it
    'is clearly bieng trapped as you can see below. I tried placing DoEvents,
    'and extra 'On Error GoTo errhandler' s.. but to no avail.
    '-----------> Is this a Microsoft BUG? <---------------------------------
    
    DoEvents
    On Error GoTo errhandler
    For x = 1 To host.Count
    DoEvents
    On Error GoTo errhandler
        If host(x).State = 8 Then
            host(x).Close
            List2.AddItem Format(x - 1, "0000"): Unload host(x)
            List1.AddItem "Disconnected: " & Format(x, "0000")
        End If
next1:
    Next x
Exit Sub
errhandler:
If Err.Number = 340 Then
    On Error GoTo errhandler
    GoTo next1
Else: MsgBox Err.Number & ": " & Err.Description, vbOKOnly, "Error"
End If
End Sub
