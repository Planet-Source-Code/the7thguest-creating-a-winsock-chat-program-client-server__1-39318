VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example Chat Server"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Server 
      Left            =   2040
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cbStop 
      Caption         =   "Stop Server"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cbStart 
      Caption         =   "Start Server"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox tbSay 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox tbMessages 
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbStart_Click()
cbStop.Enabled = True
cbStart.Enabled = False
Server.LocalPort = 2500 'Sets the servers local port to 2500
Server.Listen           'Tells the server to listen for incoming connections
End Sub

Private Sub cbStop_Click()
cbStart.Enabled = True
cbStop.Enabled = False
Server.Close            'Tells the server to close
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)
Server.Close            'Establishes the connection
Server.Accept requestID 'Gets the connected computers ID
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call MsgBox(Description, bvExclimation, "Error Num." & Number)
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Server.GetData strData
tbMessages.Text = tbMessages & "Client: " & strData & vbCrLf
End Sub

Private Sub tbSay_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim strMessage As String
If KeyAscii = (13) Then
strMessage = tbSay.Text
tbMessages.Text = tbMessages.Text & "Server: " & tbSay & vbCrLf
Server.SendData strMessage
tbSay.Text = ""
End If
End Sub
