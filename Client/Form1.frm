VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example Client"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbIP 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cbDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   2040
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox tbMessages 
      Height          =   2175
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox tbSay 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cbConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbConnect_Click()
On Error GoTo Error:
Client.RemotePort = 2500
Client.RemoteHost = tbIP.Text
Client.Connect
cbConnect.Enabled = False
cbDisconnect.Enabled = True
Error: Exit Sub
End Sub

Private Sub cbDisconnect_Click()
Client.Close
cbConnect.Enabled = True
cbDisconnect.Enabled = False
End Sub

Private Sub tbSay_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim strData As String
If KeyAscii = (13) Then
strData = tbSay.Text
tbMessages.Text = tbMessages.Text & "Client: " & tbSay & vbCrLf
Client.SendData strData
tbSay.Text = ""
End If
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Client.GetData strData
tbMessages.Text = tbMessages & "Server: " & strData & vbCrLf
End Sub

Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call MsgBox(Description, vbExclamation, "Error Num." & Number)
End Sub
