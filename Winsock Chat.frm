VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chat Room"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      Height          =   1095
      Left            =   3360
      TabIndex        =   9
      Top             =   960
      Width           =   1935
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdHost 
         Caption         =   "Host"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Host/Connect"
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   1935
      Begin VB.TextBox txtLocalIP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "255.255.255.255"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   2655
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4080
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chat Room"
      Height          =   2655
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   1095
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Username"
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This tutorial was made and commented by Brad
'Distribute as you wish!!!
'June 18, 2002 12:49 AM

Private Sub cmdConnect_Click()
'If There is no username entered then don't allow to connect
If txtName = "" Then
MsgBox "Please Enter Username", vbOKOnly + vbInformation, "Enter Username"
Else
'The 'Host' and 'Port' variables just make the code look nice (not needed)
Host = txtLocalIP
Port = 4588
Winsock1.Close 'Close everything, THEN Connect
Winsock1.Connect Host, Port
txtName.Enabled = False 'Don't Let them change their name
cmdDisconnect.Enabled = True 'Let user disconnect if they please
cmdHost.Enabled = False 'Don't let the user HOST
cmdConnect.Enabled = False 'Don't let the user CONNECT
End If
End Sub

Private Sub cmdDisconnect_Click()
Winsock1.Close 'Terminate the connection
End Sub

Private Sub cmdHost_Click()
'If There is no username entered then don't allow to host
If txtName = "" Then
MsgBox "Please Enter Username", vbOKOnly + vbInformation, "Enter Username"
Else
'Define the port to listen on...
Winsock1.LocalPort = "4588"
'Then Listen
Winsock1.Listen
txtName.Enabled = False 'Don't Let them change their name
cmdDisconnect.Enabled = True 'Let user disconnect if they please
cmdHost.Enabled = False 'Don't let the user HOST
cmdConnect.Enabled = False 'Don't let the user CONNECT
End If
End Sub

Private Sub cmdSend_Click()
'When we hit Send we want to update OUR chat with what WE typed
txtChat.Text = txtChat.Text + txtName.Text + ": " + txtMessage & vbCrLf
'Then we define the Data we want to send
Data = txtName + ": " + txtMessage 'Of course the name and the Message
Winsock1.SendData Data 'Then we send the data
txtMessage = "" 'and Clear the Message box
End Sub

Private Sub Form_Load()
txtLocalIP = Winsock1.LocalIP 'Load the Local IP into the Text box
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
'When talking in chat room it's convienent to push RETURN
If KeyAscii = 13 Then
Call cmdSend_Click
End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'MUST MUST!! verify that the winsock is used already
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID 'Then Accept Connection
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock1.GetData Data 'Get the Data that was sent by user
txtChat.Text = txtChat.Text + Data & vbCrLf 'Update the Chat room
End Sub


