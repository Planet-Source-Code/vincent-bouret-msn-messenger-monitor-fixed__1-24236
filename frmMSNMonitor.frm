VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMSNMonitor 
   Caption         =   "MSN Messenger Monitor"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wskComNS 
      Left            =   6120
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin MSWinsockLib.Winsock wskCom 
      Left            =   5160
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.CommandButton cListen 
      Caption         =   "Listen"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   6360
      Width           =   735
   End
   Begin MSWinsockLib.Winsock wskServDS 
      Left            =   4680
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
      LocalPort       =   1863
   End
   Begin VB.TextBox tLocalPort 
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Text            =   "1864"
      Top             =   6360
      Width           =   735
   End
   Begin VB.Frame frDS 
      Caption         =   "Dispatch Server"
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   8175
      Begin VB.TextBox tDS 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   7935
      End
   End
   Begin MSWinsockLib.Winsock wskNS 
      Left            =   8040
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin MSWinsockLib.Winsock wskDS 
      Left            =   8040
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.Frame frNS 
      Caption         =   "Notification Server"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   8175
      Begin VB.TextBox tNS 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   7935
      End
   End
   Begin MSWinsockLib.Winsock wskServNS 
      Left            =   5640
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
      LocalPort       =   1864
   End
   Begin VB.Label lMonitor 
      Caption         =   "Open port on localhost for monitoring:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label lServState 
      Caption         =   "Not connected"
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lMSNServ 
      Caption         =   "MSN Server:"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lMSNClient 
      Caption         =   "MSN Messenger:"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lCliState 
      Caption         =   "Not connected"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lCon 
      Alignment       =   2  'Center
      Caption         =   "Connection state :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image icoNet 
      Height          =   480
      Left            =   240
      Picture         =   "frmMSNMonitor.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMSNMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In MSN Messenger Service, the protocol is divided into 3 kinds of server. Dispatch server
'which redirect the client to the appropriate Notification Server, where the client will connect
'and receive information about contacts online
'For the complete MSN Messenger Protocol goto: http://www.tlsecurity.net/Textware/Misc/draft-movva-msn-messenger-protocol-00.txt
'Dont forget to change your registry settings:
'HKEY_CURRENT_USER\Software\Microsoft\MessengerService\' find the key "Server" and change it to 127.0.0.1;127.0.0.1:1863
'dont forget to backup your original values. If you already erased it, the initial value that should work
'is messenger.hotmail.com;127.0.0.1:1863

'Important Note: Since this program was made with VB 5.0, I've added a Split function found on www.planet-source-code
'I'm not pretty sure of the VB6 syntax but if it's not exactly like mine, please change every Split for another
'name. Sorry for inconvenience

Private Sub Form_Load()
wskServDS.Listen

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub wskCom_DataArrival(ByVal bytesTotal As Long)
Dim stemp As String

wskCom.GetData stemp
tDS.Text = tDS.Text & vbCrLf & "..." & stemp
Do Until wskDS.State = 7
    DoEvents
Loop
wskDS.SendData stemp

End Sub

Private Sub wskComNS_DataArrival(ByVal bytesTotal As Long)

Dim stemp As String

wskComNS.GetData stemp
tNS.Text = tNS.Text & vbCrLf & "..." & stemp
If stemp = "OUT" & vbCrLf Then
    wskNS.SendData stemp
    wskComNS.Close
    wskNS.Close
    lCliState.Caption = "Not connected"
    lServState.Caption = "Not connected"
End If
Do Until wskNS.State = 7
    DoEvents
Loop
wskNS.SendData stemp

End Sub


Private Sub wskDS_DataArrival(ByVal bytesTotal As Long)

Dim stemp As String
Dim DSArgs() As String
Dim AddIp() As String

wskDS.GetData stemp
tDS.Text = tDS.Text & vbCrLf & stemp

Split stemp, DSArgs(), " "
If DSArgs(1) = "XFR" And DSArgs(3) = "NS" Then 'In MSN Protocol, XFR is the command to redirect a client
    wskDS.Close
    Split DSArgs(4), AddIp(), ":"
    wskNS.RemoteHost = AddIp(1)
    wskNS.RemotePort = AddIp(2)
    wskServNS.Listen
    wskServDS.Close
    'XFR 37 NS 64.4.13.218:1863 0 <- What we will do is trafficate this command from the server to give a local adress instead
    stemp = "XFR " & DSArgs(2) & " NS " & "127.0.0.1:" & tLocalPort.Text & " 0" & vbCrLf
End If


wskCom.SendData stemp

End Sub

Private Sub wskNS_DataArrival(ByVal bytesTotal As Long)

Dim stemp As String

wskNS.GetData stemp
tNS.Text = tNS.Text & vbCrLf & stemp
wskComNS.SendData stemp

End Sub

Private Sub wskServDS_ConnectionRequest(ByVal requestID As Long)
wskCom.Close
wskCom.Accept requestID
wskDS.Connect "messenger.hotmail.com", 1863
lCliState.Caption = "Connected"
lServState.Caption = "Connected"
End Sub

Private Sub wskServNS_ConnectionRequest(ByVal requestID As Long)

wskComNS.Close
wskComNS.Accept requestID
wskNS.Connect


End Sub

