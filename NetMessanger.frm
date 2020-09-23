VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNetMessanger 
   Caption         =   "Network Communicator"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSendPM 
      Caption         =   "Send Private Message"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton cmdNotTop 
      Caption         =   "Not Always On Top"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Data"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3360
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbTalk 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4048
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"NetMessanger.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbSend 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"NetMessanger.frx":00D5
   End
   Begin MSWinsockLib.Winsock tcpSock 
      Left            =   2040
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      RemotePort      =   7777
      LocalPort       =   7777
   End
   Begin RichTextLib.RichTextBox rtbFind 
      Height          =   855
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393217
      TextRTF         =   $"NetMessanger.frx":01AA
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   4560
      Y1              =   2350
      Y2              =   2350
   End
End
Attribute VB_Name = "frmNetMessanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YourName

Private Sub cmdNotTop_Click()
    SetWindowPos frmNetMessanger.hWnd, HWND_NOTOPMOST, frmNetMessanger.Left / 15, frmNetMessanger.Top / 15, frmNetMessanger.Width / 15, frmNetMessanger.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    AOTValue = False
rtbSend.SetFocus
End Sub

Private Sub cmdSend_Click()
    SendtheData
End Sub

Private Sub SendtheData()
On Error Resume Next
tcpSock.SendData YourName & ": " & rtbSend.Text
If Not rtbTalk.Text = "" Then rtbTalk.Text = rtbTalk.Text & vbCrLf & YourName & ": " & rtbSend.Text Else rtbTalk.Text = YourName & ": " & rtbSend.Text
rtbSend.Text = ""
rtbSend.SetFocus
End Sub

Private Sub cmdSendPM_Click()
On Error Resume Next
ToWhom = InputBox("Who do you want to send the private message to?", "Send Private Message", "Type User Name Here")
    tcpSock.SendData "onlyfor" & ToWhom & YourName & " (private message): " & rtbSend.Text
        If Not rtbTalk.Text = "" Then rtbTalk.Text = rtbTalk.Text & vbCrLf & YourName & "(private message to " & ToWhom & "): " & rtbSend.Text Else rtbTalk.Text = YourName & "(private message to " & ToWhom & "): " & rtbSend.Text
    rtbSend.Text = ""
    rtbSend.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
    SetWindowPos frmNetMessanger.hWnd, HWND_TOPMOST, frmNetMessanger.Left / 15, frmNetMessanger.Top / 15, frmNetMessanger.Width / 15, frmNetMessanger.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    AOTValue = True
    
    Do Until Not YourName = "" And Len(YourName) <= 15
    YourName = InputBox("Name please...", "Enter User Name", "Maximum Length: 15 Characters")
    Loop
    
        tcpSock.Protocol = sckUDPProtocol
        tcpSock.LocalPort = 7777
        tcpSock.RemotePort = 7777
        tcpSock.RemoteHost = "255.255.255.255"
        
tcpSock.SendData "User On: " & YourName
End Sub

Private Sub Form_Unload(Cancel As Integer)
tcpSock.SendData "User Off: " & YourName
End Sub

Private Sub rtbSend_Change()
If rtbSend.Text = vbCrLf Then rtbSend.Text = ""
If Not rtbSend.Text = "" Then cmdSend.Enabled = True: cmdSendPM.Enabled = True Else cmdSend.Enabled = False: cmdSendPM.Enabled = False
End Sub

Private Sub rtbTalk_Change()
XX = Len(rtbTalk.Text)
rtbTalk.SelStart = XX - 1: rtbTalk.SelLength = 1
rtbTalk.SelLength = 0
End Sub

Private Sub rtbSend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendtheData
    rtbSend.Text = ""
End If
End Sub

Private Sub tcpSock_DataArrival(ByVal bytesTotal As Long)
Dim NewMsg As String
tcpSock.GetData NewMsg

rtbFind.Text = NewMsg
rtbFind.SelStart = 0: rtbFind.SelLength = 7 + Len(YourName)
If rtbFind.SelText = "onlyfor" & YourName Then
    rtbFind.SelText = ""
    
    If Not rtbTalk.Text = "" Then rtbTalk.Text = rtbTalk.Text & vbCrLf & rtbFind.Text Else rtbTalk.Text = rtbTalk.Text & rtbFind.Text

Exit Sub
End If

rtbFind.Text = NewMsg
rtbFind.SelStart = 0: rtbFind.SelLength = 7
If rtbFind.SelText = "onlyfor" Then Exit Sub


rtbFind.Text = NewMsg
rtbFind.SelStart = 0: rtbFind.SelLength = 41 + Len(YourName)
If rtbFind.SelText = "User Available Agknowledgement of " & YourName & " From: " Then
rtbFind.SelText = ""

If rtbFind.Text = YourName Then Exit Sub

    If rtbTalk.Text = "" Then rtbTalk.Text = "User Available Agknowledgement of " & YourName & " From: " & rtbFind.Text Else rtbTalk.Text = rtbTalk.Text & vbCrLf & "User Available Agknowledgement of " & YourName & " From: " & rtbFind.Text
Exit Sub
End If

rtbFind.Text = NewMsg
rtbFind.SelStart = 0: rtbFind.SelLength = Len(YourName)
If Not rtbFind.SelText = YourName Then
    If Not rtbTalk.Text = "" Then rtbTalk.Text = rtbTalk.Text & vbCrLf & NewMsg Else rtbTalk.Text = rtbTalk.Text & NewMsg
End If

rtbFind.Text = NewMsg
rtbFind.SelStart = 0: rtbFind.SelLength = 9
If rtbFind.SelText = "User On: " Then
rtbFind.SelText = ""
tcpSock.SendData "User Available Agknowledgement of " & rtbFind.Text & " From: " & YourName
End If

    SetWindowPos frmNetMessanger.hWnd, HWND_TOPMOST, frmNetMessanger.Left / 15, frmNetMessanger.Top / 15, frmNetMessanger.Width / 15, frmNetMessanger.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    AOTValue = True
rtbSend.SetFocus
End Sub
