VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "WinTel"
   ClientHeight    =   6780
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11355
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   6780
   ScaleWidth      =   11355
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   1440
      Top             =   5280
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6600
      TabIndex        =   7
      Text            =   "9000"
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "Form1.frx":1B750
      Left            =   6840
      List            =   "Form1.frx":1B752
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   4920
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1080
      Top             =   4920
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   720
      Top             =   4920
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":1B754
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   8040
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   7920
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   3795
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":1B77F
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   360
      Top             =   4920
   End
   Begin RichTextLib.RichTextBox Text 
      Height          =   3855
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6800
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":1B954
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   5640
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":1BA29
      Top             =   1320
      Width           =   7215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   23
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu start 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu stopserver 
         Caption         =   "&Stop Server"
         Enabled         =   0   'False
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu motd 
         Caption         =   "&MOTD"
      End
      Begin VB.Menu commands 
         Caption         =   "&Commands"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu other 
      Caption         =   "&Other"
      Begin VB.Menu msguser 
         Caption         =   "&Message Users"
         Enabled         =   0   'False
      End
      Begin VB.Menu kill 
         Caption         =   "&Kill all"
         Enabled         =   0   'False
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu options 
         Caption         =   "&Options"
         Begin VB.Menu autostart 
            Caption         =   "&Auto Start Services"
         End
         Begin VB.Menu noexit 
            Caption         =   "&Notify on exit"
            Checked         =   -1  'True
         End
         Begin VB.Menu beepon 
            Caption         =   "&Beep on connection"
            Checked         =   -1  'True
         End
         Begin VB.Menu autohide 
            Caption         =   "&Auto hide on start"
         End
         Begin VB.Menu line4 
            Caption         =   "-"
         End
         Begin VB.Menu timeout 
            Caption         =   "&Timeout Delay"
         End
         Begin VB.Menu deletehisotry 
            Caption         =   "&Delete History"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu helpmenu 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
      Begin VB.Menu help 
         Caption         =   "&Help"
      End
   End
   Begin VB.Menu back 
      Caption         =   "<=="
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
MsgBox "This program isn't finished yet, but was made by ^Funny^ in vb6.", vbInformation, "About"
End Sub

Private Sub autohide_Click()
    If autohide.Checked = False Then
        autohide.Checked = True
        MsgBox "Warning, If you do this you will not be able to get it back without CTRL+ALT+DEL", vbExclamation, "Warning!"
    Else
        autohide.Checked = False
    End If
End Sub

Private Sub autostart_Click()
    If autostart.Checked = False Then
        autostart.Checked = True
    Else
        autostart.Checked = False
    End If
End Sub

Private Sub back_Click()
Text3.Visible = False
Text2.Visible = False
End Sub

Private Sub beepon_Click()
    If beepon.Checked = False Then
        beepon.Checked = True
    Else
        beepon.Checked = False
    End If
End Sub

Private Sub commands_Click()
Text3.Visible = False
Text2.Visible = True
End Sub

Private Sub deletehisotry_Click()
Dim user
    user = MsgBox("Are you sure you want to delete the history file?", vbYesNo Or vbQuestion, "Delete")
        If user = vbYes Then
            Call DeleteFile(App.Path & "/history.wintel", True)
            List1.Clear
            deletehisotry.Enabled = False
        Else
            Exit Sub
        End If
End Sub

Private Sub exit_Click()
On Error Resume Next
 Dim user, a, b As String
    If Winsock1.State = sckListening Then
        Text.SelColor = vbRed
        Text.SelText = vbCrLf & "0016 Error: You must shutdown the server first befor you can quit." & vbCrLf
    Else
            If noexit.Checked = False Then
             Winsock1.SendData vbCrLf & "Local server has requested a shutdown, Goodbye"
            Timer1.Enabled = True
       iniPath$ = App.Path & "/wintel.ini"
            entry$ = 0
        r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
            entry$ = Me.WindowState
        r% = WritePrivateProfileString("main", "windowstate", entry$, iniPath$)
            entry$ = Me.Left
        r% = WritePrivateProfileString("main", "left", entry$, iniPath$)
            entry$ = Me.Top
        r% = WritePrivateProfileString("main", "top", entry$, iniPath$)
            entry$ = autostart.Checked
        r% = WritePrivateProfileString("settings", "autostart", entry$, iniPath$)
            entry$ = noexit.Checked
        r% = WritePrivateProfileString("settings", "noexit", entry$, iniPath$)
            entry$ = beepon.Checked
        r% = WritePrivateProfileString("settings", "beepon", entry$, iniPath$)
            entry$ = autohide.Checked
        r% = WritePrivateProfileString("settings", "autohide", entry$, iniPath$)
            entry$ = Winsock1.LocalPort
        r% = WritePrivateProfileString("winsock", "port", entry$, iniPath$)
            entry$ = Winsock1.LocalIP
        r% = WritePrivateProfileString("winsock", "ip", entry$, iniPath$)
            entry$ = "Last login: " & Time & " " & Date & ", from, " & lastlogin1
        r% = WritePrivateProfileString("logon", Winsock1.RemoteHostIP, entry$, iniPath$)
            entry$ = "no"
        r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
            entry$ = Text3.Text
        r% = WritePrivateProfileString("main", "motd", entry$, iniPath$)
            entry$ = Text4.Text
        r% = WritePrivateProfileString("main", "interval", entry$, iniPath$)
            Call WriteList(List1, App.Path & "/history.wintel")
        Else
        user = MsgBox("Are you sure you want to quit?", vbYesNo Or vbQuestion, "Quit?")
            If user = vbYes Then
            On Error Resume Next
            Winsock1.SendData vbCrLf & "Local server has requested a shutdown, Goodbye" & vbCrLf & "[Wintel@Pot]$ "
            Timer1.Enabled = True
                iniPath$ = App.Path & "/wintel.ini"
entry$ = 0
r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
entry$ = Me.Left
r% = WritePrivateProfileString("main", "left", entry$, iniPath$)
entry$ = Me.Top
r% = WritePrivateProfileString("main", "top", entry$, iniPath$)
entry$ = autostart.Checked
r% = WritePrivateProfileString("settings", "autostart", entry$, iniPath$)
entry$ = noexit.Checked
r% = WritePrivateProfileString("settings", "noexit", entry$, iniPath$)
entry$ = beepon.Checked
r% = WritePrivateProfileString("settings", "beepon", entry$, iniPath$)
entry$ = autohide.Checked
r% = WritePrivateProfileString("settings", "autohide", entry$, iniPath$)
entry$ = Winsock1.LocalPort
r% = WritePrivateProfileString("winsock", "port", entry$, iniPath$)
entry$ = Winsock1.LocalIP
r% = WritePrivateProfileString("winsock", "ip", entry$, iniPath$)
entry$ = "Last login: " & Time & " " & Date & ", from, " & lastlogin1
r% = WritePrivateProfileString("logon", Winsock1.RemoteHostIP, entry$, iniPath$)
entry$ = "no"
r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
entry$ = Text3.Text
r% = WritePrivateProfileString("main", "motd", entry$, iniPath$)
entry$ = Text4.Text
r% = WritePrivateProfileString("main", "interval", entry$, iniPath$)
Call WriteList(List1, App.Path & "/history.wintel")
            Else
                Exit Sub
            End If
    End If
    End If
End Sub

Private Sub Form_Load()
Dim IP, Port, use, use1, telnet As String

On Error Resume Next
    Notontop Form2
    telnet = MsgBox("Would you like to run telnet?", vbYesNo Or vbQuestion, "Telnet")
        If telnet = vbYes Then
            Shell "C:\windows\telnet.exe", vbNormalFocus
            Ontop Form2
        Else
            Ontop Form2
            GoTo resumeeeeee
        End If
        
resumeeeeee:
Text.SelColor = vbBlack
Text.SelText = "WinTel made by ^Funny^ version 2.3b: "
Text.SelColor = vbBlack
Text.SelText = "Now running, click file\start to begin" & vbCrLf
iniPath$ = App.Path & "/wintel.ini"
Me.WindowState = GetFromINI("main", "windowstate", iniPath$)
Form1.Left = GetFromINI("main", "left", iniPath$)
Form1.Top = GetFromINI("main", "top", iniPath$)
autostart.Checked = GetFromINI("settings", "autostart", iniPath$)
noexit.Checked = GetFromINI("settings", "noexit", iniPath$)
beepon.Checked = GetFromINI("settings", "beepon", iniPath$)
autohide.Checked = GetFromINI("settings", "autohide", iniPath$)
Text4.Text = GetFromINI("main", "interval", iniPath$)
Text3.Text = GetFromINI("main", "motd", iniPath$)
IP = GetFromINI("winsock", "ip", iniPath$)
use = GetFromINI("main", "use", iniPath$)
    If use = "" Then
        Text.SelColor = vbBlue
        Text.SelText = "This is your first time using WinTel please remember to set your" & vbCrLf & "MOTD, commands and other settings up befor starting this service" & vbCrLf & vbCrLf & "if you don't understand what this program is goto menu help, for a quick description of this program" & vbCrLf & "its something like a linux shell where someone uses telnet to connect" & vbCrLf & "but is made for windows 9.x i don't know about winme and win 2k systems" & vbCrLf & "if your still having trouble goto menu help"
        entry$ = 1
        r% = WritePrivateProfileString("main", "use", entry$, iniPath$)
        GoTo resumenext
    Else
    End If
use1 = use + 1
entry$ = use1
r% = WritePrivateProfileString("main", "use", entry$, iniPath$)
resumenext:
Port = GetFromINI("winsock", "port", iniPath$)
Call ReadList(List1, App.Path & "/history.wintel", True)
    If autostart.Checked = True Then
        Text.SelColor = vbRed
        Text.SelText = vbCrLf & "Auto start has been activated, Now loading settings that were lasted used" & vbCrLf
        Text.SelColor = vbBlack
        Text.SelText = "Now listening on port, " & Port & " Localip, " & IP & " Host, " & Winsock1.LocalHostName & vbCrLf & vbCrLf
        Winsock1.Listen
        start.Enabled = False
        stopserver.Enabled = True
        msguser.Enabled = True
        kill.Enabled = True
    End If
    
    If autohide.Checked = True Then
        Me.Hide
    End If
    Form2.Hide
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 Dim user, a, b As String
    If Winsock1.State = sckListening Then
        Text.SelColor = vbRed
        Text.SelText = vbCrLf & "0016 Error: You must shutdown the server first befor you can quit." & vbCrLf
        Cancel = 1
    Else
        If noexit.Checked = False Then
             Winsock1.SendData vbCrLf & "Local server has requested a shutdown, Goodbye"
            Timer1.Enabled = True
       iniPath$ = App.Path & "/wintel.ini"
            entry$ = 0
        r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
            entry$ = Me.WindowState
        r% = WritePrivateProfileString("main", "windowstate", entry$, iniPath$)
            entry$ = Me.Left
        r% = WritePrivateProfileString("main", "left", entry$, iniPath$)
            entry$ = Me.Top
        r% = WritePrivateProfileString("main", "top", entry$, iniPath$)
            entry$ = autostart.Checked
        r% = WritePrivateProfileString("settings", "autostart", entry$, iniPath$)
            entry$ = noexit.Checked
        r% = WritePrivateProfileString("settings", "noexit", entry$, iniPath$)
            entry$ = beepon.Checked
        r% = WritePrivateProfileString("settings", "beepon", entry$, iniPath$)
            entry$ = autohide.Checked
        r% = WritePrivateProfileString("settings", "autohide", entry$, iniPath$)
            entry$ = Winsock1.LocalPort
        r% = WritePrivateProfileString("winsock", "port", entry$, iniPath$)
            entry$ = Winsock1.LocalIP
        r% = WritePrivateProfileString("winsock", "ip", entry$, iniPath$)
            entry$ = "Last login: " & Time & " " & Date & ", from, " & lastlogin1
        r% = WritePrivateProfileString("logon", Winsock1.RemoteHostIP, entry$, iniPath$)
            entry$ = "no"
        r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
            entry$ = Text3.Text
        r% = WritePrivateProfileString("main", "motd", entry$, iniPath$)
            entry$ = Text4.Text
        r% = WritePrivateProfileString("main", "interval", entry$, iniPath$)
        Call WriteList(List1, App.Path & "/history.wintel")
        Else
        user = MsgBox("Are you sure you want to quit?", vbYesNo Or vbQuestion, "Quit?")
            If user = vbYes Then
            On Error Resume Next
            Winsock1.SendData vbCrLf & "Local server has requested a shutdown, Goodbye"
            Timer1.Enabled = True
                iniPath$ = App.Path & "/wintel.ini"
entry$ = 0
r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
entry$ = Me.WindowState
r% = WritePrivateProfileString("main", "windowstate", entry$, iniPath$)
entry$ = Me.Left
r% = WritePrivateProfileString("main", "left", entry$, iniPath$)
entry$ = Me.Top
r% = WritePrivateProfileString("main", "top", entry$, iniPath$)
entry$ = autostart.Checked
r% = WritePrivateProfileString("settings", "autostart", entry$, iniPath$)
entry$ = noexit.Checked
r% = WritePrivateProfileString("settings", "noexit", entry$, iniPath$)
entry$ = beepon.Checked
r% = WritePrivateProfileString("settings", "beepon", entry$, iniPath$)
entry$ = autohide.Checked
r% = WritePrivateProfileString("settings", "autohide", entry$, iniPath$)
entry$ = Winsock1.LocalPort
r% = WritePrivateProfileString("winsock", "port", entry$, iniPath$)
entry$ = Winsock1.LocalIP
r% = WritePrivateProfileString("winsock", "ip", entry$, iniPath$)
entry$ = "Last login: " & Time & " " & Date & ", from, " & lastlogin1
r% = WritePrivateProfileString("logon", Winsock1.RemoteHostIP, entry$, iniPath$)
entry$ = "no"
r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
entry$ = Text3.Text
r% = WritePrivateProfileString("main", "motd", entry$, iniPath$)
entry$ = Text4.Text
r% = WritePrivateProfileString("main", "interval", entry$, iniPath$)
Call WriteList(List1, App.Path & "/history.wintel")
            Else
                Cancel = 1
            End If
    End If
    End If
End Sub

Private Sub help_Click()
MsgBox "Help" & vbCrLf & "-------------" & vbCrLf & "Wintel made for windows 9.x systems don't know about winme or 2k systems" & vbCrLf & vbCrLf & "What this service is, is something like linux shell hosting BUT is made for windows systems." & vbCrLf & "Reason i made this is cause i was looking all over the net for one and could not find one" & vbCrLf & "all you have to do is click file\start and read instructions on what todo" & vbCrLf & "when the services are started goto start, run, and type in 'telnet'" & vbCrLf & "in telnet connect to (" & Winsock1.LocalIP & ") defualt telnet port (23)" & vbCrLf & "when your connected type in the defualt password and login login: '1' Password: '2' (Can be edited in the wintel.ini file under [login] edit login= and pass=" & vbCrLf & "it can only be 1 letter word or number sorry that will be fixed in newer versions" & vbCrLf & vbCrLf & "if you need any other help please icq 13237186 or email hax03d@hotmail.com" & vbCrLf & vbCrLf & "Enjoy!", vbInformation, "Help"
MsgBox "Opps forgot one thing... one 1 user can be logged in at a time and one 1 login and password" & vbCrLf & "can be issused to users, and you must shutdown the services befor you can quit.. other than that im pretty sure im finished =)" & vbCrLf & vbCrLf & "All graphics made by Patrick are copyrighted (c) 2001" & vbCrLf & "WinTel is also copyrighted (c) 2001 all laws will be inforced" & vbCrLf & "if needed, otherwise enjoy."
End Sub

Private Sub kill_Click()
Winsock1.SendData vbCrLf & "Local server has disconnected all users, Goodbye"
Text.SelColor = vbRed
Text.SelText = "0019 All clients have been disconnected" & vbCrLf
Timer3.Enabled = True
    iniPath$ = App.Path & "/wintel.ini"
entry$ = 0
r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
entry$ = "Last login: " & Time & " " & Date & ", from, " & lastlogin1
r% = WritePrivateProfileString("logon", Winsock1.RemoteHostIP, entry$, iniPath$)
entry$ = "no"
r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
End Sub

Private Sub motd_Click()
Text3.Visible = True
Text2.Visible = False
End Sub

Private Sub msguser_Click()
On Error Resume Next
 Dim Msg As String
    Msg = InputBox("Enter message you wish to send", "Message")
    Winsock1.SendData vbCrLf & "Message From your admin: " & Msg & vbCrLf & "[Wintel@pot]$ "
    Text.SelColor = vbRed
    Text.SelText = "0018 Message to users: "
    Text.SelColor = vbBlack
    Text.SelText = Msg & vbCrLf
End Sub

Private Sub noexit_Click()
    If noexit.Checked = False Then
        noexit.Checked = True
    Else
        noexit.Checked = False
    End If
End Sub

Private Sub start_Click()
 On Error Resume Next
 Dim user, Answer, answer1 As String
 user = InputBox("Which port do you wish to open for WinTel? (Default, = 23)", "Port", "23")
 Answer = MsgBox("Is the ip address below what you wish to use?" & vbCrLf & Winsock1.LocalIP, vbYesNo Or vbQuestion, "IP")
        If Answer = vbYes Then
            answer1 = Winsock1.LocalIP
        Else
            dataX = InputBox("Please enter a host or ip you wish to use than", "IP")
        End If
    Winsock1.LocalPort = user
    Text.SelColor = vbRed
    Text.SelText = vbCrLf & "0013 Now listening on port: "
    Text.SelColor = vbBlack
    Text.SelText = user & vbCrLf
    Text.SelColor = vbRed
    Text.SelText = "0014 Local Hostname: "
    Text.SelColor = vbBlack
    Text.SelText = Winsock1.LocalHostName & ", " & Winsock1.LocalIP & vbCrLf & vbCrLf
    Text.SelText = dataX
    start.Enabled = False
    stopserver.Enabled = True
    Winsock1.Listen
    kill.Enabled = True
    msguser.Enabled = True
End Sub

Private Sub stopserver_Click()
 Dim user As String
 user = MsgBox("Are you sure you want to shutdown the server?", vbYesNo Or vbQuestion, "Shutdown")
    If user = vbYes Then
        Winsock1.Close
        Text.SelColor = vbRed
        Text.SelText = vbCrLf & "0015 Server Shutdown, All users disconnected" & vbCrLf
        start.Enabled = True
        kill.Enabled = False
        msguser.Enabled = False
    stopserver.Enabled = False
    iniPath$ = App.Path & "/wintel.ini"
entry$ = 0
r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
entry$ = "Last login: " & Time & " " & Date & ", from, " & lastlogin1
r% = WritePrivateProfileString("logon", Winsock1.RemoteHostIP, entry$, iniPath$)
entry$ = "no"
r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
    Else
        Exit Sub
    End If

End Sub



Private Sub timeout_Click()
Dim user As String
user = InputBox("Enter timeout delay so if a user takes more than & vbcrlf & eg 10 seconds to enter a login and password they get & vbcrlf & disconnected defualt is 10 seconds & vbcrlf & vbcrlf & (1000 = 1 second 10000 = 10 seconds)", "Timeout", "1000")
Text4.Text = user
End Sub

Private Sub Timer1_Timer()
End
End Sub


Private Sub Timer2_Timer()
    If Text3.Visible Or Text2.Visible = True Then
        back.Enabled = True
    Else
        back.Enabled = False
    End If
End Sub

Private Sub Timer3_Timer()
Winsock1.Close
Winsock1.Listen
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Dim a, b As Integer
a = (Int(Rnd * 3) + 1)
b = (Int(Rnd * 3) + a)
Dim Secs, Mins, Hours, Days
Dim TotalMins, TotalHours, TotalSecs, TempSecs
Dim CaptionText
TotalSecs = Int(GetTickCount / 1000)
Days = Int(((TotalSecs / 60) / 60) / 24)
TempSecs = Int(Days * 86400)
TotalSecs = TotalSecs - TempSecs
TotalHours = Int((TotalSecs / 60) / 60)
TempSecs = Int(TotalHours * 3600)
TotalSecs = TotalSecs - TempSecs
TotalMins = Int(TotalSecs / 60)
TempSecs = Int(TotalMins * 60)
TotalSecs = (TotalSecs - TempSecs)
    If TotalHours > 23 Then
        Hours = (TotalHours - 23)
    Else
        Hours = TotalHours
    End If
    If TotalMins > 59 Then
        Mins = (TotalMins - (Hours * 60))
    Else
        Mins = TotalMins
    End If
        Winsock1.SendData vbCrLf & Time & " up " & Days & " days, " & Hours & " hours. Users: 1, Load average: 00." & a & ", 00.0, 00." & b & vbCrLf & "[Wintel@Pot]$ "
    Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
Dim lastlogin, lastlogin1, use As String
iniPath$ = App.Path & "/wintel.ini"
        use = GetFromINI("main", "use", iniPath$)
        lastlogin = GetFromINI("logon", "lastid", iniPath$)
        lastlogin1 = GetFromINI("logon", "lastip", iniPath$)
Winsock1.SendData vbCrLf & "........................." & vbCrLf
Winsock1.SendData vbCrLf & vbCrLf & "Listing." & vbCrLf
Dim i As Integer
For i = 0 To List1.ListCount - 1
    Winsock1.SendData List1.List(i) & " from history.wintel" & vbCrLf
Next i
Winsock1.SendData vbCrLf & vbCrLf & "Saving logs...<" & List1.ListCount & "> Saved." & vbCrLf
Winsock1.SendData "Saving uses...<" & use & "> Saved." & vbCrLf
Winsock1.SendData "Please wait..." & vbCrLf
Winsock1.SendData "One momment..." & vbCrLf
Winsock1.SendData "Sorry if there was a wait..." & vbCrLf
Winsock1.SendData "Configuring WinTel....Done" & vbCrLf
Winsock1.SendData "Finished" & vbCrLf
Winsock1.SendData vbCrLf & "Finishing settings" & vbCrLf & "[Wintel@Pot]$ "
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
On Error Resume Next
iniPath$ = App.Path & "/wintel.ini"
Dim data As String
data = GetFromINI("main", "Logged", iniPath$)
    If data = 0 Then
        Text.SelText = vbCrLf & "User being logged out, took to long to login" & vbCrLf
        Winsock1.SendData vbCrLf & vbCrLf & "You are now being logged out, reason: Took to long to login"
        Timer3.Enabled = True
        Timer6.Enabled = False
        entry$ = "no"
        r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
    Else
        Timer6.Enabled = False
        Exit Sub
    End If
End Sub

Private Sub Timer7_Timer()
On Error Resume Next
Timer6.Interval = Text4.Text
End Sub

Private Sub Winsock1_Close()
Text.SelColor = vbRed
Text.SelText = vbCrLf & "0021 User disconnected allowing another user to connect"
iniPath$ = App.Path & "/wintel.ini"
entry$ = 0
r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)

entry$ = "Last login: " & Time & " " & Date & ", from, " & lastlogin1
r% = WritePrivateProfileString("logon", Winsock1.RemoteHostIP, entry$, iniPath$)
entry$ = "no"
r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)

Winsock1.Close
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Open App.Path & "/" & Winsock1.RemoteHostIP & ".log" For Output As #2
        Timer6.Enabled = True
    If beepon.Checked = True Then
        Beep
        GoTo nextnext
    Else
nextnext:
Text2.Visible = False
    Winsock1.Close
    Winsock1.Accept requestID
    Winsock1.SendData "WinTel Made by ^Funny^, used for windows 9.x systems" & vbCrLf & vbCrLf
    Winsock1.SendData "MOTD: " & Text3.Text & vbCrLf & vbCrLf
    Winsock1.SendData "Login:"
    List1.AddItem "#-" & Winsock1.RemoteHostIP & " at " & Time & " " & Date
    deletehisotry.Enabled = True
    Text.SelColor = vbRed
    Text.SelText = vbCrLf & "0017 Client Connecting: "
    Text.SelColor = vbBlack
    Text.SelText = Winsock1.RemoteHostIP & ", " & Winsock1.RemoteHost & vbCrLf
    Text.SelStart = Len(Text.Text)
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dataX As String
Dim lastlogin, lastlogin1 As String
iniPath$ = App.Path & "/wintel.ini"
lastlogin = GetFromINI("logon", Winsock1.RemoteHostIP, iniPath$)
Dim Login, Pass As String
Login = GetFromINI("login", "login", iniPath$)
Pass = GetFromINI("login", "pass", iniPath$)
Winsock1.GetData dataX

    If dataX = Login Then
        Winsock1.SendData vbCrLf & "Login accepted, need password" & vbCrLf
        Winsock1.SendData "Password:"
        entry$ = "yes"
        r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
        GoTo notlogged
    End If
    
    If dataX = Pass Then
        Dim lag
        lag = GetFromINI("main", "passwordlogin", iniPath$)
            If lag = "yes" Then
                Winsock1.SendData vbCrLf & lastlogin & " " & Winsock1.RemoteHostIP & vbCrLf & vbCrLf & "[Wintel@Pot]$ "
                entry$ = 1
                r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
            Else
                Winsock1.SendData vbCrLf & "Invalid login." & vbCrLf & vbCrLf & "Login: "
                GoTo notlogged
            End If
    End If
    
    On Error Resume Next
    
Dim i As String
i = GetFromINI("main", "Logged", iniPath$)

    If i = "0" Then
        Winsock1.SendData vbCrLf & "Invalid login." & vbCrLf & vbCrLf & "Login: "
        Text.SelColor = vbRed
        Text.SelText = vbCrLf & "0020 User tried to use commands (But not logged in)"
        GoTo notlogged
    Else
    
        If dataX = "h" Then
            Winsock1.SendData vbCrLf & Text2.Text & vbCrLf & "[Wintel@Pot]$ "
            Text.SelColor = vbRed
            Text.SelText = vbCrLf & "0021 Requested Help File"
        ElseIf dataX = "q" Then
            Text.SelText = vbCrLf & "0021 requested logout"
            Winsock1.SendData vbCrLf & "Your now being logged out.." & vbCrLf
            Winsock1.SendData "Saving logs..." & vbCrLf
            Winsock1.SendData "Saving uses..." & vbCrLf
            Winsock1.SendData "Please wait..." & vbCrLf
            Winsock1.SendData "One momment..." & vbCrLf
            Winsock1.SendData "Sorry if there was a wait..." & vbCrLf
            Winsock1.SendData "getting ready to shutdown current connection, opening another" & vbCrLf
            Winsock1.SendData "Finished." & vbCrLf
            Winsock1.SendData "Finishing up" & vbCrLf
            Winsock1.SendData "Saving Login and password.." & vbCrLf
            Winsock1.SendData "Have a nice day, Goodbye." & vbCrLf
            entry$ = 0
            r% = WritePrivateProfileString("main", "Logged", entry$, iniPath$)
            entry$ = "no"
            r% = WritePrivateProfileString("main", "passwordlogin", entry$, iniPath$)
            Timer3.Enabled = True
        ElseIf dataX = "m" Then
            Winsock1.SendData vbCrLf & Text3.Text & "[Wintel@Pot]$ "
            Text.SelText = vbCrLf & "0021 requested Message of the day"
        ElseIf dataX = "d" Then
            Text.SelText = vbCrLf & "0021 requested Directory listing"
            Winsock1.SendData vbCrLf & vbCrLf & "File System" & vbCrLf
            Winsock1.SendData "=============" & vbCrLf
            Dim cnt As Integer
            For cnt1 = 0 To Dir1.ListCount - 1
                Winsock1.SendData "Dir: " & Dir1.List(cnt1) & vbCrLf
            Next cnt1
            Winsock1.SendData "" & vbCrLf
            For cnt = 0 To File1.ListCount
                Winsock1.SendData "File: " & File1.List(cnt) & ", " & FileLen(File1.List(cnt)) & " bytes" & vbCrLf
            Next cnt
            Winsock1.SendData vbCrLf & "# of files: <" & File1.ListCount & ">" & vbCrLf
            Winsock1.SendData "# of Directorys: <" & Dir1.ListCount & ">" & vbCrLf
            Winsock1.SendData vbCrLf & vbCrLf & "End of: " & Dir1.Path & vbCrLf & vbCrLf & "to change directorys type <cd Directory> if you wish to download" & vbCrLf & "If the admin has the ftp up and setup you can connect to" & vbCrLf & Winsock1.LocalIP & " Password and login for this shell" & vbCrLf & vbCrLf & "[Wintel@Pot]$ "
        ElseIf dataX = "c" Then
            Winsock1.SendData vbCrLf & "Change Login and password is not setup yet" & vbCrLf & "[Wintel@Pot]$ "
        ElseIf dataX = "t" Then
            Winsock1.SendData vbCrLf & "Current Time: " & Time & " Current Date: " & Date & vbCrLf & "[Wintel@Pot]$ "
            Text.SelText = vbCrLf & "0021 User requested time and date"
        ElseIf dataX = "u" Then
            Timer4.Enabled = True
            Text.SelText = vbCrLf & "0021 User requested uptime"
        ElseIf dataX = "v" Then
            Winsock1.SendData vbCrLf & "WinTel made for windows 9.x made by ^Funny^ ver: 3.2.043" & vbCrLf
            Timer4.Enabled = True
            Text.SelText = vbCrLf & "0021 User requested version information"
        ElseIf dataX = "f" Then
            Winsock1.SendData vbCrLf & "FTP SERVER" & vbCrLf & "==============" & vbCrLf
            Winsock1.SendData "Login: wintel" & vbCrLf & "Password: pot" & vbCrLf & "Port: " & Winsock1.LocalPort & vbCrLf & "IP: " & Winsock1.LocalIP & vbCrLf & vbCrLf & "FTP is now setup you can allow users to connect with ftp using the above settings" & vbCrLf & "[Wintel@Pot]$ "
            Text.SelText = vbCrLf & "0021 User requested ftp setup"
        ElseIf dataX = "i" Then
            Timer6.Enabled = True
            Winsock1.SendData vbCrLf & "Configuring WinTel for better use" & vbCrLf & vbCrLf
            Winsock1.SendData "Checking for all passwords and login atempts"
            Timer5.Enabled = True
            Text.SelText = vbCrLf & "0021 User requested a WinTel re-configuration"
    End If
             End If
Text.SelColor = vbRed
Text.SelText = vbCrLf & "0019: " & dataX
Text.SelStart = Len(Text.Text)
Print #2, "#-" & Time & ": " & dataX
notlogged:

End Sub
