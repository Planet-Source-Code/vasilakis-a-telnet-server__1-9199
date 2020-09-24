VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TelnetServer by Vasilis Sagonas"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7500
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdduser 
      Caption         =   "Add User"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox lstUser 
      Height          =   3000
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   1455
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1920
      Width           =   4695
   End
   Begin VB.ListBox lstPass 
      Enabled         =   0   'False
      Height          =   2790
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   360
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   23
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Logging :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Userlist :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Pass"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Authorized(who As Integer)
On Error Resume Next
Dim ws As Winsock
Clients(who).Login = True
wsock(who).SendData vbCrLf
wsock(who).SendData " -= Log on Succesfull." & vbCrLf
wsock(who).SendData vbCrLf
wsock(who).SendData "  Welcome at the Control Center!" & vbCrLf
wsock(who).SendData "  If you have any problem, send an email to 'vsag@forthnet.gr'..." & vbCrLf
wsock(who).SendData vbCrLf
DoEvents
Clients(who).Action = "COMM"
SendName who
End Sub




Sub SaveUsers()
'**************************************************************
'Save Users to registry
it = 0
For I = 0 To lstUser.ListCount - 1
    SaveSetting "TelnetServer by Vasilis Sagonas", "User" & I, "Username", lstUser.List(I)
    SaveSetting "TelnetServer by Vasilis Sagonas", "User" & I, "Password", lstPass.List(I)
it = it + 1
Next I
    SaveSetting "TelnetServer by Vasilis Sagonas", "User" & it, "Username", ""
    SaveSetting "TelnetServer by Vasilis Sagonas", "User" & it, "Password", ""
End Sub

Sub Welcome(Index)
For I = 1 To 24
    wsock(Index).SendData String$(80, "-") & vbCrLf
    DoEvents
Next I
wsock(Index).SendData vbCrLf & vbTab & vbTab & "       Easy Telnet Server by Vasilakis" & vbCrLf
    DoEvents
For I = 1 To 22
    wsock(Index).SendData String$(80, " ") & vbCrLf
    DoEvents
Next I
ScreenCls Index
wsock(Index).SendData vbCrLf & vbTab & vbTab & "       Easy Telnet Server by Vasilakis" & vbCrLf & vbCrLf
word = vbTab & vbTab & "    *    Authorized connections only    *    "
For I = 1 To Len(word)
    letr = Right(Left(word, I), 1)
    wsock(Index).SendData letr
    DoEvents
Next I
word = vbCrLf & vbCrLf & vbTab & vbTab & "    *     User Access Verification      *    " & vbCrLf & vbCrLf
For I = 1 To Len(word)
    letr = Right(Left(word, I), 1)
    wsock(Index).SendData letr
    DoEvents
Next I
DoEvents
wsock(Index).SendData vbCrLf & " Username: "
DoEvents
Clients(Index).Action = "user"
End Sub

Private Sub cmdAdduser_Click()
If txtUser.Text = "" Then txtUser.SetFocus: Exit Sub
If txtPass.Text = "" Then txtPass.SetFocus: Exit Sub
For I = 0 To lstUser.ListCount - 1
    If LCase$(lstUser.List(I)) = LCase$(txtUser.Text) Then
        txtUser.SetFocus
        txtUser.Text = ""
        txtUser.SelStart = 0
        txtUser.SelLength = Len(txtUser.Text)
        Exit Sub
    End If
Next I
lstUser.AddItem txtUser.Text
lstPass.AddItem txtPass.Text
txtUser.Text = ""
txtPass.Text = ""
SaveUsers
End Sub

Private Sub cmdDelete_Click()
rUser = lstUser.ListIndex
lstUser.RemoveItem rUser
lstPass.RemoveItem rUser
SaveUsers
cmdDelete.Enabled = False
End Sub

Private Sub Form_Load()

'**************************************************************
' Welcome to TELNET SERVER by Vasilis Sagonas!!!
' The easiest to use and change TelnetServer ever made!
' You can freely grab any code or use it as is adding your functions!
' Thanks for using it and PLEASE if you ever use it, add a thanks
' to my name with my email ;-)
' Sorry for my bad code i am only 17 years old!
'
' T H A N X !!!
'
' Vasilis Sagonas
' vsag@forthnet.gr - vasilis@lar.forthnet.gr
'
' Contact info:
'**************************************************************
' Vasilis Sagonas
' Anthimou Gazi 48, 41222
' LARISSA, GREECE
' 0030 - 41 - 612941
'**************************************************************
' For anyone who needs a packet filter for advanced Winsock
' use (internet applications), email me!
' I also created a very cool DATAPIPE utility for windows!
' Works with telnet like datapipe in linux does!
'**************************************************************
' Please Download NetFLY! v1.5 Administration Utility to see
' my best work ever!... Thanks!
'**************************************************************
' http://www.dawn.gr/netfly/
' webmaster@dawn.gr
'**************************************************************



CenterForm Me

'**************************************************************
'Load users from registry
it = 0
Do
    rUser = GetSetting("TelnetServer by Vasilis Sagonas", "User" & it, "Username", "")
    rPass = GetSetting("TelnetServer by Vasilis Sagonas", "User" & it, "Password", "")
    If rUser = "" Then Exit Do
    it = it + 1
    lstUser.AddItem rUser
    lstPass.AddItem rPass
Loop

'**************************************************************
'Initialize socket

wsock(0).Close
wsock(0).LocalPort = 23
wsock(0).RemotePort = 0
wsock(0).Listen

'**************************************************************
End Sub



Sub ScreenCls(Index)
wsock(Index).SendData Chr$(27) & "[2J": DoEvents
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstUser_Click()
cmdDelete.Enabled = True
End Sub


Private Sub txtLog_Change()
txtLog.SelStart = Len(txtLog.Text)
txtLog.SelLength = Len(txtLog.Text)
End Sub

Private Sub wsock_Close(Index As Integer)
On Error Resume Next
txtLog.Text = txtLog.Text & wsock(Index).RemoteHostIP & " closed." & vbCrLf
wsock(Index).Close
Clients(Index).Name = ""
Clients(Index).Action = ""
Unload wsock(Index)
End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
txtLog.Text = txtLog.Text & wsock(Index).RemoteHostIP & " connected." & vbCrLf
On Error Resume Next
ic = 0
Do
    ic = ic + 1
    Err = 0
    Load wsock(ic)
    If Err = 0 Or wsock(ic).State = sckClosed Then
        ReDim Preserve Clients(wsock.Count + 1)
        Clients(ic).Action = ""
        Clients(ic).Name = ""
        Load wsock(ic)
        wsock(ic).Close
        DoEvents
        wsock(ic).Accept requestID
        Welcome ic
        Exit Sub
    End If
Loop
End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim DNSa As String
Dim vtData As String
Dim Comm As String
Dim cmd As String
Dim ws As Winsock
wsock(Index).GetData vtData
If Clients(Index).Action = "" Then Exit Sub
If InStr(wsock(Index).Tag, vbCrLf) Then
cmd:
    If cmd <> "" Then Clients(Index).LastCommand = cmd 'Give the last command in case the UP button is pressed.
    cmd = wsock(Index).Tag
    wsock(Index).Tag = ""
    'Checks current user state. If he gives the password, username or can use the commands
    Select Case Clients(Index).Action
        Case "user" 'If just entered the USERNAME
            Clients(Index).Action = "pass"
            Clients(Index).Name = cmd
            wsock(Index).SendData vbCrLf
            wsock(Index).SendData " Password: "
            DoEvents
        Case "pass" 'If just entered the password.
            wsock(Index).SendData vbCrLf
            For iTMP = 0 To lstUser.ListCount - 1
                If LCase$(Clients(Index).Name) = LCase$(lstUser.List(iTMP)) And cmd = lstPass.List(iTMP) Then
                    Clients(Index).Action = ""
                    Authorized Index
                    Exit Sub
                ElseIf LCase$(Clients(Index).Name) = LCase$(lstUser.List(iTMP)) And cmd <> lstPass.List(iTMP) Then
                    Clients(Index).Action = ""
                    Clients(Index).Attempts = Clients(Index).Attempts + 1
                    If Clients(Index).Attempts = 3 Then
                        wsock(Index).SendData vbCrLf & vbCrLf & " -= Too many invalid username/passwords. Closing..." & vbCrLf
                        DoEvents
                        wsock_Close Index
                        Exit Sub
                    End If
                    Unauthorized Index
                    Exit Sub
                End If
            Next iTMP
                    Clients(Index).Attempts = Clients(Index).Attempts + 1
                    If Clients(Index).Attempts = 3 Then
                        wsock(Index).SendData vbCrLf & vbCrLf & " Too many invalid username/passwords. Closing..." & vbCrLf
                        DoEvents
                        wsock_Close Index
                        Exit Sub
                    End If
                    Unauthorized Index
                    Exit Sub
        Case "COMM"
            txtLog.Text = txtLog.Text & Clients(Index).Name & " - " & cmd
            Clients(Index).Action = "NONE"
            Comm = LCase$(GetPiece(cmd, " ", 1))
            wsock(Index).SendData vbCrLf
            Select Case Comm
               
               'Case "yourcommand"
                   
                   '************************************************
                   'You can add your commands
                   'Please view the other commands first before doing any tests ;-)
                    '************************************************
                
                Case "info"
                    wsock(Index).SendData vbCrLf
                    wsock(Index).SendData " -===============================================-" & vbCrLf
                    wsock(Index).SendData " -=                  Telnet Server              =-" & vbCrLf
                    wsock(Index).SendData " -=              Version " & programVer & "                  =-" & vbCrLf
                    wsock(Index).SendData " -===============================================-" & vbCrLf
                    wsock(Index).SendData " -=   Programmed/Developed by Vasilis Sagonas   =-" & vbCrLf
                    wsock(Index).SendData " -=               Vasilakis on IRC              =-" & vbCrLf
                    wsock(Index).SendData " -=          (C) Copyright, 1999 - 2001         =-" & vbCrLf
                    wsock(Index).SendData " -===============================================-" & vbCrLf
                    wsock(Index).SendData " -=           Email: vsag@forthnet.gr           =-" & vbCrLf
                    wsock(Index).SendData " -===============================================-" & vbCrLf & vbCrLf
                    wsock(Index).SendData "    ...That's all folks!!!" & vbCrLf & vbCrLf
                Case "time"
                    wsock(Index).SendData vbCrLf & " -= The time is : " & Time$ & vbCrLf & vbCrLf
                Case "date"
                    wsock(Index).SendData vbCrLf & " -= The date is : " & Format$(Date$, "dd/mm/yy") & vbCrLf & vbCrLf
                Case "help"
                    If GetPiece(cmd, " ", 2) = "general" Then
                        wsock(Index).SendData vbCrLf
                        wsock(Index).SendData " -= General Commands =-" & vbCrLf & vbCrLf
                        wsock(Index).SendData " info " & vbTab & vbTab & vbTab & vbTab & "[about the author" & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " cls  " & vbTab & vbTab & vbTab & vbTab & "[clears screen  " & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " help " & vbTab & vbTab & vbTab & vbTab & "[shows this help" & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " time" & vbTab & vbTab & vbTab & vbTab & "[shows the time  " & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " date" & vbTab & vbTab & vbTab & vbTab & "[shows the date  " & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " who " & vbTab & vbTab & vbTab & vbTab & "[shows connect people" & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " talk <message> " & vbTab & vbTab & "[sends global message" & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " talkto <name> <message> " & vbTab & "[sends message to someone" & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " winver" & vbTab & vbTab & vbTab & vbTab & "[windows version          " & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " userlist " & vbTab & vbTab & vbTab & "[list users" & vbTab & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " logout " & vbTab & vbTab & vbTab & "[logs out for login" & vbTab & vbTab & "]" & vbCrLf
                        wsock(Index).SendData " quit, xit, exit " & vbTab & vbTab & "[quit telnet" & vbTab & vbTab & vbTab & "]" & vbCrLf & vbCrLf
                   
                   'ElseIf GetPiece(cmd, " ", 2) = "yourcategory" Then
                        '************************************************
                        'your help for your category
                        '************************************************
                    Else
                        wsock(Index).SendData vbCrLf
                        wsock(Index).SendData " -= HELP =-" & vbCrLf & vbCrLf
                        wsock(Index).SendData " Syntax: help <general>" & vbCrLf & vbCrLf ' add here your category
                        wsock(Index).SendData "   help general - General Telnet Commands" & vbCrLf
                        wsock(Index).SendData vbCrLf
                    End If
                Case "winver"
                    wsock(Index).SendData vbCrLf & " -= Server is running on " & WinVersion & vbCrLf & vbCrLf
                Case "cls"
                    ScreenCls Index: DoEvents
                Case "adduser"
                    rUser = GetPiece(cmd, " ", 2)
                    rPass = GetPiece(cmd, " ", 3)
                    If rPass = "" Then wsock(Index).SendData vbCrLf & " -= Missing argument -- adduser <username> <password>." & vbCrLf & vbCrLf: GoTo Continue
                    For I = 0 To lstUser.ListCount - 1
                        If LCase$(lstUser.List(I)) = LCase$(rUser) Then
                            If LCase$(rUser) = LCase$(Clients(Index).Name) Then
                                wsock(Index).SendData vbCrLf & "* Huh? I think this is you ;-)" & vbCrLf & vbCrLf
                            Else
                                wsock(Index).SendData vbCrLf & "* User exists." & vbCrLf & vbCrLf
                            End If
                            DoEvents
                            txtUser.Text = ""
                            txtPass.Text = ""
                            GoTo Continue
                        End If
                    Next I
                    txtUser.Text = rUser
                    txtPass.Text = rPass
                    cmdAdduser_Click
                    SaveUsers
                    wsock(Index).SendData vbCrLf & "* User has been added." & vbCrLf & vbCrLf
                Case "setpass"
                    rPass = GetPiece(cmd, " ", 2)
                    If rPass = "" Then wsock(Index).SendData vbCrLf & " -= Missing argument -- setpass <password>." & vbCrLf & vbCrLf: GoTo Continue
                    For I = 0 To lstUser.ListCount - 1
                        If LCase$(lstUser.List(I)) = LCase$(rUser) Then
                            If LCase$(rUser) = LCase$(Clients(Index).Login) Then
                                wsock(Index).SendData vbCrLf & "* Huh? I think this is you ;-)" & vbCrLf & vbCrLf
                            Else
                                wsock(Index).SendData vbCrLf & "* User exists." & vbCrLf & vbCrLf
                            End If
                            DoEvents
                            txtUser.Text = ""
                            txtPass.Text = ""
                            GoTo Continue
                        End If
                    Next I
                    txtUser.Text = rUser
                    txtPass.Text = rPass
                    cmdAdduser_Click
                    SaveUsers
                    wsock(Index).SendData vbCrLf & "* User has been added." & vbCrLf & vbCrLf
                Case "deluser"
                    rUser = GetPiece(cmd, " ", 2)
                    If rUser = "" Then wsock(Index).SendData vbCrLf & " -= Missing argument -- adduser <username> <password>." & vbCrLf & vbCrLf: GoTo Continue
                    For I = 0 To lstUser.ListCount - 1
                        If LCase$(lstUser.List(I)) = LCase$(rUser) Then
                            lstUser.ListIndex = I
                            cmdDelete_Click
                            SaveUsers
                            DoEvents
                            wsock(Index).SendData vbCrLf & "* User has been removed." & vbCrLf & vbCrLf
                            GoTo Continue
                        End If
                    Next I
                    wsock(Index).SendData vbCrLf & "* User does not exist." & vbCrLf & vbCrLf
                    DoEvents
                Case "userlist"
                    wsock(Index).SendData vbCrLf & " -= User Database =-" & vbCrLf & vbCrLf
                    For iTMP = 0 To lstUser.ListCount - 1
                                DoEvents
                                If lstUser.Selected(iTMP) = True Then
                                    iR = vbTab & "[Admin]"
                                Else
                                    iR = ""
                                End If
                                
                                wsock(Index).SendData " " & lstUser.List(iTMP) & iR & vbCrLf
                    Next iTMP
                    wsock(Index).SendData vbCrLf
                Case "who"
                    wsock(Index).SendData vbCrLf & " -= Line" & vbTab & "Username =- " & vbCrLf
                    wsock(Index).SendData " -========" & vbTab & "=========================-" & vbCrLf
                    DoEvents
                    For Each ws In wsock
                        If ws.Index > 0 And Clients(ws.Index).Login = True Then
                                wsIP = ws.RemoteHostIP
                                If wsIP = "127.0.0.1" Then wsIP = "server"
                                wsock(Index).SendData "   " & ws.Index & vbTab & vbTab & Clients(ws.Index).Name & "@" & wsIP & vbCrLf
                                DoEvents
                            DoEvents
                        End If
                    Next ws
                    wsock(Index).SendData vbCrLf
                    DoEvents
                Case "quit", "exit", "xit"
                    wsock(Index).SendData vbCrLf
                    DoEvents
                    wsock(Index).SendData " -= Logging off..." & vbCrLf & vbCrLf
                    DoEvents
                    wsock_Close Index
                Case "logout"
                    wsock(Index).SendData vbCrLf & " -= Logging out..." & vbCrLf & vbCrLf
                    DoEvents
                    Clients(Index).Action = "user"
                    Clients(Index).Name = ""
                    Clients(Index).Login = False
                    wsock(Index).Tag = ""
                    Welcome Index
                    Exit Sub
                Case "talkto"
                    If GetPiece(cmd, " ", 3) = "" Then wsock(Index).SendData vbCrLf & " -= Parameter: talk <name> <text>." & vbCrLf & vbCrLf: GoTo Continue
                    DoEvents
                    For Each ws In wsock
                        If ws.Index > 0 And Clients(ws.Index).Login = True And LCase$(Clients(ws.Index).Name) = LCase(GetPiece(cmd, " ", 2)) And ws.Index <> Index Then
                            If Clients(ws.Index).Action = "COMM" Then
                                ws.SendData vbCrLf & vbCrLf & " -= Message from: " & Clients(Index).Name & " - " & Right(cmd, Len(cmd) - 8 - Len(GetPiece(cmd, " ", 2))) & vbCrLf & vbCrLf
                                DoEvents
                                SendName ws.Index
                                DoEvents
                                ws.SendData ws.Tag
                            End If
                            DoEvents
                        End If
                    Next ws
                    DoEvents
                Case "talk"
                    If GetPiece(cmd, " ", 2) = "" Then wsock(Index).SendData vbCrLf & " -= Parameter: talk <text>." & vbCrLf & vbCrLf: GoTo Continue
                    DoEvents
                    For Each ws In wsock
                        If ws.Index > 0 And Clients(ws.Index).Login = True And ws.Index <> Index Then
                            If Clients(ws.Index).Action = "COMM" Then
                                ws.SendData vbCrLf & vbCrLf & " -= " & Clients(Index).Name & " - " & Right(cmd, Len(cmd) - 5) & vbCrLf & vbCrLf
                                DoEvents
                                SendName ws.Index
                                DoEvents
                                ws.SendData ws.Tag
                            End If
                            DoEvents
                        End If
                    Next ws
                    DoEvents
                Case ""
                    DoEvents
                Case Else
WrongCMD:
                    wsock(Index).SendData vbCrLf
                    DoEvents
                    wsock(Index).SendData " -= Command not understood: " & Comm & vbCrLf & vbCrLf
                    DoEvents
            End Select
Continue:
            '************************************************************************************************
            'SendName sends to current user the 'User>' prompt for command
            '************************************************************************************************
            SendName Index
            Clients(Index).Action = "COMM"
            
        Case Else
            wsock(Index).SendData vbCrLf
    End Select
    Exit Sub
End If

'************************************************************************************************************
'These are the packet filter and command recorder! Please do not harm it :-)
'This algorithm can be used for several purposes like an internet tool because
'it decode the winsock packets with vbCrLf (Enter + LineFeed) characters
'and it is fast too! You can use it with any application
'************************************************************************************************************

If Clients(Index).Action = "NONE" Then Exit Sub
rTEMP = vtData
I = 0
Do
    I = I + 1
    iTEMP = Mid(rTEMP, I, 1)
    rTEMPV = wsock(Index).Tag
    vtData = Right(rTEMP, I)
    If Mid(rTEMP, I, 2) = vbCrLf Then
            vtData = Right(rTEMP, I + 1)
            GoTo cmd
        ElseIf Mid(rTEMP, I, 3) = "[A" Or Mid(rTEMP, I, 3) = "OA" Then
            'Up key pressed! Send the last command the user entered.
            If Clients(Index).Action = "COMM" Then
                If wsock(Index).Tag <> "" And Clients(Index).Action = "COMM" Then
                    wsock(Index).SendData String$(Len(wsock(Index).Tag), Chr$(8)) & String$(Len(wsock(Index).Tag), " ") & String$(Len(wsock(Index).Tag), Chr$(8))
                    DoEvents
                End If
                Clients(Index).NextCommand = wsock(Index).Tag
                wsock(Index).Tag = Clients(Index).LastCommand
                wsock(Index).SendData Clients(Index).LastCommand
                DoEvents
            End If
            vtData = Right(rTEMP, I + 2)
            I = I + 2
        ElseIf iTEMP = Chr$(8) Then
            'Backspace pressed! Remove the last character entered
            If Len(wsock(Index).Tag) <> 0 Then
                wsock(Index).Tag = Left(wsock(Index).Tag, Len(wsock(Index).Tag) - 1)
                If Clients(Index).Action <> "PASS" Then
                    wsock(Index).SendData Chr$(8) & " " & Chr$(8)
                    DoEvents
                End If
            End If
        ElseIf Mid(rTEMP, I, 3) = "[B" Or Mid(rTEMP, I, 3) = "OB" Then
            'Down key pressed! Send the command the user was entering when he pressed the UP key
            'or remove anything he is currently writing!
            If Clients(Index).Action = "COMM" Then
                If wsock(Index).Tag <> "" And Clients(Index).Action = "COMM" Then
                    If wsock(Index).Tag <> "" Then Clients(Index).LastCommand = wsock(Index).Tag
                    wsock(Index).SendData String$(Len(wsock(Index).Tag), Chr$(8)) & String$(Len(wsock(Index).Tag), " ") & String$(Len(wsock(Index).Tag), Chr$(8))
                    DoEvents
                    wsock(Index).Tag = Clients(Index).NextCommand
                    wsock(Index).SendData Clients(Index).NextCommand
                    DoEvents
                End If
            End If
            vtData = Right(rTEMP, I + 2)
            I = I + 2
        ElseIf Mid(rTEMP, I, 3) = "[C" Or Mid(rTEMP, I, 3) = "OC" Then
            'Right key pressed! Remove these characters because we don't need them.
            vtData = Right(rTEMP, I + 2)
            I = I + 2
        ElseIf Mid(rTEMP, I, 3) = "[D" Or Mid(rTEMP, I, 3) = "OD" Then
            'Left key pressed! Remove these characters because we don't need them.
            vtData = Right(rTEMP, I + 2)
            I = I + 2
        Else
            If Clients(Index).Action = "pass" Then
                'If he is entering the password then hide it with '*' character
                wsock(Index).SendData "*"
            Else
                'If user is not giving the password then send the character back.
                wsock(Index).SendData iTEMP
            End If
            wsock(Index).Tag = wsock(Index).Tag & iTEMP 'Add to tag the character i filtered from the packets
            DoEvents
    End If
Loop Until I >= Len(rTEMP)
DoEvents
End Sub

'**************************************************************
'Windows version in an easier way to read ;-)
'**************************************************************'
Function WinVersion() As String
        Dim myVer As MYVERSION
        myVer = WindowsVersion()
        
        If myVer.lMajorVersion = 4 Then
            If myVer.lExtraInfo = VER_PLATFORM_WIN32_NT Then
                strTmp = "Windows NT v"
            ElseIf myVer.lExtraInfo = VER_PLATFORM_WIN32_WINDOWS Then
                strTmp = "Windows 95 v"
            End If
        ElseIf myVer.lMajorVersion = 5 Then
            strTmp = "Windows 2000 - NT v"
        Else
            strTmp = "Windows v"
        End If
            
        WinVersion = strTmp & myVer.lMajorVersion & "." & myVer.lMinorVersion

End Function


Sub SendName(Index)
wsock(Index).SendData Clients(Index).Name & "> "
DoEvents
End Sub


Sub Unauthorized(who As Integer)
On Error Resume Next
wsock(who).SendData vbCrLf
wsock(who).SendData "  Invalid Username/Password." & vbCrLf
wsock(who).SendData vbCrLf
wsock(who).SendData " Username: "
Clients(who).Action = "user"
End Sub


Function GetPiece(from As String, delim As String, Index) As String
    Dim temp$
    Dim Count
    Dim Where
    '
    temp$ = from & delim
    Where = InStr(temp$, delim)
    Count = 0
    Do While (Where > 0)
        Count = Count + 1
        If (Count = Index) Then
            GetPiece = Left$(temp$, Where - 1)
            Exit Function
        End If
        temp$ = Right$(temp$, Len(temp$) - Where)
        Where = InStr(temp$, delim)
    DoEvents
    Loop
    If (Count = 0) Then
        GetPiece = from
    Else
        GetPiece = ""
    End If
End Function

