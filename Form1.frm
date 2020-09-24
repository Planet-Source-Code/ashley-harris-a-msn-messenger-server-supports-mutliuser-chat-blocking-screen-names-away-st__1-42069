VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "MSN server"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton showadmin 
      Caption         =   "Admin"
      Height          =   375
      Left            =   6195
      TabIndex        =   2
      Top             =   4530
      Width           =   1500
   End
   Begin VB.TextBox Log 
      Height          =   3210
      Index           =   0
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3345
      Top             =   2805
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   2790
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1863
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0006
      Height          =   1605
      Left            =   2070
      TabIndex        =   1
      Top             =   3360
      Width           =   3540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim md5 As New md5

Public userstates As New Dictionary
Public listversions As New Dictionary
Public senddone As New Dictionary

Public switchboards As New Collection

Private Sub Form_Load()
    ws(0).Close
    ws(0).LocalPort = 1863
    ws(0).Listen
    initlists
    Debug.Print ws(0).State
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'tell all connected clients that were going down for mantinence
    For Each a In userstates.Keys
        sendmsgtousr a, "OUT SSD", True
    Next
    End
End Sub

Private Sub Form_Resize()
    refreshchatlocations
End Sub

Private Sub showadmin_Click()
    admin.Show
End Sub

Private Sub Timer1_Timer()
    Dim dic As Dictionary

    'check to see if a user has unexpectidly gone offline.
    'for some reason, ws_error doesn't fire. which is more often then not
    For a = 0 To ws.UBound
        If ws(a).Tag <> "CLOSED" And ws(a).State <> 7 And ws(a).Tag <> "" Then
            ws(a).Close
            
            Set dic = pulllist("rl", CStr(ws(a).Tag))
            For b = 0 To UBound(dic.Keys)
                sendmsgtousr dic.Keys(b), "FLN " & ws(a).Tag
            Next b
            ws(a).Tag = "CLOSED"
            Log(a).Visible = False
        End If
    Next a
    If ws(0).State <> 2 Then ws(0).Close: ws(0).Listen
    
    For a = 0 To UBound(userstates.Keys)
        For b = 0 To ws.UBound
            If ws(b).Tag = userstates.Keys(a) And ws(b).State = 7 Then GoTo fine
        Next b
        userstates(userstates.Keys(a)) = ""
fine:
    Next a
End Sub

Private Sub ws_Close(Index As Integer)
    'same as above, catches some other times
    On Error Resume Next
    ws(Index).Close
    Set dic = pulllist("rl", CStr(ws(Index).Tag))
    For a = 0 To UBound(dic.Keys)
        sendmsgtousr dic.Keys(a), "FLN " & usr
    Next a
    'ws(Index).Tag = ""
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'allow new users to signin to the notification system.
    
    For a = 1 To ws.UBound
        If ws(a).State <> 7 Then Exit For
    Next a
    If a = ws.Count Then
        Load ws(a)
        Load Log(a)
        Log(a).Left = Log(a).Width * ((a - 1) Mod 3)
        Log(a).Top = Log(a).Height * ((a - 1) \ 3)
        Log(a).Visible = True
    End If
    Log(a).Text = ""
    ws(a).Close
    ws(a).Accept requestID
    ws(a).Tag = ""
    addtolog a, ws(a).RemoteHostIP & ":" & ws(a).RemotePort
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim I As String, dic As Dictionary, dic2 As Dictionary
    ws(Index).PeekData I
    
    If Right(I, 1) = vbLf And InStr(1, I, vbCrLf) = 0 Then
        'trillain forgets that a newline is crlf (sometimes)
        I = Replace(I, vbLf, vbCrLf)
    End If
        
    If InStr(1, I, vbCrLf) = 0 Then Exit Sub
    ws(Index).GetData I, , InStr(1, I, vbCrLf) + 1
    
    If Len(I) <> bytesTotal Then ws_DataArrival Index, bytesTotal - Len(I)
    
    If Right(I, 1) = vbLf And InStr(1, I, vbCrLf) = 0 Then
        'trillain forgets that a newline is crlf (sometimes)
        I = Replace(I, vbLf, vbCrLf)
    End If
    
    'ok, you've pulled the topline, now, analise it
    
    If I = "" Or I = vbCrLf Then Exit Sub
    I = Replace(I, vbCr, "")
    I = Replace(I, vbLf, "")
    
    addtolog Index, "C: " & I
    
    cl = Split(I)
    cmd = cl(0)
    On Error Resume Next
    PID = cl(1)
    usr = ws(Index).Tag
    On Error GoTo 0
    Select Case UCase(cmd)
    Case "VER"
        'handshaking, work out what version to chat with
        back = "VER " & PID & " MSNP4"
    Case "INF"
        'Login information handshake
        back = "INF " & PID & " MD5"
    Case "USR"
        'do the logon script
        If LCase(cl(3)) = "i" Then
            'usrname specified respond with hash for shared secret login
            
            If Not userexists(cl(4)) Then
                'bad username
                back = "911"
            Else
                'good username
                back = "USR " & PID & " MD5 S " & DateDiff("h", #1/1/2000#, Now)
                
                'if there logged in elsewhere, log them out
                sendmsgtousr cl(4), "OUT OTH", True
                
                ws(Index).Tag = cl(4)
            End If
        Else
            'password specified
            If LCase(md5.DigestStrToHexStr(DateDiff("h", #1/1/2000#, Now) & getpassword(usr))) = LCase(cl(4)) Then
                'login ok
                back = "USR " & PID & " OK " & usr & " " & getscreenname(usr)
            Else
                'bad login
                back = "911 " & PID
            End If
        End If
    Case "SYN"
        'Client has requested syncronisation of the contact lists
        
        userstates(usr) = ""
        back = "SYN " & PID & " 999"
        
        'send the users setting back to the client, They can't change them in this server.
        'to tell you the truth, I don't know what they are, I suspect handling of blocked
        'or new users. but, These work for both MSN and trillian, so, I don't care.
        
        'I note that setting the first one to N (I've observed it being another possible
        'setting for the GTC command), adding contacts doesn't work. It apears to be a
        ' 'lock' of the contact lists of sorts. (don't ask me, i barely know the protocol)
        
        back = back & vbCrLf & "GTC " & PID & " 999 A"
        back = back & vbCrLf & "BLP " & PID & " 999 AL"
        
        'Send the users contact list back to the client
        Set dic = pulllist("fl", CStr(usr))
        For a = 0 To UBound(dic.Keys)
            back = back & vbCrLf & "LST " & PID & " FL 999 " & a + 1 & " " & UBound(dic.Keys) + 1 & " " & dic.Keys(a) & " " & dic(dic.Keys(a))
        Next a
        If UBound(dic.Keys) = -1 Then back = back & vbCrLf & "LST " & PID & " FL 999 0 0"
        
        'Send the users 'allow list' back to the client
        Set dic = pulllist("al", CStr(usr))
        For a = 0 To UBound(dic.Keys)
            back = back & vbCrLf & "LST " & PID & " AL 999 " & a + 1 & " " & UBound(dic.Keys) + 1 & " " & dic.Keys(a) & " " & dic(dic.Keys(a))
        Next a
        If UBound(dic.Keys) = -1 Then back = back & vbCrLf & "LST " & PID & " AL 999 0 0"
        
        'send the users 'blocked list' back to the client
        Set dic = pulllist("bl", CStr(usr))
        For a = 0 To UBound(dic.Keys)
            back = back & vbCrLf & "LST " & PID & " BL 999 " & a + 1 & " " & UBound(dic.Keys) + 1 & " " & dic.Keys(a) & " " & dic(dic.Keys(a))
        Next a
        If UBound(dic.Keys) = -1 Then back = back & vbCrLf & "LST " & PID & " BL 999 0 0"

        'send the users 'return list' back to the client. Basically a list of
        'everybody who has you on their list
        Set dic = pulllist("rl", CStr(usr))
        For a = 0 To UBound(dic.Keys)
            back = back & vbCrLf & "LST " & PID & " RL 999 " & a + 1 & " " & UBound(dic.Keys) + 1 & " " & dic.Keys(a) & " " & dic(dic.Keys(a))
        Next a
        If UBound(dic.Keys) = -1 Then back = back & vbCrLf & "LST " & PID & " RL 999 0 0"

        'msg = _
        "MIME-Version: 1.0" & vbCrLf & _
        "Content-Type: text/x-msmsgspro file; charset=UT" & vbCrLf & _
        "LoginTime: " & Fix(Now * 86400) & vbCrLf & _
        "EmailEnabled: 1" & vbCrLf & _
        "MemberIdHigh: 84736" & vbCrLf & _
        "MemberIdLow: -1434729391" & vbCrLf & _
        "lang _preference: 103" & vbCrLf & _
        "preferredEmail: " & usr & vbCrLf & _
        "country: AU" & vbCrLf & _
        "PostalCode: " & vbCrLf & _
        "Gender: M" & vbCrLf & _
        "Kid:0" & vbCrLf & _
        "Age: 22" & vbCrLf & _
        "sid: 517" & vbCrLf & _
        "kv: 2" & vbCrLf & _
        "MSPAuth: 2AAAAAAAADU0p4uxxxJtDJozJSlUTS0i7YpwnC9PUHRv56YKxxxCTWmg$$" & vbCrLf & vbCrLf
        'back = back & vbCrLf & "MSG Hotmail Hotmail " & Len(msg) & vbCrLf & msg
        
        back = Replace(back, "999", cl(2) + 1)
        
        listversions(usr) = cl(2) + 1
        
    Case "CHG"
        'change online state
        
        back = "CHG " & PID & " " & cl(2)
        
        If userstates(usr) = "" Then
            'ok, here, after their first changestate command, they recieve
            'A list of all thier contacts who are online. Doesn't make sense to me.
            'I think they should get the online list after they get their reverse list.
            'but, I didn't write the protocol.
            Set dic = pulllist("fl", CStr(usr))
            For a = 0 To UBound(dic.Keys)
                If userstates(dic.Keys(a)) <> "" And userstates(dic.Keys(a)) <> "HDN" Then
                    Set dic2 = pulllist("bl", CStr(dic.Keys(a)))
                    If dic2(usr) = "" Then back = back & vbCrLf & "ILN " & PID & " " & userstates(dic.Keys(a)) & " " & dic.Keys(a) & " " & getscreenname(dic.Keys(a))
                End If
            Next a
        End If
        
        If cl(2) = "FLN" Then cl(2) = ""
        
        changestate usr, cl(2)

        
    Case "CVR"
        'is there a new version of MSN out? only the microsft version of MSN asks this.
        'I have no idea what the individual sections mean. I know the last 2 are the
        'download url and the information url. but, I have no idea about the others.
        back = "CVR " & PID & " 4.5.0127 4.5.0127 1.0.0863 http://download.microsoft.com/download/msnmessenger/install/4.5/win98me/en-us/mmssetup.exe http://messenger.microsoft.com"
    
    Case "OUT"
        'a user is logging out the correct way
        back = "OUT"
        
        Set dic = pulllist("rl", CStr(usr))
        For a = 0 To UBound(dic.Keys)
            sendmsgtousr dic.Keys(a), "FLN " & usr
        Next a
        
        ws(Index).Tag = "LOGOFF"
        
    Case "ADD"
        'add a user to your list. not nesicarily your contact list, it can be
        'your forward list (contact list), block list, allow list, or reverse list
        If Not userexists(cl(3)) Then
            back = "205"
            GoTo sendback
        End If
        addusertolist usr, cl(2), cl(3)
        
        If LCase(cl(2)) = "fl" Then
            'if a adds b to their contact list, then let b know
            'then, send the 'online' notifyer to them
            
            addusertolist cl(3), "al", usr
            'addusertolist cl(3), "fl", usr
            'listversions(cl(3)) = listversions(cl(3)) + 1
            sendmsgtousr cl(3), "ADD 0 RL " & listversions(cl(3)) & " " & usr & " " & getscreenname(usr)
            'listversions(cl(3)) = listversions(cl(3)) + 1
            'sendmsgtousr cl(3), "ADD 0 FL " & listversions(cl(3)) & " " & usr & " " & getscreenname(usr)
            
            'sendmsgtousr cl(3), "ADD 0 AL 1 " & usr & " " & getscreenname(usr)
            
            listversions(usr) = listversions(usr) + 1
            
            'back = back & vbCrLf & "ADD RL " & listversions(usr) & " " & cl(3) & " " & getscreenname(cl(3))
            
            If userstates(cl(3)) <> "HDN" And userstates(cl(3)) <> "" And userstates(cl(3)) <> "FLN" Then
                back = back & vbCrLf & "NLN " & userstates(cl(3)) & " " & cl(3) & " " & getscreenname(cl(3))
            End If
            
            If userstates(usr) <> "HDN" And userstates(usr) <> "" And userstates(usr) <> "FLN" Then
                back = back & vbCrLf & "NLN " & userstates(usr) & " " & usr & " " & getscreenname(usr)
            End If
            
        End If
        
        back = "ADD " & PID & " " & cl(2) & " 1 " & cl(3) & " " & getscreenname(cl(3))
        
        If LCase(cl(2)) = "bl" Then
            'usr a has just blocked usr b, so, tell b that a's offline
        
            sendmsgtousr cl(3), "FLN " & usr
            
        End If
        
        changestate usr, userstates(usr)
        changestate cl(3), userstates(cl(3))
        
    Case "REM"
        'remove a user from one of the 4 lists alocated to each user.
        
        removeuserfromlist usr, cl(2), cl(3)
        back = "REM " & PID & " " & cl(2) & " 1 " & cl(3)
        
        If LCase(cl(2)) = "bl" Then
            'usr a has just unblocked usr b, so, tell b that a's yadda yadda
            sendmsgtousr cl(3), "NLN " & userstates(usr) & " " & usr & " " & getscreenname(usr)
        End If
    Case "REA"
    
        If LCase(cl(2)) = LCase(usr) Then
            'change my screename.
            setscreename usr, CStr(cl(3))
            back = "REA " & PID & " OK " & usr & " " & cl(3)
            
            'there was a bug, when someone in trillian signs on, it's set's
            'their screename, then sets them as online. This would cause problems
            'as other clients would be notified of the name change, but it
            'would send "" as the client state, which would crash other copies
            'of trillian, and MSN would reject the line. so, easy fix:
            If userstates(usr) = "" Then GoTo sendback
            
            Set dic = pulllist("rl", CStr(usr))
            Set dic2 = pulllist("bl", CStr(usr))
            
            'notify all those on your list about the change
            For a = 0 To UBound(dic.Keys)
                If dic2(dic.Keys(a)) = "" Then sendmsgtousr dic.Keys(a), "NLN " & userstates(usr) & " " & usr & " " & getscreenname(usr)
            Next a
        Else
            'this ones weird, but I think it means "confirm that user cl(2)'s name is cl(3)" ?????? I dunno. it's a microsoft creation.
            back = "REA " & PID & " OK " & cl(2) & " " & getscreenname(cl(2))
        End If
    Case "XFR"
        'The user wants to be refered to another server, I think this is allways SB
        'for switchboard (for chat sessions) but, I dunno. (Voice chats? Files? I dunno)
        Dim sw As New sb
        back = "XFR " & PID & " SB " & ws(0).LocalIP & ":" & sw.setup & " CKI " & CKInewcookie
        sw.Top = Me.Top + Me.Height + sw.Height * ((switchboards.Count - 1) \ 3)
        sw.Left = Me.Left + sw.Width * ((switchboards.Count) Mod 3)
        
        sw.Visible = Me.WindowState <> vbMinimized
        
        switchboards.Add sw
        
    Case Else
        'something else was recieved, respond with a failure message
        back = "200 " & PID
    End Select
    
    
sendback:
    'now, send it back to the user
    back = back & vbCrLf
    For Each a In Split(back, vbCrLf)
        If a <> "" Then
            addtolog Index, "S: " & a
            ws(Index).SendData a & vbCrLf
        End If
    Next
    If back = "OUT" & vbCrLf Then ws(Index).Tag = "LOGOFF"
End Sub

Public Sub sendmsgtousr(usrname, msg, Optional logoff As Boolean)
    For a = 0 To ws.UBound
        If ws(a).Tag = usrname Then
            If ws(a).State = 7 Then
                If logoff Then ws(a).Tag = "LOGOFF"
                ws(a).SendData msg & vbCrLf
                waitforsend a
                addtolog a, "S: " & msg
            End If
        End If
    Next
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'On Error Resume Next
    ws(Index).Close
    Dim dic As Dictionary
    Set dic = pulllist("al", CStr(ws(Index).Tag))
    For a = 0 To UBound(dic.Keys)
        sendmsgtousr dic.Keys(a), "FLN " & ws(Index).Tag
    Next a
    'ws(Index).Tag = ""
End Sub

Private Sub ws_SendComplete(Index As Integer)
    If ws(Index).Tag = "LOGOFF" Then ws(Index).Close
    senddone(Index) = True
End Sub

Public Sub waitforsend(I)
    senddone(I) = False
    t = Timer + 5
    While senddone(I) = False And Timer < t
        DoEvents
    Wend
End Sub

Public Sub addtolog(n, t)
    Log(n).Visible = True
    Log(n).Text = Log(n).Text & t & vbCrLf
    Log(n).SelStart = Len(Log(n).Text) + 1
End Sub

Public Sub changestate(usr, newstate)
    'tell the entire world our state, with the exception of those who we have blocked
    Dim dic As Dictionary, dic2 As Dictionary
    
    Set dic = pulllist("al", CStr(usr))
    Set dic2 = pulllist("bl", CStr(usr))
    If newstate = "FLN" Or newstate = "HDN" Then
        'when the user goes offline, or pretends to go offline,
        'tell the world that they've gone offline.
        For a = 0 To UBound(dic.Keys)
            If Not dic2.Exists(dic.Keys(a)) Then sendmsgtousr dic.Keys(a), "FLN " & usr
        Next a
    Else
        'when they come online, tell the world that their now online.
        'this will also occur when the user changes state.
        For a = 0 To UBound(dic.Keys)
            If Not dic2.Exists(dic.Keys(a)) Then sendmsgtousr dic.Keys(a), "NLN " & newstate & " " & usr & " " & getscreenname(usr)
        Next a
    End If
    userstates(usr) = newstate
End Sub

Public Sub refreshchatlocations()
    Dim sw As sb
    For a = 1 To switchboards.Count
        Set sw = switchboards(a)
        sw.Top = Me.Top + Me.Height + sw.Height * ((a - 1) \ 3)
        sw.Left = Me.Left + sw.Width * ((a - 1) Mod 3)
        sw.Visible = WindowState <> vbMinimized
    Next a
End Sub
