VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form sb 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Switchboard Server/Chat Session"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2715
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   465
      Top             =   285
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   1380
      Top             =   870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "sb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sessionid As Long

Private senddone As New Dictionary

Public lastmsg As Long

Public sessiondata As New Dictionary

Private loglifename As String
Dim ts As TextStream

Public Sub sendtouser(usr, cmd)
    Debug.Print "S-" & usr & ": " & cmd
    For a = 1 To ws.UBound
        DoEvents
        If ws(a).Tag = usr And ws(a).State = 7 Then
            ws(a).SendData cmd & vbCrLf
            waitforsend a
            DoEvents
        End If
    Next a
End Sub

Public Sub sendtoall(cmd)
    DoEvents
    Debug.Print "S-ALL:" & cmd
    For a = 1 To ws.UBound
        If ws(a).Tag <> "NEWBIE" And ws(a).State = 7 Then
            ws(a).SendData cmd & vbCrLf
            waitforsend a
            DoEvents
        End If
    Next a
End Sub

Public Sub sendtoallbut(excludeusr, cmd)
    DoEvents
    Debug.Print "S-ALLBUTONE:" & cmd
    For a = 1 To ws.UBound
        If ws(a).Tag <> "NEWBIE" And ws(a).State = 7 And ws(a).Tag <> excludeusr Then
            ws(a).SendData cmd & vbCrLf
            waitforsend a
            DoEvents
        End If
    Next a
End Sub

Private Sub waitforsend(I)
    senddone(I) = False
    t = Timer + 5
    While senddone(I) = False And ws(I).State = 7 And t > Timer
        DoEvents
    Wend
    
End Sub

Public Function setup() As Long
    'sets up a switchboard server, returns the port number
    
    ws(0).Close
    ws(0).LocalPort = 32768 + Int(Rnd() * 32768)
    ws(0).Listen
    
    setup = ws(0).LocalPort
    sessionid = Int(Rnd() * 100000000#)
    
    loglifename = fso.BuildPath(App.Path, CDbl(Time) & ".txt")
    Set ts = fso.OpenTextFile(loglifename, ForWriting, True)
End Function

Private Sub Timer1_Timer()
    For a = 1 To ws.UBound
        If ws(a).State <> 7 And ws(a).Tag <> "" And ws(a).Tag <> "NEWBIE" Then
            DoEvents
            sendtoall "BYE " & ws(a).Tag
            ws(a).Close
            ws(a).Tag = ""
        End If
        If ws(a).State = 7 Then inchat = inchat + 1
    Next
    Me.Caption = "chat: " & inchat & " participants"
    If inchat = 0 Or (inchat = 1 And lastmsg < Timer - 100) Or (inchat = 2 And lastmsg < Timer - 300) Then
        
        For a = Form1.switchboards.Count To 1 Step -1
            If Form1.switchboards(a) Is Me Then
                Form1.switchboards.Remove a
                Me.Visible = False
                Unload Me
            End If
        Next a
        Form1.refreshchatlocations
    End If
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    For a = 1 To ws.UBound
        If ws(a).State <> 7 Then Exit For
    Next a
    If a = ws.Count Then Load ws(a)
    ws(a).Close
    ws(a).Accept requestID
    ws(a).Tag = "NEWBIE"
    Debug.Print "NEW CONNECTION " & a
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim z As String
    ws(Index).GetData z
    If Right(z, 1) = vbLf Then z = Mid(z, 1, Len(z) - 1)
    If Right(z, 1) = vbCr Then z = Mid(z, 1, Len(z) - 1)
    cl = Split(z, " ")
    On Error Resume Next
    PID = cl(1)
    usr = ws(Index).Tag
    On Error GoTo 0
    
    Debug.Print "C-" & usr & ": " & z
    lastmsg = Timer
    DoEvents
    Select Case UCase(cl(0))
    Case "USR"
        'A user is signing in to a new chat room
        If CKIcookieisvalid(cl(3)) Then
            ws(Index).Tag = cl(2)
            sendtouser cl(2), "USR " & PID & " OK " & cl(2) & " " & getscreenname(cl(2))
        Else
            Debug.Print "S-NEWBIE: 200 " & PID & vbCrLf; ""
            ws(Index).SendData "200 " & PID & vbCrLf
        End If
        
    Case "CAL"
        'a user is 'calling' another user into the chat room]
        
        'call them
        Form1.sendmsgtousr cl(2), "RNG " & sessionid & " " & ws(0).LocalIP & ":" & ws(0).LocalPort & " CKI " & CKInewcookie & " " & usr & " " & getscreenname(usr)
    
        'acknowledge that we're calling them
        
        sendtouser usr, "CAL " & PID & " RINGING " & sessionid
        
    Case "ANS"
        '"answering" a "call" (an added user is joining)
        If CKIcookieisvalid(cl(3)) Then
            ws(Index).Tag = cl(2)
            usr = cl(2)
            Dim col As New Collection
            
            For a = 1 To ws.UBound
                If ws(a).Tag <> "" And ws(a).Tag <> "NEWBIE" And ws(a).State = 7 Then
                    col.Add ws(a).Tag
                End If
            Next a
            
            For a = 1 To col.Count
                sendtouser usr, "IRO " & PID & " " & a & " " & col.Count & " " & col(a) & " " & getscreenname(col(a))
            Next a
            
            sendtouser usr, "ANS " & PID & " OK"
            
            sendtoall "JOI " & usr & " " & getscreenname(usr)
        Else
            sendtouser usr, "200 " & PID
        End If
    
    Case "OUT"
        'they want out, give them out
        ws(Index).Close
    
    Case "MSG"
        'they are sending somehing
        'this could be anything from text, to the details of how to transfer
        'files.
        z = Mid(z, InStr(1, z, vbCrLf) + 2)
        sendtoallbut usr, "MSG " & usr & " " & getscreenname(usr) & " " & Len(z) & vbCrLf & z
        sendtouser usr, "ACK " & PID
        z = Mid(z, InStr(1, z, vbCrLf & vbCrLf) + 4)
        z = Replace(z, vbCrLf, "<nl>")
        If z <> "" Then Text1 = Text1 & Mid(usr, 1, 5) & ": " & z & vbCrLf
        If z <> "" Then ts.WriteLine Mid(usr, 1, 5) & ": " & z
        
        Text1.SelStart = Len(Text1) + 1
    End Select
End Sub

Private Sub ws_SendComplete(Index As Integer)
    senddone(Index) = True
End Sub
