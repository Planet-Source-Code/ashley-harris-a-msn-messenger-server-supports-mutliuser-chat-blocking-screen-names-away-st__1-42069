Attribute VB_Name = "userinfo"
Public Function userexists(usr) As Boolean
    On Error GoTo nope
    finduser CStr(usr)
    If users("ID") > 0 Then userexists = True
nope:
End Function

Public Sub finduser(emailaddr As String)
    users.MoveFirst
    On Error GoTo out
    While users.PercentPosition < 100 And LCase(users("Email Address")) <> LCase(emailaddr)
        users.MoveNext
    Wend
out:
End Sub

Public Function getscreenname(emailaddr) As String
    finduser LCase(CStr(emailaddr))
    getscreenname = users("Friendlyname")
End Function

Public Function getpassword(emailaddr) As String
    On Error Resume Next
    finduser LCase(CStr(emailaddr))
    getpassword = users("Password")
End Function

Public Sub setscreename(emailaddr, newname)
    finduser LCase(CStr(emailaddr))
    users.Edit
    users("Friendlyname") = newname
    users.Update
End Sub
