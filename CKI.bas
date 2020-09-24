Attribute VB_Name = "CKI"
Private cookies As New Collection

Public Function CKInewcookie() As String
    Dim C As String
    Randomize Timer
    C = Format(Int(Rnd() * 100000000#), "0") & "." & Format(Int(Rnd() * 100000000#), "0")
    cookies.Add CStr(C)
    CKInewcookie = C
End Function

Public Function CKIcookieisvalid(cookie) As Boolean
    For a = 1 To cookies.Count
        If cookies(a) = CStr(cookie) Then CKIcookieisvalid = True: cookies.Remove a: Exit Function
    Next
End Function
