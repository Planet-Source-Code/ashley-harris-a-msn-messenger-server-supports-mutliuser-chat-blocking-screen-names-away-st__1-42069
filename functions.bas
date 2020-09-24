Attribute VB_Name = "functions"
'yep, every server I've released so far on PSC contains this function, well spotted

Public fso As New FileSystemObject

Public Function fromhttpstringtostring(httpstring As String) As String
    'turns 'This%20is%20cool' into 'This is cool'
    httpstring = Replace(httpstring, "+", " ")
    While InStr(1, httpstring, "%")
        fromhttpstringtostring = fromhttpstringtostring & Mid(httpstring, 1, InStr(1, httpstring, "%") - 1)
        httpstring = Mid(httpstring, InStr(1, httpstring, "%"))
        esc = Mid(httpstring, 1, 3)
        ch = Chr(hexdiget(Mid(esc, 2, 1)) * 16 + hexdiget(Mid(esc, 3, 1)))
        httpstring = Replace(httpstring, esc, ch)
    Wend
    fromhttpstringtostring = fromhttpstringtostring & httpstring
End Function

Public Function hexdiget(d) As Integer
    'converts a number from 0-15 into a hexeqiverlant (ie a=10)
    If d = Val(CStr(d)) Then hexdiget = d: Exit Function
    Select Case LCase(d)
    Case "a"
        hexdiget = 10
    Case "b"
        hexdiget = 11
    Case "c"
        hexdiget = 12
    Case "d"
        hexdiget = 13
    Case "e"
        hexdiget = 14
    Case "f"
        hexdiget = 15
    End Select
End Function

Public Function parseheaders(h As String) As Dictionary
    'turn
    'cookie: name=ashley
    'referer: www.pornrus.com
    'accept: all the stuff that goes here.
    'langauge: en-au
    'etc.
    'into a datadictonary.
    
    'the msn 'msg' command (after the first newline) follows this
    'protocol for headers. (same as email, http, etc.)
    Dim K As String, v As String
    Set p = New Dictionary
    p.CompareMode = TextCompare
    h = h & vbNewLine
    h = Replace(h, ": ", ":")
    h = Replace(h, vbCrLf & " ", "")
    h = Replace(h, vbCrLf & vbTab, "")
    While h <> vbNewLine And h <> "" And h <> vbNewLine & vbNewLine
        K = LCase(Mid(h, 1, InStr(1, h, ":") - 1))
        h = Mid(h, Len(K) + 2)
        v = Mid(h, 1, InStr(1, h, vbNewLine) - 1)

        h = Mid(h, Len(v) + 3)
        p(K) = v
    Wend
    Set parseheaders = p
End Function

