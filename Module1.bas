Attribute VB_Name = "ContactLists"
Public fl As Recordset
Public rl As Recordset
Public al As Recordset
Public bl As Recordset

Public users As Recordset

Public db As Database



Public Sub initlists()
    
    dao.CompactDatabase "database.mdb", "database2.mdb"
    fso.CopyFile "database2.mdb", "database.mdb", True
    fso.DeleteFile "database2.mdb"
    
    Set db = dao.Workspaces(0).OpenDatabase("database.mdb")
    'Set fl = db.QueryDefs("FWlist").OpenRecordset
    'Set rl = db.QueryDefs("RLlist").OpenRecordset
    'Set al = db.QueryDefs("ALlist").OpenRecordset
    'Set bl = db.QueryDefs("BLlist").OpenRecordset
    
    Set users = db.TableDefs("users").OpenRecordset
    
    'For Each t In Array(fl, rl, al, bl, users)
    '    t.MoveLast
    '    t.MoveFirst
    'Next
    
End Sub

Public Sub refreshlists()
    On Error Resume Next
    For Each t In Array(fl, rl, al, bl, users)
        t.Requery
    Next
End Sub

Public Sub getlists(emailaddr)
    'Dim q As QueryDef
    For Each q In Array(db.QueryDefs("FWlist"), db.QueryDefs("RLlist"), db.QueryDefs("ALlist"), db.QueryDefs("BLlist"))
        q.Parameters.Refresh
        q.Parameters("User") = CStr(emailaddr)
    Next
    
    Set fl = db.QueryDefs("FWlist").OpenRecordset
    Set rl = db.QueryDefs("RLlist").OpenRecordset
    Set al = db.QueryDefs("ALlist").OpenRecordset
    Set bl = db.QueryDefs("BLlist").OpenRecordset
    
    refreshlists
End Sub

Public Function pulllist(l As String, email As String) As Dictionary
    Dim a As New Dictionary, list As Recordset
    
    getlists email

    Select Case l
    Case "al"
        Set list = al
    Case "bl"
        Set list = bl
    Case "rl"
        Set list = rl
    Case "fl"
        Set list = fl
    End Select

    On Error GoTo out
    list.MoveLast
    list.MoveFirst
    
    On Error GoTo out
    While list.PercentPosition < 100
        a(list("ToAddr").value) = list("ToNam").value
        list.MoveNext
    Wend
out:
    Set pulllist = a
    
End Function

Public Sub removeuserfromlist(fromaddr, list, addr)
    On Error GoTo out
    finduser CStr(fromaddr)
    fromid = users("ID")
    
    finduser CStr(addr)
    toid = users("ID")
    
    Dim con As Recordset
    Set con = db.TableDefs("contact").OpenRecordset
    On Error GoTo out
    con.MoveLast
    con.MoveFirst
    While con.PercentPosition < 100
        If con("fid") = fromid And con("tid") = toid And LCase(con("type")) = LCase(list) Then
            'con.Edit
            con.Delete 'htf does delete work?, if you ask me, it doesnt!
            'con.Update
            GoTo out
        End If
        con.MoveNext
    Wend
out:
    refreshlists
End Sub

Public Sub addusertolist(fromaddr, list, addr)
    finduser LCase(CStr(fromaddr))
    fromid = users("ID")
    
    finduser LCase(CStr(addr))
    toid = users("ID")

    Dim con As Recordset
    Set con = db.TableDefs("contact").OpenRecordset
    con.AddNew
    con("fid") = fromid
    con("tid") = toid
    con("type") = list
    con.Update
End Sub


