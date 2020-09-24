VERSION 5.00
Begin VB.Form admin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Database"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox scname 
      Height          =   330
      Left            =   2925
      TabIndex        =   6
      Top             =   1170
      Width           =   1680
   End
   Begin VB.TextBox pw 
      Height          =   345
      Left            =   2880
      TabIndex        =   5
      Top             =   390
      Width           =   1710
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   360
      Left            =   1410
      TabIndex        =   2
      Top             =   2190
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   2205
      Width           =   675
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   2520
   End
   Begin VB.Label Label2 
      Caption         =   "Screenname:"
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   180
      Width           =   1785
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    users.AddNew
    users("Email Address") = InputBox("What is their screenname? ie email address", "Ashleys MSN server")
    users("Password") = "password"
    users("FriendlyName") = "Newbie"
    users.Update
    updateusernamelist
    
End Sub

Private Sub Command2_Click()
    users.Delete
    updateusernamelist
End Sub

Private Sub Form_Load()
    updateusernamelist
End Sub

Public Sub updateusernamelist()
    List1.Clear
    On Error Resume Next
    users.MoveFirst
    Err.Clear
    While Err.Number = 0
        List1.AddItem users("Email Address")
        users.MoveNext
    Wend
End Sub

Private Sub List1_Click()
    pw = getpassword(List1.list(List1.ListIndex))
    scname = getscreenname(List1.list(List1.ListIndex))
    finduser List1.list(List1.ListIndex)
End Sub

Private Sub pw_Change()
    If pw = "" Then Exit Sub
    users.Edit
    users("Password") = pw
    users.Update
End Sub

Private Sub scname_Change()
    If scname = "" Then Exit Sub
    users.Edit
    users("Friendlyname") = scname
    users.Update
End Sub
