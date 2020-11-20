VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Login"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pass"
      Top             =   120
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Beep
    End
End Sub
Private Sub Form_Load()
    Data1.Visible = False
    Data1.DatabaseName = App.Path + "\pas.mdb"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
        SendKeys "{home}+{end}"
        Text1.Text = LCase(Text1.Text)
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Data1.RecordSource = ("select * from pass where password='" & _
        Text2.Text & "' and user = '" & Text1.Text & "'")
        Data1.Refresh

    If Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "Access Denied", vbCritical, "Password"
        Text1.SetFocus
        SendKeys "{home}+{end}"
    Else
        MsgBox "Access Granted", vbInformation, "Password"
        Unload Password
      
End If
End If

End Sub


