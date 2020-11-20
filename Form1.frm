VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   4950
   ClientTop       =   5685
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "rights"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   5955
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   2820
         TabIndex        =   9
         Top             =   780
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   2820
         TabIndex        =   8
         Top             =   300
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD : "
         Height          =   315
         Index           =   1
         Left            =   1620
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "USER NAME :"
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   6
         Top             =   300
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   420
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "pass"
      Top             =   120
      Width           =   2355
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "user"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "password"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "This is for security system wall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Pls. Supply the following Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   300
      Width           =   6315
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag2 As Integer

Private Sub Command1_Click()
Beep
End
End Sub
Private Sub Form_Load()
Data1.Visible = False
Data1.DatabaseName = App.Path + "\Pas.mdb"
End Sub

Private Sub Label4_Click()
Beep
End
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
Data1.RecordSource = ("select * from pass where password='" & Text2.Text & "' and user = '" & Text1.Text & "'")
Data1.Refresh
If Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Access Denied", vbCritical, "Password"
Text1.SetFocus
SendKeys "{home}+{end}"

Else
flag2 = Text5.Text
MsgBox "Access Granted", vbInformation, "Password"

MDIForm1.stabar.Panels(1).Text = "User :  " & UCase(Text1.Text)

Unload frmlogin
Load MDIForm1
MDIForm1.Show

End If
End If
End Sub
