VERSION 5.00
Begin VB.Form change 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2085
   ClientLeft      =   6285
   ClientTop       =   6975
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5610
   Begin VB.Frame Frame3 
      Caption         =   "VERIFY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1275
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.Frame Frame2 
         Caption         =   "NEW PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1275
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   5055
         Begin VB.Frame Frame1 
            Caption         =   "OLD PASSWORD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1275
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   5055
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1440
               TabIndex        =   18
               Top             =   300
               Width           =   2235
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1440
               PasswordChar    =   "*"
               TabIndex        =   17
               Top             =   780
               Width           =   2235
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "USER NAME :"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "PASSWORD :"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   21
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label Label2 
               BackColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   1500
               TabIndex        =   20
               Top             =   840
               Width           =   2235
            End
            Begin VB.Label Label2 
               BackColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   1500
               TabIndex        =   19
               Top             =   360
               Width           =   2235
            End
            Begin VB.Image Image1 
               Height          =   540
               Index           =   0
               Left            =   4080
               Picture         =   "change.frx":0000
               Stretch         =   -1  'True
               Top             =   420
               Width           =   645
            End
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   300
            Width           =   2235
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1440
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   780
            Width           =   2235
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PASSWORD :"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "USER NAME :"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   13
            Top             =   360
            Width           =   2235
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   1500
            TabIndex        =   12
            Top             =   840
            Width           =   2235
         End
         Begin VB.Image img 
            Height          =   540
            Index           =   1
            Left            =   4080
            Picture         =   "change.frx":0442
            Stretch         =   -1  'True
            Top             =   420
            Width           =   645
         End
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         DataField       =   "password"
         DataSource      =   "Data1"
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   780
         Width           =   2235
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         DataField       =   "user"
         DataSource      =   "Data1"
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME :"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   1500
         TabIndex        =   6
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   1500
         TabIndex        =   5
         Top             =   840
         Width           =   2235
      End
      Begin VB.Image img 
         Height          =   540
         Index           =   2
         Left            =   4080
         Picture         =   "change.frx":0884
         Stretch         =   -1  'True
         Top             =   420
         Width           =   645
      End
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   795
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
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
a = 0
Data1.Visible = False
Data1.DatabaseName = App.Path + "\Pas.mdb"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).BackColor = &H8000&
End Sub

Private Sub Label2_Click(Index As Integer)
On Error Resume Next
If Frame3.Visible = True Or Frame2.Visible = True Then
Data1.Recordset.CancelUpdate
Frame1.Visible = True
Frame2.Visible = True
Text1.Text = ""
Text2.Text = ""
change.Height = 2340
Unload change
End If
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).BackColor = vbRed
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If LCase(Text1.Text) = "verify" Then
change.Height = 4890
Else
Text2.SetFocus
SendKeys "{home}+{end}"
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
If LCase(Text1.Text) = "add" And LCase(Text2.Text) = "add" Then
Frame1.Visible = False
Text3.SetFocus
Else
Data1.RecordSource = ("select * from Pass where password = '" & Text2.Text & "' and user = '" & Text1.Text & "'")
Data1.Refresh
If Text5.Text = "" And Text6.Text = "" Then
MsgBox "Invalid Username and Password!", vbCritical, "Error"
Text1.SetFocus
SendKeys "{home}+{end}"
Text7.Text = Val(Text7.Text) + 1

'for database will refresh
Data1.RecordSource = ("select * from Pass order by user")
Data1.Refresh

'counter for failed
If Text7.Text = 2 Then
Text1.Enabled = False
Text2.Enabled = False
Text1.SetFocus
MsgBox "It seems you forgot your password... I'll give you another chance, but if you failed I will be forced to terminate this system!", vbInformation, "Remarks"
End If

If Text7.Text = 3 Then
MsgBox "System Shut Down", vbCritical
End
End If

Else
Frame1.Visible = False
Text3.SetFocus
End If
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
LCase (Text3.Text)
Text4.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim a
If KeyAscii = 13 Then

Data1.RecordSource = ("select * from Pass where  user = '" & Text3.Text & "'")
Data1.Refresh

If Text5.Text = "" Then

If Text1.Text = "add" Then
LCase (Text4.Text)
Frame2.Visible = False
Data1.Recordset.AddNew
Text5.SetFocus
Else
jon:
Data1.RecordSource = ("select * from Pass where password='" & Text2.Text & "' and user= '" & Text1.Text & "'")
Data1.Refresh
LCase (Text4.Text)
Data1.Recordset.Edit
Frame2.Visible = False
Text5.SetFocus
SendKeys "{home}+{end}"
End If

Else

If LCase(Text1.Text) = "add" Then
MsgBox "User Name Duplicated!", vbInformation, "Duplicated"
Text3.Text = ""
Text4.Text = ""
Text3.SetFocus
Else
a = MsgBox("User Name Duplicated! Change the Password?", vbInformation + vbYesNo, "Duplicated")
If a = vbYes Then
If Text1.Text = Text3.Text Then
GoTo jon
Else
MsgBox "Warning! this is not your Existing Login Name", vbCritical, "Warning"
End If
End If
Text3.Text = ""
Text4.Text = ""
Text3.SetFocus
Data1.RecordSource = ("select * from Pass where password='" & Text2.Text & "'")
Data1.Refresh
End If
End If
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
SendKeys "{home}+{end}"
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If LCase(Text5.Text) = Text3.Text And LCase(Text6.Text) = Text4.Text Then
Data1.Recordset.Update
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Frame1.Visible = True
Frame2.Visible = True
Text1.SetFocus
Data1.RecordSource = ("select* from pass order by user")
Unload Me
Else
Data1.Recordset.CancelUpdate
MsgBox "Not Valid!, Check your Username and Password", vbCritical
Frame2.Visible = True
Text3.SetFocus
Exit Sub
End If
End If
End Sub



