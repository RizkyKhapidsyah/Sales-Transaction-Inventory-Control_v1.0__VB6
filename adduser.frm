VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form useradd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   3855
   ClientLeft      =   5460
   ClientTop       =   6510
   ClientWidth     =   6675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6675
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   5955
      Begin VB.CommandButton cmdok 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   2400
         Width           =   1155
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   240
         Top             =   2280
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Jawad\pas.mdb"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Jawad\pas.mdb"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "adduser.frx":0000
         Left            =   2880
         List            =   "adduser.frx":000A
         TabIndex        =   0
         Text            =   "Select Rights Level"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text3 
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
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   2400
         Width           =   1155
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
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
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
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "CONFIRM PASSWORD : "
         Height          =   375
         Index           =   5
         Left            =   840
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   2940
         TabIndex        =   11
         Top             =   1980
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   420
         Picture         =   "adduser.frx":0014
         Stretch         =   -1  'True
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "USER NAME :"
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD : "
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   2940
         TabIndex        =   8
         Top             =   1020
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   2940
         TabIndex        =   7
         Top             =   1500
         Width           =   2775
      End
   End
End
Attribute VB_Name = "useradd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Sub cmdcancel_Click()
    Unload useradd
    
End Sub

Private Sub Cmdok_Click()
    Unload useradd
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
        SendKeys "{home}+{end}" 'Selects Text within textbox
        Text1.Text = LCase(Text1.Text)
    End If

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
        SendKeys "{home}+{end}"
        Text2.Text = LCase(Text2.Text)
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text2.Text = Text3.Text Then
            sql2 = "insert into pass(user, password, rights) values ('" & Text1.Text & "','" & Text2.Text & "','" & Combo1.Text & "')"
            con.Open ("Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Jawad\pas.mdb")
            Set rst = con.Execute(sql2)
        End If
        MsgBox "New User Added"
    End If
    
End Sub
