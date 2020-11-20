VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Sales Transaction and Inventory Control (Work Flow)"
   ClientHeight    =   8325
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11325
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stabar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7830
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIForm1.frx":0000
            Object.ToolTipText     =   "USER"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   11906
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "4:41 PM"
            Object.ToolTipText     =   "TIME"
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Index           =   0
      WindowList      =   -1  'True
      Begin VB.Menu login 
         Caption         =   "L&ogin"
         Index           =   1
      End
      Begin VB.Menu adduser 
         Caption         =   "&Add New User"
      End
      Begin VB.Menu changep 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Index           =   3
      End
   End
   Begin VB.Menu menuitems 
      Caption         =   "&Menu Items"
      HelpContextID   =   1
      Index           =   1
      Begin VB.Menu menudetails 
         Caption         =   "Menu &Details"
      End
      Begin VB.Menu category 
         Caption         =   "&Food Categories"
      End
   End
   Begin VB.Menu sales 
      Caption         =   "&Sales"
   End
   Begin VB.Menu purchase 
      Caption         =   "&Purchase Orders"
      Index           =   2
   End
   Begin VB.Menu supplier 
      Caption         =   "&Suppliers"
      Begin VB.Menu Supplerinfo 
         Caption         =   "S&upplier Info"
      End
      Begin VB.Menu shipmethod 
         Caption         =   "S&hipping Methods"
      End
   End
   Begin VB.Menu employee 
      Caption         =   "&Employees"
   End
   Begin VB.Menu rpts 
      Caption         =   "&Reports"
      Begin VB.Menu rptfoodmenu 
         Caption         =   "&Food Menu details"
      End
      Begin VB.Menu rptsales 
         Caption         =   "&Sales"
      End
      Begin VB.Menu rptpurchaseorder 
         Caption         =   "&Purchase Orders"
      End
      Begin VB.Menu rptsuppliers 
         Caption         =   "S&uppliers"
      End
   End
   Begin VB.Menu about 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub adduser_Click()
    Load useradd
    useradd.Show
End Sub

Private Sub category_Click()
    Load ffoodcategory
    ffoodcategory.Show
End Sub

Private Sub changep_Click()
    Load change
    change.Show
End Sub

Private Sub employee_Click()
    Load femployee
    femployee.Show
End Sub

Private Sub exit_Click(Index As Integer)
    Dim a
    a = MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit")
    If a = vbYes Then
        Beep
        End
    End If
End Sub

Private Sub login_Click(Index As Integer)
    Unload MDIForm1
    Load frmlogin
    frmlogin.Show
End Sub


Private Sub menudetails_Click()
    Load fmenudetail
    fmenudetail.Show
End Sub

Private Sub purchase_Click(Index As Integer)
    Load fpurchaseOrder
    fpurchaseOrder.Show
End Sub

Private Sub rptfoodmenu_Click()
    datarptmenuitems.Show
End Sub

Private Sub rptpurchaseorder_Click()
    datarptpurchaseorder.Show
End Sub

Private Sub rptsales_Click()
    datarptsales.Show
End Sub

Private Sub rptsuppliers_Click()
    datarptsuppliers.Show
End Sub

Private Sub sales_Click()
    Load fsales
    fsales.Show
End Sub

Private Sub shipmethod_Click()
    Load fshipmethod
    fshipmethod.Show
End Sub

Private Sub Supplerinfo_Click()
    Load fsupplier
    fsupplier.Show
End Sub

Private Sub vuser_Click()
    Load fusersview
    fusersview.Show
End Sub
