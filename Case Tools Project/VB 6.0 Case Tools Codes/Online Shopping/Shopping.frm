VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Shopping 
   Caption         =   "Shopping"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15840
   LinkTopic       =   "Form2"
   ScaleHeight     =   9315
   ScaleWidth      =   15840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   27
      Top             =   8400
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc shoppingado 
      Height          =   855
      Left            =   11760
      Top             =   8040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Shopping.frx":0000
      OLEDBString     =   $"Shopping.frx":0087
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "shopping"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   25
      Top             =   8400
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   24
      Top             =   8400
      Width           =   2535
   End
   Begin VB.TextBox txtcode 
      DataField       =   "Code word"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      TabIndex        =   23
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox txtdob 
      DataField       =   "Date of Birth"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   22
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtgender 
      DataField       =   "Gender"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   21
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtadd 
      DataField       =   "Address"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   20
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox txtphone 
      DataField       =   "Contact No"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   19
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox txtproduct 
      DataField       =   "Product Name"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   18
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtprice 
      DataField       =   "Price"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   17
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtquantity 
      DataField       =   "Quantity"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   16
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtmail 
      DataField       =   "Mail Id"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   15
      Top             =   7200
      Width           =   2415
   End
   Begin VB.TextBox txtamount 
      DataField       =   "Total Amount"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   14
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox txtpayment 
      DataField       =   "Payment"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      TabIndex        =   13
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox txtname 
      DataField       =   "Name"
      DataSource      =   "shoppingado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   12
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "Online Shopping"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   26
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label12 
      Caption         =   "Code Word"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   11
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Date of Birth "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Contact No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Mail Id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   7320
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Shopping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
shoppingado.Recordset.Fields("name") = txtname.Text
shoppingado.Recordset.Fields("date of birth") = txtdob.Text
shoppingado.Recordset.Fields("gender") = txtgender.Text
shoppingado.Recordset.Fields("Address") = txtadd.Text
shoppingado.Recordset.Fields("Contact no") = txtphone.Text
shoppingado.Recordset.Fields("mail id") = txtmail.Text
shoppingado.Recordset.Fields("product name") = txtproduct.Text
shoppingado.Recordset.Fields("price") = txtprice.Text
shoppingado.Recordset.Fields("quantity") = txtquantity.Text
shoppingado.Recordset.Fields("total amount") = txtamount.Text
shoppingado.Recordset.Fields("payment") = txtpayment.Text
shoppingado.Recordset.Fields("code word") = txtcode.Text
shoppingado.Recordset.Update
MsgBox "User Shopping Successful"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
txtname.Text = ""
txtdob.Text = ""
txtgender.Text = ""
txtadd.Text = ""
txtphone.Text = ""
txtmail.Text = ""
txtproduct.Text = ""
txtprice.Text = ""
txtquantity.Text = ""
txtamount.Text = ""
txtpayment.Text = ""
txtcode.Text = ""
End Sub

Private Sub Form_Load()
shoppingado.Recordset.AddNew
End Sub
