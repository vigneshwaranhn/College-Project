VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form books 
   Caption         =   "books"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   ScaleHeight     =   8805
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   11040
      TabIndex        =   14
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txtauthor 
      DataField       =   "authorname"
      DataSource      =   "booksado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   13
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox txtlendingbooks 
      DataField       =   "lendingbooks"
      DataSource      =   "booksado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   12
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox txtreturnbooks 
      DataField       =   "returnbooks"
      DataSource      =   "booksado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   11
      Top             =   8040
      Width           =   2655
   End
   Begin VB.TextBox txtbook 
      DataField       =   "bookname"
      DataSource      =   "booksado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   10
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtname 
      DataField       =   "name"
      DataSource      =   "booksado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtreg 
      DataField       =   "regno"
      DataSource      =   "booksado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc booksado 
      Height          =   735
      Left            =   11160
      Top             =   7800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
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
      Connect         =   $"books.frx":0000
      OLEDBString     =   $"books.frx":0087
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "books"
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
   Begin VB.CommandButton exitbtn 
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
      Height          =   855
      Left            =   11040
      TabIndex        =   1
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton conbtn 
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      TabIndex        =   0
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Library Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   15
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label6 
      Caption         =   "Book Name"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Author Name"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Date of Lending Books"
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
      Left            =   1560
      TabIndex        =   7
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Date of Return Books"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Label Label2 
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
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Reg no"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
txtreg.Text = ""
txtname.Text = ""
txtbook.Text = ""
txtauthor.Text = ""
txtlendingbooks.Text = ""
txtreturnbooks.Text = ""

End Sub

Private Sub conbtn_Click()
booksado.Recordset.Fields("regno") = txtreg.Text
booksado.Recordset.Fields("name") = txtname.Text
booksado.Recordset.Fields("bookname") = txtbook.Text
booksado.Recordset.Fields("authorname") = txtauthor.Text
booksado.Recordset.Fields("lendingbooks") = txtlendingbooks.Text
booksado.Recordset.Fields("returnbooks") = txtreturnbooks.Text
booksado.Recordset.Update
MsgBox "User books Successful"
End Sub
Private Sub exitbtn_Click()
End
End Sub

Private Sub Form_Load()
booksado.Recordset.AddNew
End Sub

